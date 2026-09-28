"""Save filtered Outlook messages to SQL Server, then move them to Processed."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence

import pyodbc
import yaml


PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
	sys.path.insert(0, str(PROJECT_ROOT))

from shared.logging import setup_logging
from shared.sharepoint.graph import GraphClient
from shared.sharepoint.outlook import OutlookClient


CONFIG_PATH = PROJECT_ROOT / "config" / "config.yaml"
AUTH_PATH = PROJECT_ROOT / "config" / "auth.yaml"
DEFAULT_SCHEMA = "raw"
DEFAULT_TABLE = "watchoffice_email"
DEFAULT_TOP = 5
RUN_SCRIPT_SETTINGS = True

SCRIPT_SETTINGS: dict[str, Any] = {
	"folder": "Inbox/Info Sources/Watch Office",
	"processed_folder": "Inbox/Info Sources/Watch Office/Reviewed",
	"sender_contains": None,
	"subject_contains": None,
	"schema": DEFAULT_SCHEMA,
	"table": DEFAULT_TABLE,
	"top": DEFAULT_TOP,
	"create_table": False,
	"preview_only": False,
	"preview_character_limit": 2000,
}

logger = setup_logging("email_to_safari")


@dataclass(frozen=True)
class ProcessResult:
	saved_count: int
	moved_count: int


def _merge_dicts(base: dict[str, Any], override: Mapping[str, Any]) -> dict[str, Any]:
	merged = dict(base)
	for key, value in override.items():
		if isinstance(value, Mapping) and isinstance(merged.get(key), Mapping):
			merged[key] = _merge_dicts(dict(merged[key]), value)
		else:
			merged[key] = value
	return merged


def load_config(
	config_path: Path = CONFIG_PATH,
	auth_path: Path = AUTH_PATH,
) -> dict[str, Any]:
	config: dict[str, Any] = {}
	for path in (config_path, auth_path):
		if path.exists():
			with path.open("r", encoding="utf-8") as config_file:
				config = _merge_dicts(config, yaml.safe_load(config_file) or {})

	sql_config = config.setdefault("sqlserver", {})
	sql_config.setdefault("username", os.getenv("SQLSERVER_USERNAME"))
	sql_config.setdefault("password", os.getenv("SQLSERVER_PASSWORD"))
	sql_config.setdefault("database", "PROD_SAFARI")
	return config


def validate_identifier(identifier: str, label: str) -> str:
	if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", identifier):
		raise ValueError(f"Invalid {label}: {identifier!r}")
	return identifier


def build_connection_string(sql_config: Mapping[str, Any]) -> str:
	required = ("host", "database", "driver")
	missing = [name for name in required if not sql_config.get(name)]
	if missing:
		raise ValueError("Missing SQL Server settings: " + ", ".join(missing))

	server = str(sql_config["host"])
	if sql_config.get("port"):
		server = f"{server},{sql_config['port']}"

	parts = [
		f"DRIVER={{{sql_config['driver']}}}",
		f"SERVER={server}",
		f"DATABASE={sql_config['database']}",
	]
	if sql_config.get("trusted_connection", False):
		parts.append("Trusted_Connection=yes")
	else:
		username = sql_config.get("username")
		password = sql_config.get("password")
		if not username or password is None:
			raise ValueError(
				"SQL Server username and password are required. Set them in "
				"config/auth.yaml or SQLSERVER_USERNAME/SQLSERVER_PASSWORD."
			)
		parts.extend((f"UID={username}", f"PWD={password}"))

	parts.extend(("Encrypt=yes", "TrustServerCertificate=yes"))
	return ";".join(parts) + ";"


def message_body(message: Mapping[str, Any]) -> str:
	body = message.get("body") or {}
	return body.get("content") or message.get("bodyPreview") or ""


def message_json(message: Mapping[str, Any]) -> str:
	return json.dumps(message, ensure_ascii=False, default=str, indent=2)


def parse_graph_datetime(value: Any) -> datetime | None:
	if not value:
		return None

	if isinstance(value, datetime):
		parsed_datetime = value
	else:
		parsed_datetime = datetime.fromisoformat(str(value).replace("Z", "+00:00"))

	if parsed_datetime.tzinfo is None:
		return parsed_datetime

	return parsed_datetime.astimezone(timezone.utc).replace(tzinfo=None)


def message_to_record(message: Mapping[str, Any]) -> tuple[Any, ...]:
	if not message.get("id"):
		raise ValueError("Outlook message does not contain a Graph id.")

	sender = (message.get("from") or {}).get("emailAddress") or {}
	body = message.get("body") or {}
	fields = (
		parse_graph_datetime(message.get("receivedDateTime")),
		parse_graph_datetime(message.get("sentDateTime")),
		message.get("subject"),
		sender.get("name"),
		sender.get("address"),
		body.get("contentType"),
		body.get("content"),
	)
	for label, value, limit in zip(
		("subject", "sender_name", "sender_email", "body_content_type"),
		fields[2:6],
		(100, 50, 50, 50),
	):
		if value is not None and len(value) > limit:
			raise ValueError(f"{label} exceeds NVARCHAR({limit}) for message {message['id']}")
	return fields


def message_matches_filters(
	message: Mapping[str, Any],
	sender_contains: str | None = None,
	subject_contains: str | None = None,
) -> bool:
	sender = (
		message.get("from", {})
		.get("emailAddress", {})
		.get("address", "")
	)
	subject = message.get("subject", "")

	if sender_contains and sender_contains.casefold() not in sender.casefold():
		return False

	if subject_contains and subject_contains.casefold() not in subject.casefold():
		return False

	return True


def received_datetime_sort_key(message: Mapping[str, Any]) -> tuple[bool, str]:
	received_datetime = message.get("receivedDateTime") or message.get("sentDateTime")
	return received_datetime is None, str(received_datetime or "")


def get_filtered_messages(
	outlook: OutlookClient,
	folder: str,
	sender_contains: str | None = None,
	subject_contains: str | None = None,
	top: int = DEFAULT_TOP,
) -> list[Mapping[str, Any]]:
	messages = outlook.get_folder_messages(folder_name=folder, top=top)
	filtered_messages = [
		message
		for message in messages
		if message_matches_filters(
			message,
			sender_contains=sender_contains,
			subject_contains=subject_contains,
		)
	]
	return sorted(filtered_messages, key=received_datetime_sort_key)


def ensure_table(cursor: pyodbc.Cursor, schema: str, table: str) -> None:
	qualified_table = f"[{schema}].[{table}]"
	cursor.execute(
		f"""
IF OBJECT_ID(N'{qualified_table}', N'U') IS NULL
BEGIN
	CREATE TABLE {qualified_table} (
		[message_id] BIGINT IDENTITY(1,1) NOT NULL,
		[received_datetime_utc] DATETIME2(7) NULL,
		[sent_datetime_utc] DATETIME2(7) NULL,
		[subject] NVARCHAR(100) NULL,
		[sender_name] NVARCHAR(50) NULL,
		[sender_email] NVARCHAR(50) NULL,
		[body_content_type] NVARCHAR(50) NULL,
		[body_content] NVARCHAR(MAX) NULL
	);
END;
"""
	)


def save_messages_to_sql(
	connection: pyodbc.Connection,
	messages: Sequence[Mapping[str, Any]],
	schema: str,
	table: str,
	create_table: bool = True,
) -> int:
	schema = validate_identifier(schema, "schema")
	table = validate_identifier(table, "table")
	cursor = connection.cursor()
	if create_table:
		ensure_table(cursor, schema, table)

	sql = f"""
	INSERT INTO [{schema}].[{table}] (
		[received_datetime_utc],
		[sent_datetime_utc],
		[subject],
		[sender_name],
		[sender_email],
		[body_content_type],
		[body_content]
	)
	OUTPUT INSERTED.[message_id]
	VALUES (?, ?, ?, ?, ?, ?, ?);
"""

	saved_count = 0
	seen_message_ids: set[str] = set()
	try:
		cursor.execute(
			"SELECT COLUMNPROPERTY(OBJECT_ID(?), 'message_id', 'IsIdentity')",
			f"[{schema}].[{table}]",
		)
		if cursor.fetchone()[0] != 1:
			raise ValueError(f"[{schema}].[{table}].[message_id] must be an IDENTITY column")
		for message in messages:
			message_id = str(message.get("id") or "")
			if not message_id:
				raise ValueError("Outlook message does not contain a Graph id.")
			if message_id in seen_message_ids:
				raise ValueError(f"Duplicate Graph message id in folder results: {message_id}")
			seen_message_ids.add(message_id)
			record = message_to_record(message)
			cursor.execute(sql, *record)
			inserted = cursor.fetchone()
			if inserted is None or inserted[0] is None:
				raise RuntimeError(f"Message was not saved to SQL: {message_id}")
			saved_count += 1
		connection.commit()
	except Exception:
		connection.rollback()
		raise
	finally:
		cursor.close()

	return saved_count


def ensure_mail_folder(outlook: OutlookClient, folder_path: str) -> str:
	parts = [part for part in folder_path.split("/") if part]
	if not parts:
		raise ValueError("Processed folder path must not be empty.")

	parent_id = None
	for part in parts:
		if parent_id is None:
			folders = outlook.graph.get_all("me/mailFolders")
		else:
			folders = outlook.graph.get_all(f"me/mailFolders/{parent_id}/childFolders")

		match = next(
			(
				folder
				for folder in folders
				if folder.get("displayName", "").casefold() == part.casefold()
			),
			None,
		)
		if match is None:
			payload = {"displayName": part}
			if parent_id is None:
				match = outlook.graph.post("me/mailFolders", payload)
			else:
				match = outlook.graph.post(
					f"me/mailFolders/{parent_id}/childFolders",
					payload,
				)

		parent_id = match["id"]

	return parent_id


def move_saved_messages(
	outlook: OutlookClient,
	messages: Sequence[Mapping[str, Any]],
	source_folder: str,
	processed_folder: str,
) -> int:
	ensure_mail_folder(outlook, processed_folder)
	moved_count = 0
	for message in messages:
		message_id = message.get("id")
		if not message_id:
			raise ValueError("Outlook message does not contain a Graph id.")
		outlook.move_message(str(message_id), processed_folder)
		moved_count += 1

	source_folder_id = outlook.get_folder_id(source_folder)
	remaining_ids = {
		str(message["id"])
		for message in outlook.get_folder_messages(source_folder)
	}
	for message in messages:
		message_id = str(message["id"])
		if message_id in remaining_ids:
			outlook.graph.delete(f"me/mailFolders/{source_folder_id}/messages/{message_id}")
	return moved_count


def process_inbox_to_sql(
	folder: str,
	schema: str = DEFAULT_SCHEMA,
	table: str = DEFAULT_TABLE,
	processed_folder: str | None = None,
	sender_contains: str | None = None,
	subject_contains: str | None = None,
	top: int = DEFAULT_TOP,
	create_table: bool = True,
) -> ProcessResult:
	"""Save filtered folder messages to SQL Server, then move saved messages."""
	processed_folder = processed_folder or f"{folder.rstrip('/')}/Processed"
	config = load_config()
	outlook = OutlookClient(GraphClient())

	logger.info("Reading Outlook folder: %s", folder)
	messages = get_filtered_messages(
		outlook=outlook,
		folder=folder,
		sender_contains=sender_contains,
		subject_contains=subject_contains,
		top=top,
	)
	logger.info("Found %s messages after filtering", len(messages))

	connection = pyodbc.connect(
		build_connection_string(config["sqlserver"]),
		timeout=30,
		autocommit=False,
	)
	try:
		saved_count = save_messages_to_sql(
			connection=connection,
			messages=messages,
			schema=schema,
			table=table,
			create_table=create_table,
		)
	finally:
		connection.close()

	moved_count = move_saved_messages(
		outlook=outlook,
		messages=messages,
		source_folder=folder,
		processed_folder=processed_folder,
	)
	logger.info(
		"Saved %s messages to [%s].[%s] and moved %s to %s",
		saved_count,
		schema,
		table,
		moved_count,
		processed_folder,
	)
	return ProcessResult(saved_count=saved_count, moved_count=moved_count)


def preview_inbox_messages(
	folder: str,
	sender_contains: str | None = None,
	subject_contains: str | None = None,
	top: int = 5,
	message_character_limit: int = 2000,
) -> list[Mapping[str, Any]]:
	"""Read a few messages and print what would be saved, without SQL or moves."""
	outlook = OutlookClient(GraphClient())
	messages = get_filtered_messages(
		outlook=outlook,
		folder=folder,
		sender_contains=sender_contains,
		subject_contains=subject_contains,
		top=top,
	)

	print(f"Messages found: {len(messages)}")
	for index, message in enumerate(messages, start=1):
		sender = (
			message.get("from", {})
			.get("emailAddress", {})
			.get("address", "")
		)
		message_text = message_json(message)
		if message_character_limit and len(message_text) > message_character_limit:
			message_text = message_text[:message_character_limit] + "\n... [message truncated]"

		print("=" * 100)
		print(f"Message {index}")
		print(f"MessageId         : {message.get('id', '')}")
		print(f"InternetMessageId : {message.get('internetMessageId', '')}")
		print(f"Sender            : {sender}")
		print(f"Subject           : {message.get('subject', '')}")
		print("Message JSON:")
		print(message_text)

	return messages


def run_from_script_settings(
	settings: Mapping[str, Any] = SCRIPT_SETTINGS,
) -> ProcessResult | list[Mapping[str, Any]]:
	folder = settings.get("folder")
	if not folder:
		raise ValueError("Set SCRIPT_SETTINGS['folder'] before running.")

	if settings.get("preview_only", False):
		return preview_inbox_messages(
			folder=str(folder),
			sender_contains=settings.get("sender_contains"),
			subject_contains=settings.get("subject_contains"),
			top=int(settings.get("top", DEFAULT_TOP)),
			message_character_limit=int(
				settings.get("preview_character_limit", 2000)
			),
		)

	return process_inbox_to_sql(
		folder=str(folder),
		schema=str(settings.get("schema", DEFAULT_SCHEMA)),
		table=str(settings.get("table", DEFAULT_TABLE)),
		processed_folder=settings.get("processed_folder"),
		sender_contains=settings.get("sender_contains"),
		subject_contains=settings.get("subject_contains"),
		top=int(settings.get("top", DEFAULT_TOP)),
		create_table=bool(settings.get("create_table", True)),
	)


def parse_args() -> argparse.Namespace:
	parser = argparse.ArgumentParser(
		description="Save Outlook messages to SQL Server and move them to Processed."
	)
	parser.add_argument("--folder", required=True, help='Example: "Inbox/Watch Office"')
	parser.add_argument(
		"--processed-folder",
		help='Defaults to "<folder>/Processed" when omitted.',
	)
	parser.add_argument("--sender-contains", help="Only process sender emails containing this text.")
	parser.add_argument("--subject-contains", help="Only process subjects containing this text.")
	parser.add_argument("--schema", default=DEFAULT_SCHEMA)
	parser.add_argument("--table", default=DEFAULT_TABLE)
	parser.add_argument("--top", type=int, default=DEFAULT_TOP)
	parser.add_argument(
		"--preview-only",
		action="store_true",
		help="Print matching messages without saving to SQL or moving them.",
	)
	parser.add_argument(
		"--preview-character-limit",
		type=int,
		default=2000,
		help="Maximum message JSON characters to print per email during preview.",
	)
	parser.add_argument(
		"--no-create-table",
		action="store_true",
		help="Do not create the SQL table if it does not already exist.",
	)
	return parser.parse_args()


def main() -> None:
	if RUN_SCRIPT_SETTINGS:
		run_from_script_settings()
		return

	args = parse_args()
	if args.preview_only:
		preview_inbox_messages(
			folder=args.folder,
			sender_contains=args.sender_contains,
			subject_contains=args.subject_contains,
			top=args.top,
			message_character_limit=args.preview_character_limit,
		)
		return

	process_inbox_to_sql(
		folder=args.folder,
		schema=args.schema,
		table=args.table,
		processed_folder=args.processed_folder,
		sender_contains=args.sender_contains,
		subject_contains=args.subject_contains,
		top=args.top,
		create_table=not args.no_create_table,
	)


if __name__ == "__main__":
	main()
