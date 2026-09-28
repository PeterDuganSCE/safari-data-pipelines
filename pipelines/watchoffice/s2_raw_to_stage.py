"""Extract Watch Office mail from raw into staging without changing raw rows."""

from __future__ import annotations

import sys
from pathlib import Path

import pyodbc

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
	sys.path.insert(0, str(PROJECT_ROOT))

from pipelines.watchoffice.s1_email_to_safari import build_connection_string, load_config
from pipelines.watchoffice.extract import UNKNOWN, extract_incident_data, isolate_incident_narrative


ENCRYPTED_BODY_TEXT = "Watch Office email is encrypted"


def extract_body_text(body_content: str | None) -> tuple[str, bool]:
	text = extract_incident_data(body_content or "")["incident_narrative"]
	if text == UNKNOWN or (
		"encrypt" in text.casefold() and not isolate_incident_narrative(text)
	):
		return ENCRYPTED_BODY_TEXT, True
	return text, False


def stage_watchoffice(connection: pyodbc.Connection) -> int:
	cursor = connection.cursor()
	try:
		cursor.execute("SELECT [message_id] FROM [staging].[watchoffice]")
		existing_ids = {row[0] for row in cursor.fetchall()}
		cursor.execute(
			"""
SELECT [message_id],
       CONVERT(smalldatetime,
           [received_datetime_utc] AT TIME ZONE 'UTC' AT TIME ZONE 'Pacific Standard Time'),
       [body_content]
FROM [raw].[watchoffice_email]
ORDER BY [message_id]
"""
		)
		rows = cursor.fetchall()
		seen_ids: set[int] = set()
		inserted_count = 0
		for message_id, received_datetime_pst, body_content in rows:
			if message_id is None or message_id in seen_ids:
				raise ValueError(f"Missing or duplicate raw message_id: {message_id}")
			seen_ids.add(message_id)
			if message_id in existing_ids:
				continue
			body_text, is_encrypted = extract_body_text(body_content)
			cursor.execute(
				"""
INSERT INTO [staging].[watchoffice]
    ([message_id], [received_datetime_pst], [body_text], [is_encrypted])
OUTPUT INSERTED.[message_id]
VALUES (?, ?, ?, ?)
""",
				message_id,
				received_datetime_pst,
				body_text,
				is_encrypted,
			)
			inserted = cursor.fetchone()
			if inserted is None or inserted[0] != message_id:
				raise RuntimeError(f"Failed to stage message_id {message_id}")
			inserted_count += 1
		connection.commit()
		return inserted_count
	except Exception:
		connection.rollback()
		raise
	finally:
		cursor.close()


def main() -> None:
	connection = pyodbc.connect(
		build_connection_string(load_config()["sqlserver"]),
		timeout=30,
		autocommit=False,
	)
	try:
		print(f"Staged {stage_watchoffice(connection)} Watch Office emails")
	finally:
		connection.close()


if __name__ == "__main__":
	main()
