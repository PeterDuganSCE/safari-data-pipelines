"""Load parsed Watch Office incidents from staging into the scrape table."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import pyodbc

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
	sys.path.insert(0, str(PROJECT_ROOT))

from pipelines.watchoffice.extract import UNKNOWN, extract_incident_data
from pipelines.watchoffice.s1_email_to_safari import build_connection_string, load_config


def known(value: Any) -> Any:
	return None if value == UNKNOWN else value


def scrape_record(body_text: str | None, received_datetime_pst: datetime | None) -> tuple[Any, ...]:
	parsed = extract_incident_data(body_text or "", email_date=received_datetime_pst)
	incident_date = known(parsed["incident_date"])
	incident_time = known(parsed["incident_time"])
	incident_datetime = (
		datetime.fromisoformat(f"{incident_date}T{incident_time}")
		if incident_date and incident_time else None
	)
	circuit = known(parsed["circuit"])
	substation = known(parsed["substation"])
	fire_department = known(parsed["fire_department"])
	for label, value in (
		("circuit", circuit),
		("substation", substation),
		("fire_extinguished_by", fire_department),
	):
		if value is not None and len(value) > 200:
			raise ValueError(f"{label} exceeds NVARCHAR(200)")

	return (
		incident_datetime,
		known(parsed.get("customers_disrupted")),
		circuit,
		float(voltage) if (voltage := known(parsed.get("voltage_kv"))) is not None else None,
		substation,
		{"Yes": True, "No": False}.get(parsed["hfra"]),
		{"Yes": True, "No": False}.get(parsed.get("fire")),
		fire_department,
	)


def scrape_watchoffice(connection: pyodbc.Connection) -> int:
	cursor = connection.cursor()
	try:
		cursor.execute(
			"""
SELECT source.[message_id], source.[body_text], source.[received_datetime_pst]
FROM [staging].[watchoffice] AS source
WHERE NOT EXISTS (
	SELECT 1 FROM [staging].[watchoffice_scrape] AS scraped
	WHERE scraped.[message_id] = source.[message_id]
)
ORDER BY source.[message_id]
"""
		)
		pending = cursor.fetchall()
		seen_ids: set[int] = set()
		inserted_count = 0
		for message_id, body_text, received_datetime_pst in pending:
			if message_id is None or message_id in seen_ids:
				raise ValueError(f"Missing or duplicate staging message_id: {message_id}")
			seen_ids.add(message_id)
			values = scrape_record(body_text, received_datetime_pst)
			cursor.execute(
				"""
INSERT INTO [staging].[watchoffice_scrape]
	([message_id], [incident_datetime_pst], [customers_disrupted],
	 [circuit], [voltage_kv], [substation], [hfra], [fire], [fire_extinguished_by])
OUTPUT INSERTED.[message_id]
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
""",
				message_id,
				*values,
			)
			inserted = cursor.fetchone()
			if inserted is None or inserted[0] != message_id:
				raise RuntimeError(f"Failed to scrape message_id {message_id}")
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
		print(f"Scraped {scrape_watchoffice(connection)} Watch Office emails")
	finally:
		connection.close()


if __name__ == "__main__":
	main()
