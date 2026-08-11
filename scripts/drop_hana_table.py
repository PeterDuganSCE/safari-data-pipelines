"""
Utility script to drop a SAP HANA table.

Run:
    python -m scripts.drop_hana_table --schema MY_SCHEMA --table MY_TABLE
"""

import argparse
import os
from pathlib import Path
from typing import Any, Dict

import yaml
from hdbcli import dbapi

from shared.logging import setup_logging
from shared.paths import AUTH_PATH, CONFIG_PATH

logger = setup_logging("drop_hana_table")


def load_hana_credentials(
	config_path: Path = CONFIG_PATH,
	auth_path: Path = AUTH_PATH,
) -> Dict[str, Any]:
	with open(config_path, "r", encoding="utf-8") as f:
		config = yaml.safe_load(f) or {}
	with open(auth_path, "r", encoding="utf-8") as f:
		auth = yaml.safe_load(f) or {}

	hana = {**config.get("hana", {}), **auth.get("hana", {})}
	hana.setdefault("username", os.getenv("HANA_USERNAME"))
	hana.setdefault("password", os.getenv("HANA_PASSWORD"))
	return hana


def drop_table(schema: str, table: str) -> None:
	creds = load_hana_credentials()

	connection = dbapi.connect(
		address=creds["host"],
		port=creds["port"],
		user=creds["username"],
		password=creds["password"],
	)
	connection.setautocommit(False)

	quoted_schema = f'"{schema}"'
	quoted_table = f'"{table}"'

	cursor = connection.cursor()
	try:
		cursor.execute(
			"SELECT COUNT(*) FROM SYS.TABLES WHERE SCHEMA_NAME = ? AND TABLE_NAME = ?",
			(schema, table),
		)
		exists = cursor.fetchone()[0] > 0

		if not exists:
			logger.info("Table %s.%s does not exist. Nothing to drop.", schema, table)
			return

		logger.info("Dropping table %s.%s.", schema, table)
		cursor.execute(f"DROP TABLE {quoted_schema}.{quoted_table}")
		connection.commit()
		logger.info("Table %s.%s dropped successfully.", schema, table)

	except Exception:
		logger.exception("Failed to drop table %s.%s.", schema, table)
		connection.rollback()
		raise
	finally:
		cursor.close()
		connection.close()


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description="Drop a SAP HANA table.")
	parser.add_argument("--schema", required=True, help="HANA schema name")
	parser.add_argument("--table", required=True, help="HANA table name")
	args = parser.parse_args()

	drop_table(schema=args.schema, table=args.table)
