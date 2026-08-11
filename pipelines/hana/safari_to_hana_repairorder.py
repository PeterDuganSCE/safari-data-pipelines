"""
SQL Server to SAP HANA ETL Pipeline

Purpose:
	Extract data from a SQL Server table, transform it using pandas,
	and load it into a SAP HANA table.

Install required packages:
	pip install pandas sqlalchemy pyodbc hdbcli

Run:
	python -m pipelines.hana.safari_to_hana_repairorder
"""

import argparse
import datetime
import numbers
import os
import re
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, Iterator, Optional

import pandas as pd
import yaml
from hdbcli import dbapi
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

from shared.logging import setup_logging
from shared.paths import AUTH_PATH, CONFIG_PATH, PROJECT_ROOT

# ---------------------------------------------------------------------
# Logging configuration
# ---------------------------------------------------------------------
logger = setup_logging("safari_to_hana_repairorder")


# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
logger.info("Project root directory: %s", PROJECT_ROOT)
logger.info("Using config file: %s", CONFIG_PATH)
logger.info("Using auth file: %s", AUTH_PATH)

PIPELINE_DEFAULTS = {
	"sqlserver": {
		"host": "D259321",  # ip address or hostname
		"port": 49172,
		"database": "RepairOrder",
		"driver": "ODBC Driver 17 for SQL Server",
		"schema": "dbo",
		"table": "vw_AllRepairOrderData",
	},
	"hana": {
		"host": "vp55db51.sce.com",
		"port": 30015,
		"schema": "SCE_TD",
		"table": "FI_SAFARI_REPAIR_ORDERS_2",
	},
	"etl": {
		# Number of rows extracted from SQL Server at a time
		"chunksize": 50000,

		# Number of rows inserted into SAP HANA per batch
		"insert_batch_size": 5000,
	},
}


def merge_nested_dicts(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
	"""Recursively merges override values into base dictionary."""
	merged = dict(base)

	for key, value in override.items():
		if isinstance(value, dict) and isinstance(merged.get(key), dict):
			merged[key] = merge_nested_dicts(merged[key], value)
		else:
			merged[key] = value

	return merged


def load_pipeline_config(
	config_path: Path = CONFIG_PATH,
	auth_path: Path = AUTH_PATH,
) -> Dict[str, Any]:
	"""Loads and merges base config and auth config for pipeline execution."""
	with open(config_path, "r", encoding="utf-8") as config_file:
		config_yaml = yaml.safe_load(config_file) or {}

	with open(auth_path, "r", encoding="utf-8") as auth_file:
		auth_yaml = yaml.safe_load(auth_file) or {}

	merged = merge_nested_dicts(PIPELINE_DEFAULTS, config_yaml)
	merged = merge_nested_dicts(merged, auth_yaml)

	# Allow secrets to be provided through environment variables when not in auth.yaml
	merged.setdefault("sqlserver", {})
	merged["sqlserver"].setdefault("username", os.getenv("SQLSERVER_USERNAME"))
	merged["sqlserver"].setdefault("password", os.getenv("SQLSERVER_PASSWORD"))

	return merged


# ---------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------

def validate_config(config: dict) -> None:
	"""
	Ensures required configuration values are present before running ETL.
	"""
	required_fields = [
		("sqlserver", "host"),
		("sqlserver", "database"),
		("sqlserver", "username"),
		("sqlserver", "password"),
		("sqlserver", "schema"),
		("sqlserver", "table"),
		("hana", "host"),
		("hana", "port"),
		("hana", "username"),
		("hana", "password"),
		("hana", "schema"),
		("hana", "table"),
	]

	missing = []

	for section, key in required_fields:
		if config.get(section, {}).get(key) in [None, ""]:
			missing.append(f"{section}.{key}")

	if missing:
		raise ValueError(
			"Missing required configuration values: "
			+ ", ".join(missing)
		)


def quote_hana_identifier(identifier: str) -> str:
	"""
	Safely quotes SAP HANA identifiers such as schema, table, and column names.
	"""
	escaped = identifier.replace('"', '""')
	return f'"{escaped}"'


# ---------------------------------------------------------------------
# Connection functions
# ---------------------------------------------------------------------

def create_sqlserver_engine(config: dict) -> Engine:
	"""
	Creates a SQLAlchemy engine for SQL Server using pyodbc.
	"""
	sql_cfg = config["sqlserver"]

	server = sql_cfg["host"]
	if sql_cfg.get("port"):
		server = f"{server},{sql_cfg['port']}"

	connection_url = (
		"mssql+pyodbc://"
		f"{sql_cfg['username']}:{sql_cfg['password']}"
		f"@{server}/{sql_cfg['database']}"
		f"?driver={sql_cfg['driver'].replace(' ', '+')}"
		"&TrustServerCertificate=yes"
	)

	logger.info("Creating SQL Server engine.")

	return create_engine(
		connection_url,
		fast_executemany=True,
	)


def create_hana_connection(config: dict):
	"""
	Creates a SAP HANA connection using hdbcli.
	"""
	hana_cfg = config["hana"]

	logger.info("Creating SAP HANA connection.")

	connection = dbapi.connect(
		address=hana_cfg["host"],
		port=hana_cfg["port"],
		user=hana_cfg["username"],
		password=hana_cfg["password"],
	)
	# Auto-commit must be off to allow LOB (NCLOB) streaming
	connection.setautocommit(False)
	return connection


# ---------------------------------------------------------------------
# Extract
# ---------------------------------------------------------------------

def extract_sqlserver_data(
	engine: Engine,
	source_schema: str,
	source_table: str,
	chunksize: int,
) -> Iterator[pd.DataFrame]:
	"""
	Extracts data from SQL Server in chunks.

	Chunking helps avoid loading the entire source table into memory.
	"""
	query = text(f"""
		SELECT *
		FROM [{source_schema}].[{source_table}]
	""")

	logger.info(
		"Starting extraction from SQL Server table [%s].[%s].",
		source_schema,
		source_table,
	)

	try:
		with engine.connect() as connection:
			for chunk_df in pd.read_sql_query(
				sql=query,
				con=connection,
				chunksize=chunksize,
			):
				logger.info("Extracted chunk with %s rows.", len(chunk_df))
				yield chunk_df

	except Exception:
		logger.exception("Failed while extracting data from SQL Server.")
		raise


# ---------------------------------------------------------------------
# Transform
# ---------------------------------------------------------------------

def transform_dataframe(df: pd.DataFrame) -> pd.DataFrame:
	"""
	Applies common dataframe transformations.
	"""
	logger.info("Starting transformation for %s rows.", len(df))

	transformed = df.copy()

	try:
		transformed.rename(columns={
			# Add source column renames here as needed, e.g.:
			"Voltage (kV)": "VOLTAGE_KV",
			"Category": "EQUIP_CATEGORY",
			"Specifics": "EQUIP_SPECIFICS",
			"Sub-Component": "EQUIP_SUB_COMPONENT",
			"Location": "EQUIP_LOCATION",
			"Voltage Classification": "EQUIP_VOLTAGE",
			"Material": "EQUIP_MATERIAL",
			"size": "EQUIP_SIZE",
			"Application": "EQUIP_APPLICATION",
			"Control Type": "EQUIP_CONTROL_TYPE",
			"Orientation": "EQUIP_ORIENTATION",
			"kVA Rating": "EQUIP_KVA_RATING",
			"Age": "EQUIP_AGE",
			"Travel Method": "EQUIP_TRAVEL_METHOD",
			"Primary Failed Equipment": "EQUIP_PRIMARY_FAILED",
		}, inplace=True)

		transformed.columns = transformed.columns.map(lambda x: x.replace(" ", "_").replace("/", "_"))
		transformed.columns = transformed.columns.str.upper()
		transformed["REFRESH_DATE"] = pd.Timestamp("today")

		logger.info("Transformation completed successfully.")
		return transformed

	except Exception:
		logger.exception("Failed while transforming dataframe.")
		raise


# ---------------------------------------------------------------------
# Load helpers
# ---------------------------------------------------------------------

def prepare_hana_rows(df: pd.DataFrame) -> list:
	"""
	Converts pandas values to SAP HANA-compatible Python values.

	Handles:
		pandas NA / NaN / NaT -> None
		pandas Timestamp -> Python datetime
		numeric values -> int, bool, Decimal
	"""

	def convert_value(value):
		if pd.isna(value):
			return None

		# Guard against stringified nulls that pd.isna() won't catch
		if isinstance(value, str) and value.strip().lower() in ("none", "nan", "nat", "null", ""):
			return None

		if isinstance(value, bool):
			return bool(value)

		if isinstance(value, pd.Timestamp):
			return value.to_pydatetime()

		if isinstance(value, datetime.datetime):
			return value

		if isinstance(value, datetime.date) and not isinstance(value, datetime.datetime):
			return value

		if isinstance(value, datetime.time):
			return value

		if isinstance(value, Decimal):
			return value

		if isinstance(value, numbers.Integral):
			return int(value)

		if isinstance(value, numbers.Real):
			return Decimal(str(value))

		return str(value)

	rows = []
	for record in df.itertuples(index=False, name=None):
		rows.append(tuple(convert_value(value) for value in record))

	return rows


def hana_table_exists(
	hana_connection,
	target_schema: str,
	target_table: str,
) -> bool:
	"""Returns True if the target SAP HANA table already exists."""
	cursor = hana_connection.cursor()
	try:
		cursor.execute(
			"""
			SELECT COUNT(*)
			FROM SYS.TABLES
			WHERE SCHEMA_NAME = ?
			  AND TABLE_NAME = ?
			""",
			(target_schema, target_table),
		)
		result = cursor.fetchone()
		return bool(result and result[0] > 0)
	finally:
		cursor.close()


# Add column name -> SAP HANA type mappings for all source columns here.
HANA_COLUMN_TYPE_MAP = {
	"ID": "BIGINT NOT NULL",
	"TITLE": "NVARCHAR(255)",
	"ENTRY_METHOD": "NVARCHAR(255)",
	"EMAIL_DATE": "TIMESTAMP",
	"INCIDENT_DATE": "TIMESTAMP",
	"REPORTED_BY_EMAIL": "NVARCHAR(255)",
	"DISTRICT_NAME": "NVARCHAR(255)",
	"DISTRICT_NUMBER": "INTEGER",
	"EVENT_TYPE": "NVARCHAR(255)",
	"ENGINEER_ASSIGNED": "NVARCHAR(255)",
	"HFRA_CIRCUIT": "NVARCHAR(255)",
	"CIRCUIT": "NVARCHAR(255)",
	"VOLTAGE_KV": "DOUBLE",
	"STRUCTURE": "NVARCHAR(255)",
	"STRUCTURE_TYPE": "NVARCHAR(255)",
	"HFRA_STRUCTURE": "NVARCHAR(255)",
	"SUBSTATION": "NVARCHAR(255)",
	"SWITCHING_CENTER": "NVARCHAR(255)",
	"COMMENTS": "NCLOB",
	"EVENT_DRIVER": "NVARCHAR(255)",
	"EVENT_DRIVER_SPECIFICS": "NVARCHAR(255)",
	"VOLTAGE_INVOLVED": "NVARCHAR(255)",
	"REVIEWED_DATE": "TIMESTAMP",
	"CONTRIBUTING_FACTORS": "NVARCHAR(255)",
	"CC_ON_STRUCTURE_CHECKBOX": "NVARCHAR(255)",
	"INFO_SOURCES": "NVARCHAR(255)",
	"SEQ_NUM": "NVARCHAR(255)",
	"SHAREPOINT_FOLDER": "NVARCHAR(255)",
	"CC_ON_STRUCTURE": "NVARCHAR(255)",
	"CREATED": "TIMESTAMP",
	"CREATED_BY": "NVARCHAR(255)",
	"DISTRICT_INBOX": "NVARCHAR(255)",
	"MODIFIED": "TIMESTAMP",
	"MODIFIED_BY": "NVARCHAR(255)",
	"REVIEWED": "BOOLEAN",
	"GMC_ZONE": "NVARCHAR(255)",
	"INBOX_EMAIL": "NVARCHAR(255)",
	"REGION": "NVARCHAR(255)",
	"ATTACHMENT_NAME": "NVARCHAR(255)",
	"FIRE_LENGTH": "DOUBLE",
	"FIRE_WIDTH": "DOUBLE",
	"FIRE_ACRES": "DOUBLE",
	"WIRE_DETACHED_FROM_STRUCTURE": "NVARCHAR(255)",
	"WIRE_IN_REACH_OF_PUBLIC": "NVARCHAR(255)",
	"WIRE_IN_CONTACT_WITH_OBJECT": "NVARCHAR(255)",
	"WIRE_SEPARATION": "NVARCHAR(255)",
	"LINE_OR_LOAD_SIDE_DOWN": "NVARCHAR(255)",
	"ENERGIZED_UPON_ARRIVAL": "NVARCHAR(255)",
	"UG_EXPLOSION_TYPE": "NVARCHAR(255)",
	"EQUIP_ID": "NVARCHAR(255)",
	"EQUIP_CATEGORY": "NVARCHAR(255)",
	"EQUIP_SPECIFICS": "NVARCHAR(255)",
	"EQUIP_SUB_COMPONENT": "NVARCHAR(255)",
	"EQUIP_LOCATION": "NVARCHAR(255)",
	"EQUIP_VOLTAGE": "NVARCHAR(255)",
	"EQUIP_MATERIAL": "NVARCHAR(255)",
	"EQUIP_SIZE": "NVARCHAR(255)",
	"EQUIP_APPLICATION": "NVARCHAR(255)",
	"EQUIP_CONTROL_TYPE": "NVARCHAR(255)",
	"EQUIP_ORIENTATION": "NVARCHAR(255)",
	"EQUIP_KVA_RATING": "NVARCHAR(255)",
	"EQUIP_AGE": "NVARCHAR(255)",
	"EQUIP_TRAVEL_METHOD": "NVARCHAR(255)",
	"EQUIP_PRIMARY_FAILED": "NVARCHAR(255)",
	"EQUIP_CREATED": "TIMESTAMP",
	"EQUIP_MODIFIED": "TIMESTAMP",
	"RO_PROBLEM": "NVARCHAR(255)",
	"RO_CAUSE": "NVARCHAR(255)",
	"TM_REMARKS": "NCLOB",
	"TM_RADIO_NO": "NVARCHAR(255)",
	"CAD_ID": "NVARCHAR(255)",
	"SEQ_NO": "NVARCHAR(255)",
	"ATTACHMENT_URL": "NVARCHAR(1024)",
	"REFRESH_DATE": "TIMESTAMP",
}


def build_hana_column_type_map(df: pd.DataFrame) -> dict:
	"""Returns a hard-coded mapping of transformed dataframe column names to SAP HANA data types."""
	missing_columns = []
	hana_column_types = {}

	for column in df.columns:
		if column in HANA_COLUMN_TYPE_MAP:
			hana_column_types[column] = HANA_COLUMN_TYPE_MAP[column]
		else:
			missing_columns.append(column)

	if missing_columns:
		logger.warning(
			"The following transformed columns were not found in HANA_COLUMN_TYPE_MAP and will be skipped: %s",
			missing_columns,
		)

	return hana_column_types


def create_hana_table_if_missing(
	hana_connection,
	df: pd.DataFrame,
	target_schema: str,
	target_table: str,
	hana_column_types: Optional[dict] = None,
) -> None:
	"""Creates the target SAP HANA table from the DataFrame if it does not exist."""
	if hana_table_exists(hana_connection, target_schema, target_table):
		return

	if hana_column_types is None:
		hana_column_types = build_hana_column_type_map(df)

	column_defs = []
	for column, column_type in hana_column_types.items():
		quoted_column = quote_hana_identifier(column)
		column_defs.append(f"{quoted_column} {column_type}")

	quoted_schema = quote_hana_identifier(target_schema)
	quoted_table = quote_hana_identifier(target_table)
	create_sql = (
		f"CREATE COLUMN TABLE {quoted_schema}.{quoted_table} ("
		+ ", ".join(column_defs)
		+ ")"
	)

	logger.info(
		"Creating SAP HANA table %s.%s because it does not exist.",
		target_schema,
		target_table,
	)

	cursor = hana_connection.cursor()
	try:
		cursor.execute(create_sql)
		hana_connection.commit()
		logger.info("SAP HANA target table created successfully.")
	except Exception:
		logger.exception("Failed to create SAP HANA target table.")
		hana_connection.rollback()
		raise
	finally:
		cursor.close()


# ---------------------------------------------------------------------
# Load
# ---------------------------------------------------------------------

def load_dataframe_to_hana(
	hana_connection,
	df: pd.DataFrame,
	target_schema: str,
	target_table: str,
	batch_size: int,
	hana_column_types: Optional[dict] = None,
) -> None:
	"""
	Loads transformed dataframe rows into SAP HANA using parameterized inserts.
	"""
	if df.empty:
		logger.info("Dataframe is empty. Skipping SAP HANA load.")
		return

	create_hana_table_if_missing(
		hana_connection=hana_connection,
		df=df,
		target_schema=target_schema,
		target_table=target_table,
		hana_column_types=hana_column_types,
	)

	columns = list(df.columns)

	quoted_schema = quote_hana_identifier(target_schema)
	quoted_table = quote_hana_identifier(target_table)
	quoted_columns = ", ".join(quote_hana_identifier(col) for col in columns)
	placeholders = ", ".join(["?"] * len(columns))

	insert_sql = f"""
		INSERT INTO {quoted_schema}.{quoted_table}
		({quoted_columns})
		VALUES ({placeholders})
	"""

	logger.info(
		"Starting SAP HANA load into %s.%s for %s rows.",
		target_schema,
		target_table,
		len(df),
	)

	cursor = None

	try:
		cursor = hana_connection.cursor()
		rows = prepare_hana_rows(df)

		for start in range(0, len(rows), batch_size):
			batch = rows[start:start + batch_size]
			cursor.executemany(insert_sql, batch)

			logger.info(
				"Inserted batch rows %s to %s.",
				start + 1,
				start + len(batch),
			)

		hana_connection.commit()
		logger.info("SAP HANA load committed successfully.")

	except Exception:
		logger.exception("Failed while loading data into SAP HANA. Rolling back.")
		hana_connection.rollback()
		raise

	finally:
		if cursor is not None:
			cursor.close()


# ---------------------------------------------------------------------
# Optional helper for full refresh
# ---------------------------------------------------------------------

def truncate_hana_table(
	hana_connection,
	target_schema: str,
	target_table: str,
) -> None:
	"""
	Optional helper if this pipeline should perform a full refresh.

	Use carefully. Uncomment the call in run_pipeline() if you want to clear
	the target table before loading new data.
	"""
	quoted_schema = quote_hana_identifier(target_schema)
	quoted_table = quote_hana_identifier(target_table)

	sql = f"TRUNCATE TABLE {quoted_schema}.{quoted_table}"

	logger.info("Truncating SAP HANA table %s.%s.", target_schema, target_table)

	cursor = None

	try:
		cursor = hana_connection.cursor()
		cursor.execute(sql)
		hana_connection.commit()
		logger.info("SAP HANA target table truncated successfully.")

	except Exception:
		logger.exception("Failed while truncating SAP HANA table.")
		hana_connection.rollback()
		raise

	finally:
		if cursor is not None:
			cursor.close()


def drop_hana_table(
	hana_connection,
	target_schema: str,
	target_table: str,
) -> None:
	"""Drops the target SAP HANA table if it exists."""
	if not hana_table_exists(hana_connection, target_schema, target_table):
		logger.info(
			"SAP HANA table %s.%s does not exist. Nothing to drop.",
			target_schema,
			target_table,
		)
		return

	quoted_schema = quote_hana_identifier(target_schema)
	quoted_table = quote_hana_identifier(target_table)
	sql = f"DROP TABLE {quoted_schema}.{quoted_table}"

	logger.info("Dropping SAP HANA table %s.%s.", target_schema, target_table)

	cursor = None
	try:
		cursor = hana_connection.cursor()
		cursor.execute(sql)
		hana_connection.commit()
		logger.info("SAP HANA table dropped successfully.")
	except Exception:
		logger.exception("Failed while dropping SAP HANA table.")
		hana_connection.rollback()
		raise
	finally:
		if cursor is not None:
			cursor.close()


# ---------------------------------------------------------------------
# Main orchestration
# ---------------------------------------------------------------------

def run_pipeline(config: dict, drop_and_recreate: bool = False) -> None:
	"""
	Main ETL orchestration function.
	"""
	validate_config(config)

	sql_engine: Optional[Engine] = None
	hana_connection = None
	total_rows_loaded = 0

	try:
		sql_engine = create_sqlserver_engine(config)
		hana_connection = create_hana_connection(config)

		source_schema = config["sqlserver"]["schema"]
		source_table = config["sqlserver"]["table"]

		target_schema = config["hana"]["schema"]
		target_table = config["hana"]["table"]

		chunksize = config["etl"]["chunksize"]
		insert_batch_size = config["etl"]["insert_batch_size"]

		if drop_and_recreate:
			drop_hana_table(hana_connection, target_schema, target_table)

		# Uncomment this line if the pipeline should fully replace target data.
		truncate_hana_table(hana_connection, target_schema, target_table)

		for source_chunk in extract_sqlserver_data(
			engine=sql_engine,
			source_schema=source_schema,
			source_table=source_table,
			chunksize=chunksize,
		):
			transformed_chunk = transform_dataframe(source_chunk)

			hana_column_types = build_hana_column_type_map(transformed_chunk)
			logger.info(
				"SAP HANA column types from transformed chunk: %s",
				hana_column_types,
			)

			load_dataframe_to_hana(
				hana_connection=hana_connection,
				df=transformed_chunk,
				target_schema=target_schema,
				target_table=target_table,
				batch_size=insert_batch_size,
				hana_column_types=hana_column_types,
			)

			total_rows_loaded += len(transformed_chunk)

		logger.info(
			"ETL pipeline completed successfully. Total rows loaded: %s",
			total_rows_loaded,
		)

	except Exception:
		logger.exception("ETL pipeline failed.")
		raise

	finally:
		if sql_engine is not None:
			sql_engine.dispose()
			logger.info("SQL Server engine disposed.")

		if hana_connection is not None:
			hana_connection.close()
			logger.info("SAP HANA connection closed.")

		logger.info("-" * 50)


if __name__ == "__main__":
	parser = argparse.ArgumentParser()
	parser.add_argument(
		"--drop-and-recreate",
		action="store_true",
		help="Drop the target HANA table before loading so it is fully recreated.",
	)
	args = parser.parse_args()
	run_pipeline(load_pipeline_config(), drop_and_recreate=args.drop_and_recreate)
