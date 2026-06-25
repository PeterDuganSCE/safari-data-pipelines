"""
SQL Server to SAP HANA ETL Pipeline

Purpose:
    Extract data from a SQL Server table, transform it using pandas,
    and load it into a SAP HANA table.

Install required packages:
    pip install pandas sqlalchemy pyodbc hdbcli

Run:
    python sqlserver_to_hana_etl.py
"""

import sys
import logging
import datetime
import numbers
from decimal import Decimal
from typing import Iterator, Optional

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from hdbcli import dbapi


# ---------------------------------------------------------------------
# Logging configuration
# ---------------------------------------------------------------------

log_file = "etl.log"
formatter = logging.Formatter(
    "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(formatter)

file_handler = logging.FileHandler(log_file, mode="w", encoding="utf-8")
file_handler.setFormatter(formatter)

logger = logging.getLogger("sqlserver_to_hana_etl")
logger.setLevel(logging.INFO)
logger.addHandler(console_handler)
logger.addHandler(file_handler)


# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
# Update these values directly.
# Credentials are shown as placeholders.
# In production, avoid committing real usernames/passwords to source control.

CONFIG = {
    "sqlserver": {
        "host": "D259321,49172",  # ip address or hostname, include port if needed
        "database": "PROD_SAFARI",
        "username": "FIPA_APP",
        "password": "2023Welc0meFIPA",
        "driver": "ODBC Driver 17 for SQL Server",
        "schema": "dbo",
        "table": "vw_UGEAllFields",
    },
    "hana": {
        "host": "vp55db.sce.com",
        "port": 30015,
        "username": "duganpr",
        "password": "BlueForest1",
        "schema": "SCE_TD",
        "table": "FI_SAFARI_UGE",
    },
    "etl": {
        # Number of rows extracted from SQL Server at a time
        "chunksize": 50000,

        # Number of rows inserted into SAP HANA per batch
        "insert_batch_size": 5000,
    },
}


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

    connection_url = (
        "mssql+pyodbc://"
        f"{sql_cfg['username']}:{sql_cfg['password']}"
        f"@{sql_cfg['host']}/{sql_cfg['database']}"
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

    return dbapi.connect(
        address=hana_cfg["host"],
        port=hana_cfg["port"],
        user=hana_cfg["username"],
        password=hana_cfg["password"],
    )


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
    Applies explicit pandas dtype conversions.

    Expected source columns:
        id          INT
        event_date  DATETIME
        description VARCHAR
        amount      DECIMAL
    """
    logger.info("Starting transformation for %s rows.", len(df))

    transformed = df.copy()

    try:
        # Convert column names
        transformed.rename(columns={
            "Date/Time": "Date Time",
            "ILS No.": "ILS No"
        }, inplace=True)

        rename_func = lambda x: x.replace(' ', '_')
        transformed.columns = transformed.columns.map(rename_func)
        transformed.columns = transformed.columns.str.upper()
        transformed['REFRESH_DATE'] = pd.Timestamp('today')

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
    quoted_schema = quote_hana_identifier(target_schema)
    quoted_table = quote_hana_identifier(target_table)

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


def hana_type_for_dtype(dtype) -> str:
    """Maps a pandas dtype to a SAP HANA column type."""
    if pd.api.types.is_bool_dtype(dtype):
        return "BOOLEAN"

    if pd.api.types.is_integer_dtype(dtype):
        return "BIGINT"

    if pd.api.types.is_float_dtype(dtype):
        return "DOUBLE"

    if pd.api.types.is_datetime64_any_dtype(dtype):
        return "TIMESTAMP"

    if pd.api.types.is_timedelta64_dtype(dtype):
        return "BIGINT"

    if pd.api.types.is_string_dtype(dtype) or pd.api.types.is_object_dtype(dtype):
        return "NVARCHAR(5000)"

    return "NVARCHAR(5000)"


HANA_COLUMN_TYPE_MAP = {
    "INCIDENT_ID": "INTEGER NOT NULL",
    "PARENT_ID": "INTEGER",
    "EVENT_ID": "INTEGER NOT NULL",
    "TITLE": "NVARCHAR(100)",
    "DATE_TIME": "TIMESTAMP",
    "STATUS": "NVARCHAR(100)",
    "ENGINEER": "NVARCHAR(100)",
    "UGE_CATEGORY": "NVARCHAR(100)",
    "UGE_CATEGORY_DETAILS": "NVARCHAR(200)",
    "ENERGY_RELEASE": "NVARCHAR(100)",
    "WATER_PRESENT": "NVARCHAR(100)",
    "FAILED_WITHIN_24HRS_OF_PUMPING": "NVARCHAR(100)",
    "CPRR_COMPLETED": "NVARCHAR(100)",
    "CPRR_PRESENT": "NVARCHAR(100)",
    "CPRR_EFFECTIVE": "NVARCHAR(100)",
    "CIRCUIT_NAME": "NVARCHAR(100)",
    "VOLTAGE": "DOUBLE",
    "VOLTAGE_FACILITY": "NVARCHAR(100)",
    "SUBSTATION": "NVARCHAR(100)",
    "SWITCHING_CENTER": "NVARCHAR(100)",
    "ELECTRICAL_SYSTEM": "NVARCHAR(100)",
    "SAIDI_RANK": "INTEGER",
    "CIRCUIT_RELIABILITY": "NVARCHAR(100)",
    "PROTECTION_TYPE": "NVARCHAR(100)",
    "OPERATED_PROTECTION": "NVARCHAR(100)",
    "FAST_CURVE_ENABLED": "NVARCHAR(100)",
    "OPERATION_COUNT": "INTEGER",
    "LOCKOUT": "NVARCHAR(50)",
    "INTERRUPTIONS": "INTEGER",
    "OMS_ID": "INTEGER",
    "ILS_NO": "INTEGER",
    "RECLOSE_COUNT": "INTEGER",
    "CABLE_TYPE": "NVARCHAR(100)",
    "CABLE_SIZE": "NVARCHAR(100)",
    "TRAVEL_METHOD": "NVARCHAR(100)",
    "MAIN_OR_TAP": "NVARCHAR(100)",
    "INSTALL_YEAR": "INTEGER",
    "FLOC": "NVARCHAR(100)",
    "STRUCTURE_TYPE": "NVARCHAR(100)",
    "STRUCTURE_CATEGORY": "NVARCHAR(100)",
    "LAT": "DOUBLE",
    "LONG": "DOUBLE",
    "ADDRESS": "NVARCHAR(500)",
    "START_UP": "NVARCHAR(50)",
    "DISTRICT_NO": "INTEGER",
    "DISTRICT_NAME": "NVARCHAR(50)",
    "HFRA": "NVARCHAR(50)",
    "FUEL_BED": "NVARCHAR(50)",
    "LAND_USE": "NVARCHAR(50)",
    "COMMENTS": "NVARCHAR(5000)",
    "KEY_LEARNINGS": "NVARCHAR(5000)",
    "ROOT_CAUSE": "NVARCHAR(100)",
    "ROOT_SPECIFICS": "NVARCHAR(500)",
    "ROOT_FAULT_TYPE": "NVARCHAR(50)",
    "ROOT_FAULT_LOCATION": "NVARCHAR(100)",
    "ROOT_FAULT_MAGNITUDE": "NVARCHAR(50)",
    "ROOT_EQUIP_LOCATION": "NVARCHAR(50)",
    "ROOT_EQUIP_CATEGORY": "NVARCHAR(500)",
    "ROOT_EQUIP_SUBCATEGORY": "NVARCHAR(500)",
    "INTERMEDIATE_CAUSE": "NVARCHAR(100)",
    "INTERMEDIATE_SPECIFICS": "NVARCHAR(500)",
    "INTERMEDIATE_FAULT_TYPE": "NVARCHAR(50)",
    "INTERMEDIATE_FAULT_LOCATION": "NVARCHAR(100)",
    "INTERMEDIATE_FAULT_MAGNITUDE": "NVARCHAR(50)",
    "INTERMEDIATE_EQUIP_LOCATION": "NVARCHAR(50)",
    "INTERMEDIATE_EQUIP_CATEGORY": "NVARCHAR(500)",
    "INTERMEDIATE_EQUIP_SUBCATEGORY": "NVARCHAR(500)",
    "PRIMARY_CAUSE": "NVARCHAR(100)",
    "PRIMARY_SPECIFICS": "NVARCHAR(500)",
    "PRIMARY_FAULT_TYPE": "NVARCHAR(50)",
    "PRIMARY_FAULT_LOCATION": "NVARCHAR(50)",
    "PRIMARY_FAULT_MAGNITUDE": "NVARCHAR(50)",
    "PRIMARY_EQUIP_LOCATION": "NVARCHAR(50)",
    "PRIMARY_EQUIP_CATEGORY": "NVARCHAR(500)",
    "PRIMARY_EQUIP_SUBCATEGORY": "NVARCHAR(500)",
    "ADI/ATI": "NVARCHAR(50)",
    "LAST_PLP": "NVARCHAR(50)",
    "OH_IR_DATE": "NVARCHAR(50)",
    "EOI_DATE": "NVARCHAR(50)",
    "LAST_EOI": "NVARCHAR(50)",
    "LAST_IPI": "NVARCHAR(50)",
    "LSI": "NVARCHAR(50)",
    "LAST_ODI/UDI": "NVARCHAR(50)",
    "INFO_SOURCE": "NVARCHAR(50)",
    "REPAIR_ORDER": "INTEGER",
    "CAD_ID": "NVARCHAR(50)",
    "SEQ_NO": "NVARCHAR(50)",
    "SHAREPOINT": "NVARCHAR(5000)",
    "ALL_CIRCUITS_INVOLVED": "NVARCHAR(5000)",
    "ALL_STRUCTURES_INVOLVED": "NVARCHAR(5000)",
    "NOTIFICATION": "INTEGER",
    "WORK_ORDER": "INTEGER",
    "CREATED": "TIMESTAMP",
    "ASSIGNED": "TIMESTAMP",
    "COMPLETED": "TIMESTAMP",
    "REFRESH_DATE": "TIMESTAMP",
}


def build_hana_column_type_map(df: pd.DataFrame) -> dict:
    """Returns a hard-coded mapping of transformed dataframe column names to SAP HANA data types."""
    missing_columns = [
        column for column in df.columns
        if column not in HANA_COLUMN_TYPE_MAP
    ]

    if missing_columns:
        logger.warning(
            "The following transformed columns were not found in HANA_COLUMN_TYPE_MAP and will be skipped: %s",
            missing_columns,
        )

    return {
        column: HANA_COLUMN_TYPE_MAP[column]
        for column in df.columns
        if column in HANA_COLUMN_TYPE_MAP
    }


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


# ---------------------------------------------------------------------
# Main orchestration
# ---------------------------------------------------------------------

def run_pipeline(config: dict) -> None:
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

        # Uncomment this line if the pipeline should fully replace target data.
        truncate_hana_table(hana_connection, target_schema, target_table)

        for source_chunk in extract_sqlserver_data(
            engine=sql_engine,
            source_schema=source_schema,
            source_table=source_table,
            chunksize=chunksize,
        ):
            transformed_chunk = transform_dataframe(source_chunk)

            print(transformed_chunk.head())  # Debug: Show first few rows of transformed data
            print(transformed_chunk.dtypes.to_string())  # Debug: Show data types of transformed data

            hana_column_types = build_hana_column_type_map(transformed_chunk)
            logger.info(
                "Explicit SAP HANA column types from transformed chunk: %s",
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


if __name__ == "__main__":
    run_pipeline(CONFIG)