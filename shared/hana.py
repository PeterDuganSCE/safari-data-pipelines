import argparse
from typing import Any
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


def create_hana_connection(config: dict, logger: Logger) -> dbapi.Connection:
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