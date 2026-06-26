import logging
import yaml

from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional


def setup_logging(
        log_name: str,
    log_level: Optional[str] = None,
    format: Optional[str] = None,
        config_path: str = "config/config.yaml"
    ) -> logging.Logger:

    # load logging configuration from YAML file
    with open(config_path, "r") as f:
        config = yaml.safe_load(f)

    logging_config = config.get("logging", {})
    resolved_level = (log_level or logging_config.get("level", "INFO")).upper()
    resolved_format = format or logging_config.get(
        "format", "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    project_root = Path(config_path).resolve().parent.parent

    log_dir = project_root / config["logging"]["folder"]
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"{log_name}.log"

    # create logger
    logger = logging.getLogger(log_name)
    logger.setLevel(resolved_level)

    # Prevent duplicate handlers if called multiple times
    if logger.hasHandlers():
        logger.handlers.clear()

    # Formatter
    formatter = logging.Formatter(resolved_format)

    # File handler (with rotation!)
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=5 * 1024 * 1024,  # 5 MB
        backupCount=5
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.DEBUG)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    console_handler.setLevel(getattr(logging, resolved_level, logging.INFO))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger
