from pathlib import Path

# Anchored to this file's location — always resolves to the project root.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = PROJECT_ROOT / "config" / "config.yaml"
AUTH_PATH = PROJECT_ROOT / "config" / "auth.yaml"
