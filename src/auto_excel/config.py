"""Configuration module for auto-excel.

Defines path constants for desktop working directory and related subdirectories.
"""

from pathlib import Path

BASE_DIR = Path.home() / "Desktop" / "marketing analysis"
RAW_DIR = BASE_DIR / "Raw"
NEW_DIR = BASE_DIR / "New"
LOG_DIR = BASE_DIR / "log"
STATE_FILE = LOG_DIR / "processed.json"
