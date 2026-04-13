"""State management module for auto-excel.

Tracks which files have been processed to avoid re-processing on subsequent runs.
State is persisted as a JSON file at config.STATE_FILE.
"""

import json
from datetime import datetime

from auto_excel.config import STATE_FILE


def load_state() -> dict:
    """Load the processing state from the state file.

    Returns an empty dict if the file does not exist or contains invalid JSON.
    """
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError:
        return {}


def save_file_state(filename: str) -> None:
    """Record a file as successfully processed.

    Loads the current state, adds/updates the entry for *filename*, and
    immediately writes the result back to STATE_FILE.
    """
    state = load_state()
    state[filename] = {
        "processed_at": datetime.now().isoformat(),
        "status": "success",
    }
    STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")


def is_processed(filename: str) -> bool:
    """Return True if *filename* has already been recorded in the state file."""
    return filename in load_state()
