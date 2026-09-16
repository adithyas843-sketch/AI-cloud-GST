"""Local persistence for uploaded source files and reconciliation state."""
from __future__ import annotations

from pathlib import Path
import shutil
from typing import Any

import pandas as pd


DATA_DIR = Path(__file__).parent / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
STATE_FILE = DATA_DIR / "reconciliation_state.pkl"


def ensure_storage() -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


def save_uploaded_file(source_name: str, filename: str, content: bytes) -> Path:
    """Persist the latest uploaded source file using a safe, predictable filename."""
    ensure_storage()
    suffix = Path(filename).suffix.lower() or ".xlsx"
    safe_name = "_".join(source_name.lower().replace("-", " ").split())
    target = UPLOAD_DIR / f"{safe_name}{suffix}"
    target.write_bytes(content)
    return target


def save_state(state: dict[str, Any]) -> None:
    ensure_storage()
    pd.to_pickle(state, STATE_FILE)


def load_state() -> dict[str, Any] | None:
    if not STATE_FILE.exists():
        return None
    try:
        state = pd.read_pickle(STATE_FILE)
    except Exception:
        return None
    return state if isinstance(state, dict) else None


def clear_uploaded_files() -> None:
    if UPLOAD_DIR.exists():
        shutil.rmtree(UPLOAD_DIR)
    ensure_storage()


def clear_reconciliation_results() -> None:
    if STATE_FILE.exists():
        STATE_FILE.unlink()


def reset_application() -> None:
    if DATA_DIR.exists():
        shutil.rmtree(DATA_DIR)
    ensure_storage()
