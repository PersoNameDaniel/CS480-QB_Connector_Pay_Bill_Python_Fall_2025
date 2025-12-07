"""Utilities for writing JSON reports."""

from __future__ import annotations

import json
from datetime import datetime, timezone, date
from pathlib import Path
from typing import Any, Dict


def _serialize_for_json(obj: Any) -> Any:
    """Convert non-JSON-serializable objects to JSON-compatible types."""
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()
    if isinstance(obj, dict):
        return {k: _serialize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_serialize_for_json(item) for item in obj]
    return obj


def save_json_report(discrepancies: Dict[str, Any], output_file: Path) -> None:
    """Save discrepancy report to JSON file."""
    try:
        # Convert datetime objects to ISO format strings
        serializable_data = _serialize_for_json(discrepancies)

        with open(output_file, "w") as f:
            json.dump(serializable_data, f, indent=2)
    except Exception as e:
        raise Exception(f"Failed to save JSON report to {output_file}: {e}")


def write_report(payload: Dict[str, Any], output_path: Path) -> Path:
    """Serialise the payload to JSON at output_path.

    The file is encoded as UTF-8 with an indent of two spaces. The payload is
    returned unchanged to ease chaining in higher-level functions.
    """

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)
    return output_path


def iso_timestamp() -> str:
    """Return a UTC timestamp suitable for JSON serialisation."""

    return datetime.now(timezone.utc).isoformat()


__all__ = ["save_json_report", "write_report", "iso_timestamp"]
