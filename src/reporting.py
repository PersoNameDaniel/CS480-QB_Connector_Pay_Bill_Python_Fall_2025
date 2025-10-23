"""Utilities for writing JSON reports."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict


def save_json_report(data: Dict[str, Any], output_path: str | Path) -> None:
    """Save the comparison results to a JSON file.

    Args:
        data: The dictionary containing discrepancy results.
        output_path: Path where the JSON file should be saved.

    Raises:
        OSError: If the file cannot be written.
    """
    path = Path(output_path)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        # Print(f"Report successfully saved to {path}")
        print(f"Report successfully saved to {path}")
    except Exception as e:
        raise OSError(f"Failed to save JSON report to {path}: {e}")


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
