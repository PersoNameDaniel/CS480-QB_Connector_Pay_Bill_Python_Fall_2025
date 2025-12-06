"""
reporting.py

Utilities for writing JSON reports and generating timestamps.
"""

from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict


def write_report(payload: Dict[str, Any], output_path: Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)

    return output_path


def iso_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat()


__all__ = ["write_report", "iso_timestamp"]
