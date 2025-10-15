"""Excel extraction helpers for account debit worksheets."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import load_workbook


def _default_company_workbook() -> Path:
    # company_data.xlsx expected in project root (one level above src/)
    return Path(__file__).resolve().parents[1] / "company_data.xlsx"


def _read_sheet_as_dicts(
    workbook_path: Path, sheet_name: str
) -> List[Dict[str, object]]:
    """Read the given worksheet and return a list of dicts mapping header -> value.

    Each returned row dict will include:
      - the original columns mapped by header (normalized to str)
      - 'source': 'excel'
      - '__sheet__': the worksheet name

    Raises FileNotFoundError if workbook is missing and ValueError if the sheet is not found.
    """
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(filename=workbook_path, read_only=True, data_only=True)
    try:
        try:
            ws = wb[sheet_name]
        except KeyError as exc:
            raise ValueError(f"Worksheet '{sheet_name}' not found in workbook") from exc

        rows = ws.iter_rows(values_only=True)
        header_row = next(rows, None)
        if header_row is None:
            return []

        # Normalize headers; empty headers get a generated name
        headers = [
            (
                str(h).strip()
                if h is not None and str(h).strip() != ""
                else f"column_{i}"
            )
            for i, h in enumerate(header_row)
        ]

        results: List[Dict[str, object]] = []
        for row in rows:
            row_dict: Dict[str, object] = {}
            for i, header in enumerate(headers):
                value = row[i] if i < len(row) else None
                row_dict[header] = value
            row_dict["source"] = "excel"
            row_dict["__sheet__"] = sheet_name
            results.append(row_dict)

        return results
    finally:
        wb.close()


def extract_account_debit_vendor(
    workbook_path: Optional[Path] = None,
) -> List[Dict[str, object]]:
    """Return rows from the 'account debit vendor' worksheet as list of dicts.

    If workbook_path is None, company_data.xlsx at the project root is used.
    """
    path = workbook_path or _default_company_workbook()
    return _read_sheet_as_dicts(path, "account debit vendor")


def extract_account_debit_nonvendor(
    workbook_path: Optional[Path] = None,
) -> List[Dict[str, object]]:
    """Return rows from the 'account debit nonvendor' worksheet as list of dicts.

    If workbook_path is None, company_data.xlsx at the project root is used.
    """
    path = workbook_path or _default_company_workbook()
    return _read_sheet_as_dicts(path, "account debit nonvendor")


__all__ = [
    "extract_account_debit_vendor",
    "extract_account_debit_nonvendor",
]
