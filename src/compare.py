from __future__ import annotations
from typing import List, Dict, Any
from decimal import Decimal


def compare_records(
    excel_data: List[Dict[str, Any]],
    qb_data: List[Dict[str, Any]],
    amount_tolerance: float = 0.01,
) -> Dict[str, Any]:
    """Compare financial records from Excel and QuickBooks.

    Args:
        excel_data: A list of dictionaries representing financial records from Excel.
            Each record must have 'id' and 'amount' fields.
        qb_data: A list of dictionaries representing financial records from QuickBooks.
            Each record must have 'id' and 'amount' fields.
        amount_tolerance: Maximum allowed difference between amounts (default: 0.01)

    Returns:
        A dictionary containing:
            - missing_in_excel: Records present in QB but not Excel
            - missing_in_qb: Records present in Excel but not QB
            - amount_mismatches: Records where amounts differ

    Raises:
        KeyError: If records are missing required fields
        TypeError: If amount values are not numeric
    """
    # Validate input records
    for record in excel_data + qb_data:
        if not isinstance(record, dict):
            raise TypeError(f"Invalid record format: {record}")
        if "id" not in record or "amount" not in record:
            raise KeyError(f"Record missing required fields: {record}")
        if not isinstance(record["amount"], (int, float, Decimal)):
            raise TypeError(f"Amount must be numeric: {record}")

    discrepancies: Dict[str, Any] = {
        "missing_in_excel": [],
        "missing_in_qb": [],
        "amount_mismatches": [],
    }

    # Create lookup dictionaries
    excel_dict = {record["id"]: record for record in excel_data}
    qb_dict = {record["id"]: record for record in qb_data}

    # Check for records missing in Excel
    for qb_id, qb_record in qb_dict.items():
        if qb_id not in excel_dict:
            discrepancies["missing_in_excel"].append(qb_record)

    # Check for records missing in QuickBooks
    for excel_id, excel_record in excel_dict.items():
        if excel_id not in qb_dict:
            discrepancies["missing_in_qb"].append(excel_record)

    # Check for amount mismatches
    for record_id in set(excel_dict.keys()) & set(qb_dict.keys()):
        excel_amount = float(excel_dict[record_id]["amount"])
        qb_amount = float(qb_dict[record_id]["amount"])
        if abs(excel_amount - qb_amount) > amount_tolerance:
            discrepancies["amount_mismatches"].append(
                {"id": record_id, "excel_amount": excel_amount, "qb_amount": qb_amount}
            )

    # Sort results for consistent ordering
    for key in discrepancies:
        discrepancies[key].sort(key=lambda x: x["id"] if isinstance(x, dict) else x)

    return discrepancies
