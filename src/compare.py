"""
Bill Payment Comparer Module

Compares payment data between Excel and QuickBooks for existing bills.
"""

from typing import Dict, List, Any
from decimal import Decimal

# from .models import BillPayment


def normalize_amount(value: Any) -> float:
    """Convert amount to float for comparison."""
    if value is None:
        return 0.0
    if isinstance(value, Decimal):
        return float(value)
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def get_bill_id(bill: Dict[str, Any]) -> str:
    """
    Extract bill ID from bill record.
    Uses Parent-Child ID if child exists, otherwise uses Parent ID.
    """
    parent_id = bill.get("Parent ID", "").strip()
    child_id = bill.get("Child ID", "").strip()

    if child_id:
        return f"{parent_id}-{child_id}"
    return parent_id


def compare_records(
    excel_data: List[Dict[str, Any]],
    qb_data: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Compare records between Excel and QuickBooks.

    Args:
        excel_data: List of records from Excel with 'id' and 'amount' keys
        qb_data: List of records from QuickBooks with 'id' and 'amount' keys

    Returns:
        Dict containing:
            - missing_in_excel: Records in QB but not in Excel
            - missing_in_qb: Records in Excel but not in QB
            - amount_mismatches: Records with matching IDs but different amounts
    """
    # Index records by id (filter out None keys)
    excel_by_id: Dict[Any, Dict[str, Any]] = {
        rec.get("id"): rec for rec in excel_data if rec.get("id") is not None
    }
    qb_by_id: Dict[Any, Dict[str, Any]] = {
        rec.get("id"): rec for rec in qb_data if rec.get("id") is not None
    }

    results: Dict[str, Any] = {
        "missing_in_excel": [],
        "missing_in_qb": [],
        "amount_mismatches": [],
    }

    # Find records missing in Excel (in QB but not Excel)
    for qb_id, qb_rec in qb_by_id.items():
        if qb_id not in excel_by_id:
            results["missing_in_excel"].append(qb_rec)

    # Find records missing in QB (in Excel but not QB)
    for excel_id, excel_rec in excel_by_id.items():
        if excel_id not in qb_by_id:
            results["missing_in_qb"].append(excel_rec)

    # Find amount mismatches for matching IDs
    for rec_id in excel_by_id.keys() & qb_by_id.keys():
        excel_amount = normalize_amount(excel_by_id[rec_id].get("amount"))
        qb_amount = normalize_amount(qb_by_id[rec_id].get("amount"))

        if abs(excel_amount - qb_amount) > 0.01:  # tolerance of 1 cent
            results["amount_mismatches"].append(
                {
                    "id": rec_id,
                    "excel_amount": excel_amount,
                    "qb_amount": qb_amount,
                }
            )

    return results


def main():
    """Main function."""
    print("Bill Payment Comparer Module loaded successfully.")


if __name__ == "__main__":
    main()


__all__ = ["compare_records"]
