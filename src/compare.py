"""
Bill Payment Comparer Module

Compares payment data between Excel and QuickBooks for existing bills.
"""

from typing import Dict, List, Any
from decimal import Decimal


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


def compare_records(
    excel_data: List[Dict[str, Any]],
    qb_data: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Compare payment records between Excel and QuickBooks.

    Args:
        excel_data: List of payment records from Excel with 'id' and 'amount_to_pay' keys
        qb_data: List of payment records from QuickBooks with 'id' and 'amount_to_pay' keys

    Returns:
        Dict containing:
            - same_records: Count of payments that match between Excel and QB
            - conflicts: Payments with mismatches or only in QB
            - to_add_to_qb: Payments only in Excel (to be added)
    """
    # Index records by id (filter out None keys)
    excel_by_id: Dict[Any, Dict[str, Any]] = {
        rec.get("id"): rec for rec in excel_data if rec.get("id") is not None
    }
    qb_by_id: Dict[Any, Dict[str, Any]] = {
        rec.get("id"): rec for rec in qb_data if rec.get("id") is not None
    }

    results: Dict[str, Any] = {
        "same_records_count": 0,
        "conflicts": [],
        "added_bill_payments": [],
    }

    # Find payments only in QB → conflicts
    for qb_id, qb_rec in qb_by_id.items():
        if qb_id not in excel_by_id:
            results["conflicts"].append(
                {
                    "type": "only_in_qb",
                    "id": qb_id,
                    "qb_record": qb_rec,
                }
            )

    # Find payments only in Excel → add to QB
    for excel_id, excel_rec in excel_by_id.items():
        if excel_id not in qb_by_id:
            results["added_bill_payments"].append(excel_rec)

    # Compare payments that exist in both
    for rec_id in excel_by_id.keys() & qb_by_id.keys():
        excel_rec = excel_by_id[rec_id]
        qb_rec = qb_by_id[rec_id]

        excel_amount = normalize_amount(excel_rec.get("amount_to_pay"))
        qb_amount = normalize_amount(qb_rec.get("amount_to_pay"))

        if abs(excel_amount - qb_amount) > 0.01:  # tolerance of 1 cent
            results["conflicts"].append(
                {
                    "type": "data_mismatch",
                    "excel_id": rec_id,
                    "qb_id": qb_rec.get("id"),
                    "excel_date": excel_rec.get("date"),
                    "qb_date": qb_rec.get("date"),
                    "excel_amount": excel_amount,
                    "qb_amount": qb_amount,
                    "excel_vendor": excel_rec.get("vendor"),
                    "qb_vendor": qb_rec.get("vendor"),
                }
            )
        else:
            # Payments match - count them
            results["same_records_count"] += 1

    return results


def main():
    """Main function."""
    print("Bill Payment Comparer Module loaded successfully.")


if __name__ == "__main__":
    main()


__all__ = ["compare_records"]
