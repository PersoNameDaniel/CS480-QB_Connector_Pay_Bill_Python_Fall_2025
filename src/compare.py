from typing import List, Dict, Any
from decimal import Decimal


def normalize_id(value: Any) -> int | str:
    """Normalize ID to int if possible, otherwise string."""
    try:
        return int(value)
    except (ValueError, TypeError):
        return str(value).strip() if value is not None else ""


def compare_records(
    excel_data: List[Dict[str, Any]],
    qb_data: List[Dict[str, Any]],
    amount_tolerance: float = 0.01,
) -> Dict[str, Any]:
    """Compare financial records from Excel and QuickBooks."""

    # 🔹 Normalize Excel data into standard form
    normalized_excel: List[Dict[str, Any]] = []
    for record in excel_data:
        record_id = (
            record.get("Parent ID")
            or record.get("Child ID")
            or record.get("Invoice Num")
            or record.get("id")
        )
        amount = (
            record.get("Invoice Amount")
            or record.get("Check Amount")
            or record.get("AP")
            or record.get("amount")
        )
        if record_id is not None and amount is not None:
            normalized_excel.append(
                {"id": normalize_id(record_id), "amount": float(amount)}
            )

    # 🔹 Normalize QuickBooks data (if any)
    normalized_qb: List[Dict[str, Any]] = []
    for record in qb_data:
        record_id = (
            record.get("id")
            or record.get("TxnID")
            or record.get("bill")
            or record.get("Name")
        )
        amount = (
            record.get("amount")
            or record.get("amount_to_pay")
            or record.get("TotalAmount")
        )
        if record_id is not None and amount is not None:
            normalized_qb.append(
                {"id": normalize_id(record_id), "amount": float(amount)}
            )

    # 🔹 Validate
    for record in normalized_excel + normalized_qb:
        if "id" not in record or "amount" not in record:
            raise KeyError(f"Record missing required fields: {record}")
        if not isinstance(record["amount"], (int, float, Decimal)):
            raise TypeError(f"Amount must be numeric: {record}")

    discrepancies: Dict[str, Any] = {
        "missing_in_excel": [],
        "missing_in_qb": [],
        "amount_mismatches": [],
    }

    excel_dict = {normalize_id(r["id"]): r for r in normalized_excel}
    qb_dict = {normalize_id(r["id"]): r for r in normalized_qb}

    # 🔹 Missing in Excel
    for qb_id, qb_record in qb_dict.items():
        if qb_id not in excel_dict:
            discrepancies["missing_in_excel"].append(qb_record)

    # 🔹 Missing in QuickBooks
    for excel_id, excel_record in excel_dict.items():
        if excel_id not in qb_dict:
            discrepancies["missing_in_qb"].append(excel_record)

    # 🔹 Amount mismatches
    for record_id in set(excel_dict.keys()) & set(qb_dict.keys()):
        excel_amt = float(excel_dict[record_id]["amount"])
        qb_amt = float(qb_dict[record_id]["amount"])
        if abs(excel_amt - qb_amt) > amount_tolerance:
            discrepancies["amount_mismatches"].append(
                {"id": record_id, "excel_amount": excel_amt, "qb_amount": qb_amt}
            )

    # 🔹 Sort results
    for key in discrepancies:
        discrepancies[key].sort(key=lambda x: x["id"] if isinstance(x, dict) else x)

    return discrepancies


__all__ = ["compare_records"]
