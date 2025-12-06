"""
Bill Payment Comparer Module

Compares payment data between Excel and QuickBooks for existing bills.
"""

from typing import Dict, List, Any
from decimal import Decimal

from .models import BillPayment


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


def compare_bill_payments(
    excel_bills: List[Dict[str, Any]],
    qb_bills: List[Dict[str, Any]],
    qb_payments: List[Dict[str, Any]],
    amount_tolerance: float = 0.01,
) -> Dict[str, Any]:
    """
    Compare bill payment data between Excel and QuickBooks.
    
    Args:
        excel_bills: List of bills from Excel with payment data
        qb_bills: List of bills from QuickBooks (with TxnIDs)
        qb_payments: List of existing payments from QuickBooks
        amount_tolerance: Maximum acceptable difference for amount comparisons
        
    Returns:
        Dict containing:
            - same_records: List of matching BillPayment records
            - conflicts: List of payment discrepancies
            - payments_to_add: List of BillPayment objects to add to QB
    """
    
    # Index Excel bills by ID for O(1) lookup
    excel_by_id: Dict[str, Dict[str, Any]] = {
        get_bill_id(bill): bill 
        for bill in excel_bills 
        if get_bill_id(bill)
    }
    
    # Index QB bills by TxnID for O(1) lookup
    qb_by_txn: Dict[str, Dict[str, Any]] = {
        bill.get("TxnID"): bill 
        for bill in qb_bills 
        if bill.get("TxnID")
    }
    
    # Index QB payments by TxnID
    qb_payments_by_txn: Dict[str, List[Dict[str, Any]]] = {}
    for payment in qb_payments:
        txn_id = payment.get("TxnID")
        if txn_id:
            if txn_id not in qb_payments_by_txn:
                qb_payments_by_txn[txn_id] = []
            qb_payments_by_txn[txn_id].append(payment)
    
    # Results
    results = {
        "same_records": [],
        "conflicts": [],
        "payments_to_add": [],
    }
    
    # Compare bills that exist in both systems
    for bill_id in excel_by_id.keys() & qb_by_txn.keys():
        excel_bill = excel_by_id[bill_id]
        qb_bill = qb_by_txn[bill_id]
        txn_id = qb_bill.get("TxnID")
        
        # Extract payment info from Excel
        excel_amount = normalize_amount(
            excel_bill.get("Payment Amount") or excel_bill.get("Check Amount")
        )
        excel_date = excel_bill.get("Payment Date") or excel_bill.get("Check Date")
        
        # Skip if no payment in Excel
        if excel_amount == 0.0:
            continue
        
        # Get existing QB payments for this bill
        qb_bill_payments = qb_payments_by_txn.get(txn_id, [])
        
        # If no QB payment exists, add it
        if not qb_bill_payments:
            payment = BillPayment(
                id=bill_id,
                date=excel_date,
                amount_to_pay=excel_amount,
                vendor=qb_bill.get("vendor_name", ""),
            )
            results["payments_to_add"].append(payment)
            continue
        
        # Compare payment data
        payment_matched = False
        for qb_payment in qb_bill_payments:
            qb_amount = normalize_amount(
                qb_payment.get("amount") or qb_payment.get("TotalAmount")
            )
            qb_date = qb_payment.get("payment_date") or qb_payment.get("TxnDate")
            
            amount_match = abs(excel_amount - qb_amount) <= amount_tolerance
            date_match = str(excel_date) == str(qb_date)
            
            if amount_match and date_match:
                payment = BillPayment(
                    id=bill_id,
                    date=excel_date,
                    amount_to_pay=excel_amount,
                    vendor=qb_bill.get("vendor_name", ""),
                )
                results["same_records"].append(payment)
                payment_matched = True
                break
        
        # Record conflict if no match found
        if not payment_matched:
            results["conflicts"].append({
                "bill_id": bill_id,
                "txn_id": txn_id,
                "reason": "payment_mismatch",
                "excel_payment": {
                    "amount": excel_amount,
                    "date": excel_date,
                },
                "qb_payments": [
                    {
                        "amount": normalize_amount(p.get("amount") or p.get("TotalAmount")),
                        "date": p.get("payment_date") or p.get("TxnDate"),
                    }
                    for p in qb_bill_payments
                ],
            })
    
    # Check for QB payments with no Excel counterpart (conflicts)
    for txn_id, qb_bill in qb_by_txn.items():
        # Get the Excel bill ID equivalent
        bill_id = qb_bill.get("RefNumber") or qb_bill.get("parent_id")
        
        # Skip if already processed (bill exists in Excel)
        if bill_id in excel_by_id:
            continue
        
        # Payment in QB but not in Excel = conflict
        if txn_id in qb_payments_by_txn:
            results["conflicts"].append({
                "bill_id": bill_id,
                "txn_id": txn_id,
                "reason": "payment_in_qb_only",
                "excel_payment": None,
                "qb_payments": [
                    {
                        "amount": normalize_amount(p.get("amount") or p.get("TotalAmount")),
                        "date": p.get("payment_date") or p.get("TxnDate"),
                    }
                    for p in qb_payments_by_txn[txn_id]
                ],
            })
    
    return results


def main():
    """Main function."""
    print("Bill Payment Comparer Module loaded successfully.")


if __name__ == "__main__":
    main()


__all__ = ["compare_bill_payments"]