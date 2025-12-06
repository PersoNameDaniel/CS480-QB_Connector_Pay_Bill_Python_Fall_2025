"""
High-level orchestration for the pay bills CLI.

This module:

1. Reads bill payments from Excel (vendor / nonvendor sheets).
2. Reads existing bill payments from QuickBooks.
3. Compares both datasets at the PAYMENT level.
4. Adds missing Excel-only payments to QuickBooks (batch).
5. Writes a JSON report summarising the outcome.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List

from . import excel_reader, qb_gateway, compare
from .models import BillPayment, Conflict, ComparisonReport
from .reporting import iso_timestamp, write_report

DEFAULT_REPORT_NAME = "pay_bills_report.json"


def _payment_to_dict(payment: BillPayment) -> Dict[str, object]:
    """Convert a BillPayment object into a serialisable dict for added_records."""
    return {
        "record_id": payment.id,
        "amount": float(payment.amount_to_pay),
        "date": payment.date
        if isinstance(payment.date, str)
        else payment.date.isoformat(),
        "vendor": payment.vendor,
    }


def _conflict_to_dict(conflict: Conflict) -> Dict[str, object]:
    """Convert a Conflict object into a serialisable dict for the report."""
    return {
        "record_id": conflict.record_id,
        "reason": conflict.reason,
        "excel_amount": conflict.excel_amount,
        "qb_amount": conflict.qb_amount,
        "excel_date": conflict.excel_date,
        "qb_date": conflict.qb_date,
        "excel_vendor": conflict.excel_vendor,
        "qb_vendor": conflict.qb_vendor,
    }


def _qb_only_conflict(payment: BillPayment) -> Dict[str, object]:
    """Create a conflict payload for payments present only in QuickBooks."""
    return {
        "record_id": payment.id,
        "reason": "payment_only_in_quickbooks",
        "excel_amount": None,
        "qb_amount": float(payment.amount_to_pay),
        "excel_date": None,
        "qb_date": payment.date.isoformat(),
        "excel_vendor": None,
        "qb_vendor": payment.vendor,
    }


def _count_matching_payments(
    excel_payments: List[BillPayment], qb_payments: List[BillPayment]
) -> int:
    """
    Return the number of payments that exist in both sources with identical data.

    This is where we implement:
    - If payment exists in Excel AND QB AND data is the same -> same_records++.
    """
    excel_by_id = {p.id: p for p in excel_payments}
    qb_by_id = {p.id: p for p in qb_payments}

    matches = 0
    for record_id in excel_by_id.keys() & qb_by_id.keys():
        e = excel_by_id[record_id]
        q = qb_by_id[record_id]
        same_amount = abs(float(e.amount_to_pay) - float(q.amount_to_pay)) < 0.01
        same_date = e.date == q.date
        # same_vendor = (e.vendor or "") == (q.vendor or "")
        if same_amount and same_date:
            matches += 1
    return matches


def run_pay_bills(
    company_file_path: str,
    workbook_path: str,
    *,
    sheet_type: str = "vendor",  # "vendor" or "nonvendor"
    output_path: str | None = None,
) -> Path:
    """
    Contract entry point for synchronising bill payments.

    Args:
        company_file_path:
            Path to the QuickBooks company file. Use "" to reuse the currently
            open company file.
        workbook_path:
            Path to the Excel workbook containing the account debit sheet.
        sheet_type:
            "vendor" or "nonvendor" to choose correct extractor.
        output_path:
            Optional JSON output path. Defaults to pay_bills_report.json.

    Returns:
        Path to the generated JSON report.
    """

    report_path = Path(output_path) if output_path else Path(DEFAULT_REPORT_NAME)

    report_payload: Dict[str, object] = {
        "status": "success",
        "generated_at": iso_timestamp(),
        "added_records": [],
        "conflicts": [],
        "same_records": 0,
        "error": None,
    }

    try:
        # 1) Extract bill payments from Excel
        wb_path = Path(workbook_path)
        if sheet_type.lower() == "vendor":
            excel_payments = excel_reader.extract_account_debit_vendor(wb_path)
        else:
            excel_payments = excel_reader.extract_account_debit_nonvendor(wb_path)
        print(excel_payments)

        # 2) Fetch existing bill payments from QuickBooks
        qb_payments = qb_gateway.read_data()
        print("QB Payments:")
        print(qb_payments)

        # 3) Compare the two sources at PAYMENT level
        comparison: ComparisonReport = compare.compare_bill_payments(
            excel_payments, qb_payments
        )

        # 4) Add Excel-only payments to QuickBooks (batch)
        #    IMPORTANT: only the payments that QB actually accepts and returns
        #    will show up in added_records in the JSON.
        added_payments: List[BillPayment] = qb_gateway.add_bill_payments_batch(
            company_file_path, comparison.excel_only
        )

        # 5) Build conflicts list:
        #    - data_conflict from comparison.conflicts
        #    - payment_only_in_quickbooks from comparison.qb_only
        print()
        conflicts: List[Dict[str, object]] = []
        conflicts.extend(_conflict_to_dict(c) for c in comparison.conflicts)
        conflicts.extend(_qb_only_conflict(p) for p in comparison.qb_only)

        # 6) Populate report payload
        report_payload["added_records"] = [_payment_to_dict(p) for p in added_payments]
        report_payload["conflicts"] = conflicts
        report_payload["same_records"] = _count_matching_payments(
            excel_payments, qb_payments
        )

    except Exception as exc:  # pragma: no cover
        report_payload["status"] = "error"
        report_payload["error"] = str(exc)

    write_report(report_payload, report_path)
    return report_path


__all__ = ["run_pay_bills", "DEFAULT_REPORT_NAME"]
