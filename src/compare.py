"""
Comparison logic for bill payments.

Implements the rules:

- If bill exists in QB and not in Excel      -> ignore at bill level, but payments
                                              only in QB are represented in qb_only.
- If bill exists in Excel and not in QB      -> this "never happens" (we just treat
                                              them as excel_only payments).
- If bill exists in QB and Excel but with different bill data -> ignore bill metadata;
  we only compare payment data (id, date, amount, vendor).

If bill exists in QB and Excel:

- If payment data is same                    -> later counted as same_records.
- If payment data is different               -> recorded as Conflict (reason=data_conflict).
- If payment data only exist in Excel        -> included in excel_only (to add to QB).
- If payment data only exist in QB           -> included in qb_only (for reporting).
"""

from __future__ import annotations

from typing import List
from .models import BillPayment, Conflict, ComparisonReport


def _same_payment(e: BillPayment, q: BillPayment) -> bool:
    """Return True if Excel and QB payment data are considered identical."""
    same_amount = abs(float(e.amount_to_pay) - float(q.amount_to_pay)) < 0.01
    same_date = e.date == q.date
    # same_vendor = (e.vendor or "") == (q.vendor or "")
    return same_amount and same_date


def compare_bill_payments(
    excel_payments: List[BillPayment],
    qb_payments: List[BillPayment],
) -> ComparisonReport:
    """Compare Excel and QuickBooks payments and return a ComparisonReport."""

    # Index by record_id (assuming one payment per id per source)
    excel_by_id = {p.id: p for p in excel_payments}
    qb_by_id = {p.id: p for p in qb_payments}
    print("qb_by_id:", qb_by_id)

    excel_ids = set(excel_by_id.keys())
    qb_ids = set(qb_by_id.keys())

    common_ids = excel_ids & qb_ids
    excel_only_ids = excel_ids - qb_ids
    qb_only_ids = qb_ids - excel_ids

    report = ComparisonReport()

    # 1) Excel-only payments: candidates to be added to QuickBooks
    for record_id in sorted(excel_only_ids):
        p = excel_by_id[record_id]
        p.source = "excel"
        report.excel_only.append(p)

    # 2) QB-only payments: present in QuickBooks only
    for record_id in sorted(qb_only_ids):
        p = qb_by_id[record_id]
        p.source = "quickbooks"
        report.qb_only.append(p)

    # 3) Common IDs: compare payment data
    for record_id in sorted(common_ids):
        e = excel_by_id[record_id]
        q = qb_by_id[record_id]
        e.source = "excel"
        q.source = "quickbooks"

        if _same_payment(e, q):
            # No conflict; will be counted as same_records later
            continue

        # Payment data mismatch -> data_conflict
        report.conflicts.append(
            Conflict(
                record_id=record_id,
                reason="data_conflict",
                excel_amount=float(e.amount_to_pay),
                qb_amount=float(q.amount_to_pay),
                excel_date=e.date.isoformat(),
                qb_date=q.date.isoformat(),
                excel_vendor=e.vendor,
                qb_vendor=q.vendor,
            )
        )

    return report


__all__ = ["compare_bill_payments"]
