from datetime import date
from pathlib import Path

from src.models import BillPayment
from src.compare import compare_bill_payments
import src.reporting as reporting


def _payment_to_dict(p: BillPayment) -> dict:
    return {
        "record_id": p.id,
        "amount": float(p.amount_to_pay),
        "date": p.date.isoformat(),
        "vendor": p.vendor,
    }


def _conflict_to_dict(c) -> dict:
    return {
        "record_id": c.record_id,
        "reason": c.reason,
        "excel_amount": c.excel_amount,
        "qb_amount": c.qb_amount,
        "excel_date": c.excel_date,
        "qb_date": c.qb_date,
        "excel_vendor": c.excel_vendor,
        "qb_vendor": c.qb_vendor,
    }


def test_compare_and_report(tmp_path: Path) -> None:
    # Fake Excel + QB data just to test compare logic
    excel = [
        BillPayment(id="100", date=date(2024, 1, 1), amount_to_pay=200.0, vendor="A"),
        BillPayment(id="200", date=date(2024, 2, 2), amount_to_pay=500.0, vendor="B"),
    ]

    qb = [
        BillPayment(
            id="100", date=date(2024, 1, 1), amount_to_pay=200.0, vendor="A"
        ),  # same
        BillPayment(
            id="300", date=date(2024, 3, 3), amount_to_pay=900.0, vendor="C"
        ),  # QB-only
    ]

    result = compare_bill_payments(excel, qb)

    # Basic sanity-checks on comparison logic
    assert [p.id for p in result.excel_only] == ["200"]
    assert [p.id for p in result.qb_only] == ["300"]
    assert [c.record_id for c in result.conflicts] == []

    payload = {
        "status": "success",
        "generated_at": reporting.iso_timestamp(),
        "added_records": [_payment_to_dict(p) for p in result.excel_only],
        "conflicts": (
            [_conflict_to_dict(c) for c in result.conflicts]
            + [
                {
                    "record_id": p.id,
                    "reason": "payment_only_in_quickbooks",
                    "excel_amount": None,
                    "qb_amount": float(p.amount_to_pay),
                    "excel_date": None,
                    "qb_date": p.date.isoformat(),
                    "excel_vendor": None,
                    "qb_vendor": p.vendor,
                }
                for p in result.qb_only
            ]
        ),
        "same_records": 0,
        "error": None,
    }

    out_path = tmp_path / "test_compare_output.json"
    reporting.write_report(payload, out_path)

    # Check that report file was actually written
    assert out_path.exists()
    assert out_path.stat().st_size > 0
