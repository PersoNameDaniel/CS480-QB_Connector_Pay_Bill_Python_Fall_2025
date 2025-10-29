"""QuickBooks COM gateway helpers for pay bills."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from contextlib import contextmanager
from typing import Iterator, List

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    win32com = None  # type: ignore

from .models import BillPayment


APP_NAME = "Quickbooks Connector"  # do not chanege this


def _require_win32com() -> None:
    if win32com is None:  # pragma: no cover - exercised via tests
        raise RuntimeError("pywin32 is required to communicate with QuickBooks")


@contextmanager
def _qb_session() -> Iterator[tuple[object, object]]:
    _require_win32com()
    session = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    session.OpenConnection2("", APP_NAME, 1)
    ticket = session.BeginSession("", 0)
    try:
        yield session, ticket
    finally:
        try:
            session.EndSession(ticket)
        finally:
            session.CloseConnection()


def _send_qbxml(qbxml: str) -> ET.Element:
    with _qb_session() as (session, ticket):
        print(f"Sending QBXML:\n{qbxml}")  # Debug output
        raw_response = session.ProcessRequest(ticket, qbxml)  # type: ignore[attr-defined]
        print(f"Received response:\n{raw_response}")  # Debug output
    return _parse_response(raw_response)


def _parse_response(raw_xml: str) -> ET.Element:
    root = ET.fromstring(raw_xml)
    response = root.find(".//*[@statusCode]")
    if response is None:
        raise RuntimeError("QuickBooks response missing status information")

    status_code = int(response.get("statusCode", "0"))
    status_message = response.get("statusMessage", "")
    # Status code 1 means "no matching objects found" - this is OK for queries
    if status_code != 0 and status_code != 1:
        print(f"QuickBooks error ({status_code}): {status_message}")
        raise RuntimeError(status_message)
    return root


def fetch_bill_payments(company_file: str | None = None) -> List[BillPayment]:
    """Return bill payments (checks) with memo, date, bank account, and amount to pay."""

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <BillPaymentCheckQueryRq>\n"
        "      <IncludeLineItems>true</IncludeLineItems>\n"
        "    </BillPaymentCheckQueryRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )
    root = _send_qbxml(qbxml)

    payments: List[BillPayment] = []
    for ret in root.findall(".//BillPaymentCheckRet"):
        txn_id = (ret.findtext("TxnID") or "").strip()
        if not txn_id:
            continue

        memo = (ret.findtext("Memo") or "").strip()
        txn_date = ret.findtext("TxnDate")
        # bank_account = (ret.findtext("BankAccountRef/FullName") or "").strip()

        # Amount to Pay = sum of AppliedToTxnRet/PaymentAmount; fallback to header total
        from decimal import Decimal, InvalidOperation

        amount_to_pay_value: float = 0.0
        try:
            line_amounts = [
                Decimal((n.text or "0").strip())
                for n in ret.findall("AppliedToTxnRet/PaymentAmount")
            ]
            if line_amounts:
                amount_to_pay_value = float(sum(line_amounts))
            else:
                header_amt = (
                    ret.findtext("TotalAmount") or ret.findtext("Amount") or "0"
                ).strip()
                amount_to_pay_value = float(Decimal(header_amt))
        except (InvalidOperation, AttributeError, ValueError):
            amount_to_pay_value = 0.0

        # Build the BillPayment model as defined in models.py
        payments.append(
            BillPayment(
                id=memo or "Bill Payment",
                date=txn_date,
                amount_to_pay=amount_to_pay_value,
            )
        )
    return payments


def add_bill_payments_batch():
    raise NotImplementedError


def add_bill_payment():
    raise NotImplementedError


def _escape_xml(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


__all__ = ["fetch_bill_payments", "add_bill_payment", "add_bill_payments_batch"]
