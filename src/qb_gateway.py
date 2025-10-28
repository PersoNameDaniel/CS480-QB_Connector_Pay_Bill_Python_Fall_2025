"""QuickBooks COM gateway helpers for pay bills."""

from __future__ import annotations
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from typing import Iterator, List, Dict, Any

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    win32com = None  # type: ignore

from .models import BillPayment

APP_NAME = "Quickbooks Connector"  # do not change this


# ---------------------------------------------------------------------------
# Session and Common Helpers
# ---------------------------------------------------------------------------


def _require_win32com() -> None:
    if win32com is None:
        raise RuntimeError("pywin32 is required to communicate with QuickBooks")


@contextmanager
def _qb_session() -> Iterator[tuple[object, object]]:
    """Context manager for managing QuickBooks COM session lifecycle."""
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


def _escape_xml(value: str) -> str:
    """Escape special XML characters."""
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


# ---------------------------------------------------------------------------
# Send/Parse QBXML
# ---------------------------------------------------------------------------


def _send_qbxml(qbxml: str) -> ET.Element:
    """Send QBXML to QuickBooks and return parsed XML response."""
    with _qb_session() as (session, ticket):
        print(f"Sending QBXML:\n{qbxml}")
        raw_response = session.ProcessRequest(ticket, qbxml)  # type: ignore[attr-defined]
        print(f"Received response:\n{raw_response}")
    return _parse_response(raw_response)


def _parse_response(raw_xml: str) -> ET.Element:
    """Parse XML response and raise for QuickBooks error codes."""
    root = ET.fromstring(raw_xml)
    response = root.find(".//*[@statusCode]")
    if response is None:
        raise RuntimeError("QuickBooks response missing status information")

    status_code = int(response.get("statusCode", "0"))
    status_message = response.get("statusMessage", "")
    if status_code not in (0, 1):
        raise RuntimeError(f"QuickBooks error ({status_code}): {status_message}")
    return root


# ---------------------------------------------------------------------------
# Query Bill Payments
# ---------------------------------------------------------------------------


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
        txn_date = (ret.findtext("TxnDate") or "").strip()
        bank_account = (ret.findtext("BankAccountRef/FullName") or "").strip()

        from decimal import Decimal, InvalidOperation

        try:
            payment_nodes = ret.findall("AppliedToTxnRet/Amount") or ret.findall(
                "AppliedToTxnRet/PaymentAmount"
            )
            line_amounts = [Decimal((n.text or "0").strip()) for n in payment_nodes]
            amount_to_pay_value = (
                float(sum(line_amounts))
                if line_amounts
                else float(Decimal((ret.findtext("TotalAmount") or "0").strip()))
            )
        except (InvalidOperation, AttributeError, ValueError):
            amount_to_pay_value = 0.0

        payments.append(
            BillPayment(
                bill=memo or "Bill Payment",
                date=txn_date,
                bank_account=bank_account,
                amount_to_pay=amount_to_pay_value,
            )
        )
    return payments


# ---------------------------------------------------------------------------
# Add Bill Payment
# ---------------------------------------------------------------------------


def _build_bill_payment_xml(record: Dict[str, Any]) -> str:
    """Build BillPaymentCheckAddRq XML request."""
    vendor_name = _escape_xml(record.get("Supplier Name", "Unknown Vendor"))
    bank_account = _escape_xml(record.get("Bank Account", "Default Checking"))
    memo = _escape_xml(record.get("Comment", ""))
    payment_date = str(record.get("Bank Date", ""))
    txn_id = str(record.get("Parent ID", "")) or str(record.get("Child ID", ""))
    amount = str(record.get("Check Amount", record.get("Invoice Amount", 0)))

    qbxml = f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="16.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <BillPaymentCheckAddRq>
      <BillPaymentCheckAdd>
        <PayeeEntityRef>
          <FullName>{vendor_name}</FullName>
        </PayeeEntityRef>
        <BankAccountRef>
          <FullName>{bank_account}</FullName>
        </BankAccountRef>
        <TxnDate>{payment_date}</TxnDate>
        <Memo>{memo}</Memo>
        <AppliedToTxnAdd>
          <TxnID>{txn_id}</TxnID>
          <PaymentAmount>{amount}</PaymentAmount>
        </AppliedToTxnAdd>
      </BillPaymentCheckAdd>
    </BillPaymentCheckAddRq>
  </QBXMLMsgsRq>
</QBXML>"""
    return qbxml


def add_bill_payment(record: Dict[str, Any]) -> None:
    """Add a single bill payment record to QuickBooks."""
    qbxml = _build_bill_payment_xml(record)
    print(f"Attempting to add BillPayment for Vendor: {record.get('Supplier Name')}")
    _send_qbxml(qbxml)
    print(f"Successfully added payment for {record.get('Supplier Name')}")


def add_bill_payments_batch(records: List[Dict[str, Any]]) -> None:
    """Add multiple Excel-only bill payments to QuickBooks."""
    if not records:
        print("No new Excel-only payments to add.")
        return
    print(f"Adding {len(records)} new bill payments to QuickBooks...")
    for rec in records:
        try:
            add_bill_payment(rec)
        except Exception as e:
            print(f"Failed to add {rec.get('Supplier Name', 'Unknown')}: {e}")


__all__ = [
    "fetch_bill_payments",
    "add_bill_payment",
    "add_bill_payments_batch",
]
