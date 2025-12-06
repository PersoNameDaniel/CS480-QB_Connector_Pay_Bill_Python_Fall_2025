"""QuickBooks COM gateway helpers for pay bills."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from contextlib import contextmanager
from typing import Iterator, List
from datetime import date

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


# def _normalize(h: object) -> str:
#     return str(h).strip() if h is not None else ""


def _parse_qb_date(value: str | None) -> date:
    """Parse QB date or datetime to Python date. Expects 'YYYY-MM-DD' or startswith it."""
    if not value:
        raise ValueError("Missing QuickBooks date")
    s = value.strip()
    # QBXML Date is 'YYYY-MM-DD'; DateTime starts with that. Use first 10 chars safely.
    try:
        return date.fromisoformat(s[:10])
    except ValueError as e:
        raise ValueError(f"Invalid QuickBooks date: {s}") from e


def _send_qbxml(qbxml: str) -> ET.Element:
    with _qb_session() as (session, ticket):
        # print(f"Sending QBXML:\n{qbxml}")  # Debug output
        try:
            raw_response = session.ProcessRequest(ticket, qbxml)  # type: ignore[attr-defined]
            # print(f"Received response:\n{raw_response}")  # Debug output
        except Exception as e:
            print(f"ERROR during ProcessRequest: {e}")
            raise
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
        txn_date_raw = ret.findtext("TxnDate")
        try:
            txn_date = _parse_qb_date(txn_date_raw)
        except ValueError:
            continue  # skip if date is missing/invalid
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
                source="quickbooks",
                id=memo,
                date=txn_date,
                amount_to_pay=amount_to_pay_value,
            )
        )
    return payments


def add_bill_payments_batch(
    company_file: str | None, payments: List[BillPayment]
) -> List[BillPayment]:
    """Create multiple bill payments in QuickBooks in a single batch request."""

    if not payments:
        return []

    from decimal import Decimal

    requests = []
    for payment in payments:
        # Find unpaid bills for this vendor
        unpaid_bills = fetch_unpaid_bills_for_vendor(payment.vendor)

        if not unpaid_bills:
            print(
                f"Warning: No unpaid bills found for vendor: {payment.vendor}, skipping..."
            )
            continue

        # Try to find a bill that matches the payment amount exactly
        bill_txn_id = None
        bill_amount_due = 0.0

        for txn_id, amount_due in unpaid_bills:
            if abs(amount_due - payment.amount_to_pay) < 0.01:
                bill_txn_id = txn_id
                bill_amount_due = amount_due
                break

        # If no exact match, use the first bill
        if bill_txn_id is None:
            bill_txn_id, bill_amount_due = unpaid_bills[0]

        payment_amount = min(payment.amount_to_pay, bill_amount_due)

        # Build the QBXML snippet for this bill payment creation
        requests.append(
            f"    <BillPaymentCheckAddRq>\n"
            f"      <BillPaymentCheckAdd>\n"
            f"        <PayeeEntityRef>\n"
            f"          <FullName>{_escape_xml(payment.vendor)}</FullName>\n"
            f"        </PayeeEntityRef>\n"
            f"        <TxnDate>{payment.date.isoformat()}</TxnDate>\n"
            f"        <BankAccountRef>\n"
            f"          <FullName>Chase</FullName>\n"
            f"        </BankAccountRef>\n"
            f"        <IsToBePrinted>false</IsToBePrinted>\n"
            f"        <Memo>{_escape_xml(payment.id)}</Memo>\n"
            f"        <AppliedToTxnAdd>\n"
            f"          <TxnID>{_escape_xml(bill_txn_id)}</TxnID>\n"
            f"          <PaymentAmount>{Decimal(str(payment_amount)):.2f}</PaymentAmount>\n"
            f"        </AppliedToTxnAdd>\n"
            f"      </BillPaymentCheckAdd>\n"
            f"    </BillPaymentCheckAddRq>"
        )

    if not requests:
        return []

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="continueOnError">\n' + "\n".join(requests) + "\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    try:
        root = _send_qbxml(qbxml)
    except RuntimeError as exc:
        print(f"Batch add failed: {exc}")
        return []

    # Parse all responses
    added_payments: List[BillPayment] = []
    for payment_ret in root.findall(".//BillPaymentCheckRet"):
        memo = (payment_ret.findtext("Memo") or "").strip()
        vendor = (payment_ret.findtext("PayeeEntityRef/FullName") or "").strip()

        txn_date_raw = payment_ret.findtext("TxnDate")
        try:
            txn_date = _parse_qb_date(txn_date_raw) if txn_date_raw else date.today()
        except ValueError:
            txn_date = date.today()

        amount_str = payment_ret.findtext("Amount") or "0"
        try:
            amount = float(Decimal(amount_str.strip()))
        except Exception:
            amount = 0.0

        added_payments.append(
            BillPayment(
                source="",
                id=memo,
                date=txn_date,
                amount_to_pay=amount,
                vendor=vendor,
            )
        )

    return added_payments


def fetch_unpaid_bills_for_vendor(vendor_name: str) -> List[tuple[str, float]]:
    """Fetch unpaid bills for a vendor. Returns list of (TxnID, AmountDue) tuples."""

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <BillQueryRq>\n"
        f"      <EntityFilter>\n"
        f"        <FullName>{_escape_xml(vendor_name)}</FullName>\n"
        f"      </EntityFilter>\n"
        "      <PaidStatus>NotPaidOnly</PaidStatus>\n"
        "    </BillQueryRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    try:
        root = _send_qbxml(qbxml)
    except RuntimeError:
        return []

    from decimal import Decimal

    bills = []
    for bill_ret in root.findall(".//BillRet"):
        txn_id = (bill_ret.findtext("TxnID") or "").strip()
        amount_due_str = bill_ret.findtext("AmountDue") or "0"
        try:
            amount_due = float(Decimal(amount_due_str.strip()))
            if txn_id and amount_due > 0:
                print(f"Found unpaid bill: TxnID={txn_id}, AmountDue={amount_due}")
                bills.append((txn_id, amount_due))
        except Exception:
            continue

    return bills


def _escape_xml(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def read_data() -> List[BillPayment]:
    """Read bill payments from QuickBooks."""
    return fetch_bill_payments()


__all__ = ["read_data"]

if __name__ == "__main__":
    for obj in read_data():
        print(str(obj))
