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

from models import BillPayment


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
        raw_response = session.ProcessRequest(ticket, qbxml)  # type: ignore[attr-defined]
        # print(f"Received response:\n{raw_response}")  # Debug output
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
        return []  # Nothing to add; return early

    from decimal import Decimal

    # Build the QBXML with multiple BillPaymentCheckAddRq entries
    requests = []  # Collect individual add requests to embed in one batch
    for payment in payments:
        # Build the QBXML snippet for this bill payment creation
        requests.append(
            f"    <BillPaymentCheckAddRq>\n"
            f"      <BillPaymentCheckAdd>\n"
            f"        <PayeeEntityRef>\n"
            f"          <FullName>{_escape_xml(payment.id)}</FullName>\n"
            f"        </PayeeEntityRef>\n"
            f"        <TxnDate>{payment.date.isoformat()}</TxnDate>\n"
            f"        <Memo>{_escape_xml(payment.id)}</Memo>\n"
            f"        <Amount>{Decimal(str(payment.amount_to_pay))}</Amount>\n"
            f"      </BillPaymentCheckAdd>\n"
            f"    </BillPaymentCheckAddRq>"
        )

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="continueOnError">\n' + "\n".join(requests) + "\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )  # Batch request enabling partial success on errors

    try:
        root = _send_qbxml(qbxml)  # Submit the batch to QuickBooks
    except RuntimeError as exc:
        # If the entire batch fails, return empty list
        print(f"Batch add failed: {exc}")
        return []

    # Parse all responses
    added_payments: List[BillPayment] = []  # Payments confirmed/returned by QuickBooks
    for payment_ret in root.findall(".//BillPaymentCheckRet"):
        memo = (payment_ret.findtext("Memo") or "").strip()
        if not memo:
            continue

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

        added_payments.append(BillPayment(id=memo, date=txn_date, amount_to_pay=amount))

    return added_payments  # Return all payments that were added/acknowledged


def add_bill_payment(company_file: str | None, payment: BillPayment) -> BillPayment:
    """Create a single bill payment check in QuickBooks and return the stored record."""

    from decimal import Decimal

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <BillPaymentCheckAddRq>\n"
        "      <BillPaymentCheckAdd>\n"
        f"        <PayeeEntityRef>\n"
        f"          <FullName>{_escape_xml(payment.id)}</FullName>\n"
        f"        </PayeeEntityRef>\n"
        f"        <TxnDate>{payment.date.isoformat()}</TxnDate>\n"
        f"        <Memo>{_escape_xml(payment.id)}</Memo>\n"
        f"        <Amount>{Decimal(str(payment.amount_to_pay))}</Amount>\n"
        "      </BillPaymentCheckAdd>\n"
        "    </BillPaymentCheckAddRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    try:
        root = _send_qbxml(qbxml)
    except RuntimeError as exc:
        if "already exists" in str(exc).lower():
            return payment
        raise

    payment_ret = root.find(".//BillPaymentCheckRet")
    if payment_ret is None:
        return payment

    # txn_id = (payment_ret.findtext("TxnID") or "").strip()
    memo = (payment_ret.findtext("Memo") or payment.id).strip()
    txn_date_raw = payment_ret.findtext("TxnDate")

    try:
        txn_date = _parse_qb_date(txn_date_raw) if txn_date_raw else payment.date
    except ValueError:
        txn_date = payment.date

    amount_str = payment_ret.findtext("Amount") or str(payment.amount_to_pay)
    try:
        amount = float(Decimal(amount_str.strip()))
    except Exception:
        amount = payment.amount_to_pay

    return BillPayment(
        id=memo,
        date=txn_date,
        amount_to_pay=amount,
    )


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
