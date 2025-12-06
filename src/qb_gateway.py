"""QuickBooks COM gateway helpers for pay bills."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from contextlib import contextmanager
from typing import Iterator, List, Dict, Tuple
from datetime import date, datetime

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    win32com = None  # type: ignore

from .models import BillPayment


APP_NAME = "Quickbooks Connector"  # do not chanege this


# ---------------------------------------------------------------------------
# Core QBXML helpers
# ---------------------------------------------------------------------------


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


def _format_qb_date_value(value: object) -> str:
    """
    Normalize a BillPayment.date value into 'YYYY-MM-DD' for QBXML.
    Accepts date, datetime, or string like '2024-03-22 00:00:00'.
    """
    if isinstance(value, date) and not isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, datetime):
        return value.date().isoformat()

    s = str(value).strip()
    if not s:
        return date.today().isoformat()
    # e.g. '2024-03-22 00:00:00' -> '2024-03-22'
    return s[:10]


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
    if status_code not in (0, 1):
        print(f"QuickBooks error ({status_code}): {status_message}")
        raise RuntimeError(status_message)
    return root


def _escape_xml(value: str | None) -> str:
    if value is None:
        value = ""
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


# ---------------------------------------------------------------------------
# 1) Fetch BillPayment CHECKS (already-paid)
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
    from decimal import Decimal, InvalidOperation

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

        # Amount to Pay = sum of AppliedToTxnRet/PaymentAmount; fallback to header total
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
                # vendor left default/empty if your dataclass allows
            )
        )
    return payments


# ---------------------------------------------------------------------------
# 2) Fetch UNPAID bills as BillPayment-like objects (for comparison)
# ---------------------------------------------------------------------------


def fetch_unpaid_bills_as_billpayments(
    company_file: str | None = None,
) -> List[BillPayment]:
    """
    Fetch UNPAID bills from QuickBooks and return them as BillPayment-like objects.

    Mapping:
      - id      <- RefNumber
      - date    <- TxnDate
      - amount  <- AmountDue
      - vendor  <- VendorRef/FullName
    """
    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <BillQueryRq>\n"
        "      <PaidStatus>NotPaidOnly</PaidStatus>\n"
        "      <IncludeLineItems>false</IncludeLineItems>\n"
        "    </BillQueryRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    try:
        root = _send_qbxml(qbxml)
    except RuntimeError as exc:
        print(f"Error fetching unpaid bills from QuickBooks: {exc}")
        return []

    from decimal import Decimal

    bills: List[BillPayment] = []

    for bill_ret in root.findall(".//BillRet"):
        ref_number = (bill_ret.findtext("RefNumber") or "").strip()
        if not ref_number:
            continue

        txn_date_raw = bill_ret.findtext("TxnDate") or ""
        vendor_name = (bill_ret.findtext("VendorRef/FullName") or "").strip()

        try:
            txn_date = _parse_qb_date(txn_date_raw)
        except Exception:
            txn_date = date.today()

        amount_due_str = (bill_ret.findtext("AmountDue") or "0").strip()
        try:
            amount_due = float(Decimal(amount_due_str))
        except Exception:
            amount_due = 0.0

        bills.append(
            BillPayment(
                id=ref_number,
                date=txn_date,
                amount_to_pay=amount_due,
                vendor=vendor_name,
            )
        )

    return bills


# ---------------------------------------------------------------------------
# 3) Fetch ALL bills as dict[RefNumber -> (TxnID, AmountDue, TotalAmount)]
# ---------------------------------------------------------------------------


def fetch_all_bills() -> Dict[str, Tuple[str, float, float]]:
    """
    Fetch ALL bills (paid + unpaid) from QuickBooks.

    Returns:
        dict: RefNumber -> (TxnID, AmountDue, TotalAmount)
    """
    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <BillQueryRq>\n"
        "      <PaidStatus>All</PaidStatus>\n"
        "      <IncludeLineItems>false</IncludeLineItems>\n"
        "    </BillQueryRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    try:
        root = _send_qbxml(qbxml)
    except RuntimeError as exc:
        print(f"Error fetching all bills from QuickBooks: {exc}")
        return {}

    from decimal import Decimal

    bills: Dict[str, Tuple[str, float, float]] = {}

    for bill_ret in root.findall(".//BillRet"):
        ref_number = (bill_ret.findtext("RefNumber") or "").strip()
        txn_id = (bill_ret.findtext("TxnID") or "").strip()
        if not ref_number or not txn_id:
            continue

        amount_due_str = (bill_ret.findtext("AmountDue") or "0").strip()
        total_amount_str = (
            bill_ret.findtext("TotalAmount") or bill_ret.findtext("Amount") or "0"
        ).strip()

        try:
            amount_due = float(Decimal(amount_due_str))
        except Exception:
            amount_due = 0.0

        try:
            total_amount = float(Decimal(total_amount_str))
        except Exception:
            total_amount = 0.0

        bills[ref_number] = (txn_id, amount_due, total_amount)

    return bills


# ---------------------------------------------------------------------------
# 4) Unpaid bills for a single vendor
# ---------------------------------------------------------------------------


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

    bills: List[tuple[str, float]] = []
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


# ---------------------------------------------------------------------------
# 5) Add bill payments (single, batch, batch-for-existing-bills)
# ---------------------------------------------------------------------------


def add_bill_payment(company_file: str | None, payment: BillPayment) -> BillPayment:
    """Create a single bill payment check in QuickBooks and return the stored record."""

    from decimal import Decimal

    # Require a vendor name (BillPayment.vendor is str | None)
    vendor = payment.vendor
    if vendor is None or not vendor.strip():
        raise ValueError(
            "BillPayment.vendor is required to create a QuickBooks payment"
        )

    # Find unpaid bills for this vendor
    unpaid_bills = fetch_unpaid_bills_for_vendor(vendor)

    if not unpaid_bills:
        raise RuntimeError(f"No unpaid bills found for vendor: {vendor}")

    # Try to find a bill that matches the payment amount exactly
    bill_txn_id: str | None = None
    bill_amount_due = 0.0

    for txn_id, amount_due in unpaid_bills:
        if abs(amount_due - payment.amount_to_pay) < 0.01:  # Match within 1 cent
            bill_txn_id = txn_id
            bill_amount_due = amount_due
            print(f"Matched bill by amount: TxnID={txn_id}, Amount={amount_due}")
            break

    # If no exact match, use the first bill
    if bill_txn_id is None:
        bill_txn_id, bill_amount_due = unpaid_bills[0]
        print(
            f"No exact match found, using first bill: TxnID={bill_txn_id}, Amount={bill_amount_due}"
        )

    payment_amount = min(payment.amount_to_pay, bill_amount_due)
    txn_date_str = _format_qb_date_value(payment.date)

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <BillPaymentCheckAddRq>\n"
        "      <BillPaymentCheckAdd>\n"
        "        <PayeeEntityRef>\n"
        f"          <FullName>{_escape_xml(vendor)}</FullName>\n"
        "        </PayeeEntityRef>\n"
        f"        <TxnDate>{txn_date_str}</TxnDate>\n"
        "        <BankAccountRef>\n"
        "          <FullName>Chase</FullName>\n"
        "        </BankAccountRef>\n"
        "        <IsToBePrinted>false</IsToBePrinted>\n"
        f"        <Memo>{_escape_xml(payment.id)}</Memo>\n"
        "        <AppliedToTxnAdd>\n"
        f"          <TxnID>{_escape_xml(bill_txn_id)}</TxnID>\n"
        f"          <PaymentAmount>{Decimal(str(payment_amount)):.2f}</PaymentAmount>\n"
        "        </AppliedToTxnAdd>\n"
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

    memo_raw = payment_ret.findtext("Memo")
    memo = (memo_raw if memo_raw is not None else payment.id).strip()

    vendor_ret = payment_ret.findtext("PayeeEntityRef/FullName")
    if vendor_ret is not None and vendor_ret.strip():
        vendor_result = vendor_ret.strip()
    else:
        vendor_result = vendor  # fallback to original vendor string

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
        vendor=vendor_result,
    )


def add_bill_payments_batch(
    company_file: str | None, payments: List[BillPayment]
) -> List[BillPayment]:
    """Create multiple bill payments in QuickBooks in a single batch request."""

    if not payments:
        return []

    from decimal import Decimal

    requests: List[str] = []
    for payment in payments:
        vendor = payment.vendor
        if vendor is None or not vendor.strip():
            print("Warning: payment has no vendor; skipping...")
            continue

        # Find unpaid bills for this vendor
        unpaid_bills = fetch_unpaid_bills_for_vendor(vendor)

        if not unpaid_bills:
            print(f"Warning: No unpaid bills found for vendor: {vendor}, skipping...")
            continue

        # Try to find a bill that matches the payment amount exactly
        bill_txn_id: str | None = None
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
        txn_date_str = _format_qb_date_value(payment.date)

        # Build the QBXML snippet for this bill payment creation
        requests.append(
            "    <BillPaymentCheckAddRq>\n"
            "      <BillPaymentCheckAdd>\n"
            "        <PayeeEntityRef>\n"
            f"          <FullName>{_escape_xml(vendor)}</FullName>\n"
            "        </PayeeEntityRef>\n"
            f"        <TxnDate>{txn_date_str}</TxnDate>\n"
            "        <BankAccountRef>\n"
            "          <FullName>Chase</FullName>\n"
            "        </BankAccountRef>\n"
            "        <IsToBePrinted>false</IsToBePrinted>\n"
            f"        <Memo>{_escape_xml(payment.id)}</Memo>\n"
            "        <AppliedToTxnAdd>\n"
            f"          <TxnID>{_escape_xml(bill_txn_id)}</TxnID>\n"
            f"          <PaymentAmount>{Decimal(str(payment_amount)):.2f}</PaymentAmount>\n"
            "        </AppliedToTxnAdd>\n"
            "      </BillPaymentCheckAdd>\n"
            "    </BillPaymentCheckAddRq>"
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

    from decimal import Decimal

    # Parse all responses
    added_payments: List[BillPayment] = []
    for payment_ret in root.findall(".//BillPaymentCheckRet"):
        memo = (payment_ret.findtext("Memo") or "").strip()
        vendor_resp = (payment_ret.findtext("PayeeEntityRef/FullName") or "").strip()

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
                id=memo,
                date=txn_date,
                amount_to_pay=amount,
                vendor=vendor_resp or None,
            )
        )

    return added_payments


def add_bill_payments_for_existing_bills_batch(
    company_file: str | None,
    items_to_pay: List[tuple[BillPayment, str, float]],
) -> List[BillPayment]:
    """
    Create bill payments in QuickBooks for EXISTING bills only.

    items_to_pay: list of (payment, bill_txn_id, amount_due)
      - payment: BillPayment from Excel (id, date, amount_to_pay, vendor)
      - bill_txn_id: QuickBooks TxnID for that bill
      - amount_due: current AmountDue for that bill in QB
    """

    if not items_to_pay:
        return []

    from decimal import Decimal

    requests: List[str] = []

    for payment, bill_txn_id, amount_due in items_to_pay:
        vendor = payment.vendor
        if vendor is None or not vendor.strip():
            print("Warning: payment has no vendor; skipping...")
            continue

        try:
            payment_amount = float(min(payment.amount_to_pay, amount_due))
        except Exception:
            payment_amount = float(amount_due)

        txn_date_str = _format_qb_date_value(payment.date)

        requests.append(
            "    <BillPaymentCheckAddRq>\n"
            "      <BillPaymentCheckAdd>\n"
            "        <PayeeEntityRef>\n"
            f"          <FullName>{_escape_xml(vendor)}</FullName>\n"
            "        </PayeeEntityRef>\n"
            f"        <TxnDate>{txn_date_str}</TxnDate>\n"
            "        <BankAccountRef>\n"
            "          <FullName>Chase</FullName>\n"
            "        </BankAccountRef>\n"
            "        <IsToBePrinted>false</IsToBePrinted>\n"
            f"        <Memo>{_escape_xml(payment.id)}</Memo>\n"
            "        <AppliedToTxnAdd>\n"
            f"          <TxnID>{_escape_xml(bill_txn_id)}</TxnID>\n"
            f"          <PaymentAmount>{Decimal(str(payment_amount)):.2f}</PaymentAmount>\n"
            "        </AppliedToTxnAdd>\n"
            "      </BillPaymentCheckAdd>\n"
            "    </BillPaymentCheckAddRq>"
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
        print(f"Batch add (existing bills) failed: {exc}")
        return []

    from decimal import Decimal

    added_payments: List[BillPayment] = []

    for payment_ret in root.findall(".//BillPaymentCheckRet"):
        memo = (payment_ret.findtext("Memo") or "").strip()
        vendor_resp = (payment_ret.findtext("PayeeEntityRef/FullName") or "").strip()

        txn_date_raw = payment_ret.findtext("TxnDate")
        try:
            txn_date = _parse_qb_date(txn_date_raw) if txn_date_raw else date.today()
        except ValueError:
            txn_date = date.today()

        amount_str = (payment_ret.findtext("Amount") or "0").strip()
        try:
            amount = float(Decimal(amount_str))
        except Exception:
            amount = 0.0

        added_payments.append(
            BillPayment(
                id=memo,
                date=txn_date,
                amount_to_pay=amount,
                vendor=vendor_resp or None,
            )
        )

    return added_payments


# ---------------------------------------------------------------------------
# Simple read wrapper (used by __main__ demo)
# ---------------------------------------------------------------------------


def read_data() -> List[BillPayment]:
    """Read bill payments from QuickBooks."""
    return fetch_bill_payments()


__all__ = [
    "read_data",
    "fetch_bill_payments",
    "fetch_unpaid_bills_as_billpayments",
    "fetch_unpaid_bills_for_vendor",
    "fetch_all_bills",
    "add_bill_payment",
    "add_bill_payments_batch",
    "add_bill_payments_for_existing_bills_batch",
]


if __name__ == "__main__":
    for obj in read_data():
        print(str(obj))
