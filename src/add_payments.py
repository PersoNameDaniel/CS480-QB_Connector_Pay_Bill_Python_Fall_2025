from qb_gateway import add_bill_payment, add_bill_payments_batch
from models import BillPayment
from datetime import date


# Option 1: Add a single bill payment
def add_single_payment():
    payment = BillPayment(
        id="9954",
        date=date(2025, 11, 10),
        amount_to_pay=1500.00,
        vendor="ATT(cell phone)",
    )

    result = add_bill_payment(company_file=None, payment=payment)
    print(f"Added payment: {result}")


# Option 2: Add multiple bill payments in a batch
def add_multiple_payments():
    payments = [
        BillPayment(
            id="9934",
            date=date(2025, 11, 10),
            amount_to_pay=1000.00,
            vendor="ATT(cell phone)",
        ),
        BillPayment(
            id="9935",
            date=date(2025, 11, 11),
            amount_to_pay=2500.00,
            vendor="ATT(cell phone)",
        ),
        BillPayment(
            id="9936",
            date=date(2025, 11, 12),
            amount_to_pay=750.00,
            vendor="ATT(cell phone)",
        ),
    ]

    results = add_bill_payments_batch(company_file=None, payments=payments)
    print(f"Added {len(results)} payments:")
    for payment in results:
        print(f"  - {payment}")


if __name__ == "__main__":
    # Choose which function to run
    add_single_payment()
    add_multiple_payments()
