"""Domain models for payment term synchronisation."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class BillPayment:
    """Represents a bill payment."""

    bill: str
    date: str
    bank_account: str
    amount_to_pay: float

    def __str__(self) -> str:
        return (
            f"BillPayment(bill={self.bill}, date={self.date}, "
            f"bank_account={self.bank_account}, amount_to_pay={self.amount_to_pay})"
        )


__all__ = [
    "BillPayment",
]
