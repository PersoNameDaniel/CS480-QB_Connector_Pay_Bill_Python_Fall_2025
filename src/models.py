"""Domain models for payment term synchronisation."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class BillPayment:
    """Represents a bill payment."""

    bill: str
    date: str
    bank_account: str
    amount_to_pay: float


__all__ = [
    "BillPayment",
]
