"""Domain models for payment term synchronisation."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass
class BillPayment:
    """Represents a bill payment."""

    id: str
    date: date
    amount_to_pay: float

    def __str__(self) -> str:
        return (
            f"BillPayment(id={self.id}, date={self.date}, "
            f"amount_to_pay={self.amount_to_pay})"
        )


__all__ = [
    "BillPayment",
]
