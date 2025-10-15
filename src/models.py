"""Domain models for bill payment synchronisation."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["name_mismatch", "missing_in_excel", "missing_in_quickbooks"]


@dataclass(slots=True)
class BillPayment:
    """Represents a bill payment synchronised between Excel and QuickBooks."""

    bill: str
    date: str
    bank_account: str
    amount_to_pay: float


__all__ = ["BillPayment"]
