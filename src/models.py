"""Domain models for bill payment synchronisation."""

from __future__ import annotations

from dataclasses import dataclass, field
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


@dataclass(slots=True)
class Conflict:
    """Describes a discrepancy between Excel and QuickBooks payment terms."""

    bill: str
    excel_name: str | None
    qb_name: str | None
    reason: ConflictReason


@dataclass(slots=True)
class ComparisonReport:
    """Groups comparison outcomes for later processing."""

    excel_only: list[BillPayment] = field(default_factory=list)
    qb_only: list[BillPayment] = field(default_factory=list)
    conflicts: list[Conflict] = field(default_factory=list)


__all__ = [
    "BillPayment",
    "Conflict",
    "ComparisonReport",
    "ConflictReason",
    "SourceLiteral",
]
