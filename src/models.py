"""
Domain models for pay-bills synchronisation.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Literal

SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["data_conflict", "payment_only_in_quickbooks"]


@dataclass(slots=True)
class BillPayment:
    """Represents one bill payment row (either from Excel or from QuickBooks)."""

    id: str  # Parent ID / unique payment id
    date: date  # Payment date (will be normalised in __post_init__)
    amount_to_pay: float  # Payment amount (also normalised)
    vendor: str | None = None  # Vendor name (can be empty/None)
    source: SourceLiteral = "excel"

    def __post_init__(self) -> None:
        """Normalise date and amount_to_pay so the rest of the code can rely on types."""

        # --- Normalise date ---
        if isinstance(self.date, datetime):
            self.date = self.date.date()
        elif isinstance(self.date, str):
            s = self.date.strip()

            parsed: date | None = None

            # Try ISO-like first 10 chars: "YYYY-MM-DD" from "YYYY-MM-DD HH:MM:SS"
            try:
                parsed = date.fromisoformat(s[:10])
            except ValueError:
                parsed = None

            # Fallback to a few explicit formats if needed
            if parsed is None:
                for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%m/%d/%Y", "%m/%d/%y"):
                    try:
                        parsed = datetime.strptime(s, fmt).date()
                        break
                    except ValueError:
                        continue

            if parsed is None:
                raise ValueError(
                    f"Invalid date format for BillPayment.date: {self.date!r}"
                )

            self.date = parsed

        # --- Normalise amount_to_pay ---
        if isinstance(self.amount_to_pay, str):
            self.amount_to_pay = float(self.amount_to_pay.strip())

    def __str__(self) -> str:
        return (
            f"BillPayment(id={self.id}, date={self.date}, "
            f"amount={self.amount_to_pay}, vendor={self.vendor}, "
            f"source={self.source})"
        )


@dataclass(slots=True)
class Conflict:
    """Describes a discrepancy between Excel and QuickBooks payments."""

    record_id: str
    reason: ConflictReason

    excel_amount: float | None = None
    qb_amount: float | None = None

    excel_date: str | None = None  # ISO string for JSON
    qb_date: str | None = None

    excel_vendor: str | None = None
    qb_vendor: str | None = None


@dataclass(slots=True)
class ComparisonReport:
    """Result of comparing Excel vs QuickBooks payments."""

    # Payments only in Excel (candidate to add to QB)
    excel_only: list[BillPayment] = field(default_factory=list)

    # Payments only in QuickBooks (reported as payment_only_in_quickbooks conflicts)
    qb_only: list[BillPayment] = field(default_factory=list)

    # Payments with same id in both, but with different data
    conflicts: list[Conflict] = field(default_factory=list)


__all__ = [
    "BillPayment",
    "Conflict",
    "ComparisonReport",
    "SourceLiteral",
    "ConflictReason",
]
