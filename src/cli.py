"""
Command-line interface for the Pay Bills synchroniser.

This module provides the entry point for running the tool from the command
line. It parses arguments, invokes the high-level runner, and prints where the
JSON report was written.
"""

from __future__ import annotations

import argparse
import sys

from .runner import run_pay_bills  # Orchestrator for Pay Bills workflow


def main(argv: list[str] | None = None) -> int:
    # Create a user-friendly argument parser
    parser = argparse.ArgumentParser(
        description="Synchronise bill payments between Excel and QuickBooks"
    )

    # Required: Excel file path
    parser.add_argument(
        "--workbook",
        required=True,
        help="Path to the Excel workbook (company_data.xlsx)",
    )

    # Required: sheet type (vendor or nonvendor)
    parser.add_argument(
        "--sheet",
        required=True,
        choices=["vendor", "nonvendor"],
        help="Specify which account debit sheet to process: vendor or nonvendor",
    )

    # Optional: output JSON report
    parser.add_argument(
        "--output",
        help="Optional path for the generated JSON report (default: pay_bills_report.json)",
    )

    # Parse arguments (from argv or sys.argv)
    args = parser.parse_args(argv)

    # Run the synchronisation using the currently open QuickBooks company file ("")
    path = run_pay_bills(
        "",  # Use currently-open company file
        args.workbook,
        sheet_type=args.sheet,
        output_path=args.output,
    )

    print(f"Report written to {path}")
    return 0  # success exit code


if __name__ == "__main__":
    sys.exit(main())
