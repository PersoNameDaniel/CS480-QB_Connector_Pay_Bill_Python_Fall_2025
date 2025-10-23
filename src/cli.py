"""
Command-line interface (CLI) for the QuickBooks Pay Bill Connector.

This script allows you to compare financial data between an Excel workbook
and QuickBooks, then generate a JSON discrepancy report.
"""

from __future__ import annotations
import argparse
import sys
from typing import Any, List, Dict

from .excel_reader import extract_account_debit_vendor
from .compare import compare_records
from .reporting import save_json_report

# Optional import for QuickBooks COM connection
try:
    from .qb_gateway import fetch_bill_payments  # type: ignore
except ImportError:
    fetch_bill_payments = None  # type: ignore


def main() -> int:
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Compare Excel Pay Bills against QuickBooks data."
    )

    parser.add_argument(
        "--workbook",
        required=True,
        help="Path to the Excel workbook (e.g., company_data.xlsx).",
    )
    parser.add_argument(
        "--sheet",
        default="account debit vendor",
        help="Name of the Excel sheet to read.",
    )
    parser.add_argument(
        "--output",
        default="report.json",
        help="Path to save the generated discrepancy report (JSON format).",
    )
    parser.add_argument(
        "--skip-qb",
        action="store_true",
        help="Skip QuickBooks data fetching (for offline testing).",
    )
    parser.add_argument(
        "--company-file",
        type=str,
        default=None,
        help="Optional QuickBooks company file path.",
    )

    args = parser.parse_args()

    print("Reading Excel workbook...")
    try:
        excel_data = extract_account_debit_vendor(args.workbook)
        print(f"Loaded {len(excel_data)} rows from Excel.")
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return 1

    qb_data: List[Dict[str, Any]] = []
    if args.skip_qb or fetch_bill_payments is None:
        print("Skipping QuickBooks fetch (using empty dataset).")
    else:
        try:
            print("Connecting to QuickBooks...")
            qb_payments = fetch_bill_payments(args.company_file)
            qb_data = [bp.__dict__ for bp in qb_payments]
            print(f"Retrieved {len(qb_data)} QuickBooks payments.")
        except Exception as e:
            print(f"Error fetching QuickBooks data: {e}")
            return 1

    print("Comparing Excel vs QuickBooks records...")
    try:
        result = compare_records(excel_data, qb_data)
    except Exception as e:
        print(f"Comparison failed: {e}")
        return 1

    try:
        save_json_report(result, args.output)
        print(f"Report saved successfully to {args.output}")
    except Exception as e:
        print(f"Failed to save report: {e}")
        return 1

    print("\nComparison complete.")
    print("Summary:")
    print(f"  Missing in Excel: {len(result['missing_in_excel'])}")
    print(f"  Missing in QuickBooks: {len(result['missing_in_qb'])}")
    print(f"  Amount mismatches: {len(result['amount_mismatches'])}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
