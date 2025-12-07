"""
Command-line interface (CLI) for the QuickBooks Pay Bill Connector.

This script allows you to compare financial data between an Excel workbook
and QuickBooks, then generate a JSON discrepancy report.
"""

from __future__ import annotations
import argparse
import sys
from typing import Any, List, Dict
from dataclasses import asdict, is_dataclass
from pathlib import Path
from datetime import datetime

from .excel_reader import extract_account_debit_vendor, extract_account_debit_nonvendor
from .compare import compare_records
from .reporting import save_json_report
from .qb_gateway import fetch_bill_payments, add_bill_payments_batch
from .models import BillPayment


def _to_record_list(items: List[Any]) -> List[Dict[str, Any]]:
    """Convert a list of dataclass instances or dicts into list[dict]."""
    out: List[Dict[str, Any]] = []
    for item in items:
        if isinstance(item, dict):
            out.append(item)
        # ensure item is a dataclass instance (not the dataclass type)
        elif is_dataclass(item) and not isinstance(item, type):
            out.append(asdict(item))
        elif hasattr(item, "__dict__"):
            out.append(dict(item.__dict__))  # best-effort fallback
        else:
            out.append({"value": item})
    return out


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
        default="vendor",
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
        if args.sheet == "vendor":
            excel_data = extract_account_debit_vendor(args.workbook)
        elif args.sheet == "nonvendor":
            excel_data = extract_account_debit_nonvendor(args.workbook)
        print(f"Loaded {len(excel_data)} rows from Excel.")
    except Exception as e:
        print(f"Error reading Excel: {e}")
        return 1

    # convert dataclass instances to plain dicts for compare_records
    excel_records: List[Dict[str, Any]] = _to_record_list(excel_data)

    qb_data: List[Dict[str, Any]] = []
    if args.skip_qb or fetch_bill_payments is None:
        print("Skipping QuickBooks fetch (using empty dataset).")
    else:
        try:
            print("Connecting to QuickBooks...")
            qb_payments = fetch_bill_payments(args.company_file)
            qb_data = _to_record_list(qb_payments)
            print(f"Retrieved {len(qb_data)} QuickBooks payments.")
        except Exception as e:
            print(f"Error fetching QuickBooks data: {e}")
            return 1

    print("Comparing Excel vs QuickBooks records...")
    try:
        result = compare_records(excel_records, qb_data)
    except Exception as e:
        print(f"Comparison failed: {e}")
        return 1

    try:
        save_json_report(result, Path(args.output))
        print(f"Report saved successfully to {args.output}")
    except Exception as e:
        print(f"Failed to save report: {e}")
        return 1

    # Add missing records to QuickBooks
    if result["missing_in_qb"]:
        print("Adding missing records to QuickBooks...")
        try:
            # Convert missing records back to BillPayment objects
            missing_payments = [
                BillPayment(
                    source="quickbooks",
                    id=item["id"],  # Adjust based on your actual data structure
                    # Handle date conversion
                    date=(
                        datetime.strptime(item["date"], "%Y-%m-%dT%H:%M:%S")
                        if isinstance(item.get("date"), str)
                        else item["date"]  # Ensure it's a datetime object
                    ),
                    amount_to_pay=item["amount_to_pay"],
                    vendor=item["vendor"],  # Ensure this field exists in your data
                )
                for item in result["missing_in_qb"]
            ]
            added_payments = add_bill_payments_batch(
                args.company_file, missing_payments
            )
            print(f"Added {len(added_payments)} payments to QuickBooks.")
        except Exception as e:
            print(f"Failed to add payments to QuickBooks: {e}")
            return 1

    print("\nComparison complete.")
    print("Summary:")
    print(f"  Missing in Excel: {len(result['missing_in_excel'])}")
    print(f"  Missing in QuickBooks: {len(result['missing_in_qb'])}")
    print(f"  Amount mismatches: {len(result['amount_mismatches'])}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
