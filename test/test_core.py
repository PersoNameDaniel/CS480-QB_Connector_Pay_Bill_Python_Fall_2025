# This Module contains tests for the compare_records function
from __future__ import annotations
import unittest
from typing import List, Dict, Any
from src.compare import compare_records


class TestCompareRecords(unittest.TestCase):
    def setUp(self):
        self.excel_data: List[Dict[str, Any]] = [
            {"id": 1, "amount": 100.0},
            {"id": 2, "amount": 200.0},
            {"id": 3, "amount": 300.0},
        ]
        self.qb_data: List[Dict[str, Any]] = [
            {"id": 2, "amount": 200.0},
            {"id": 3, "amount": 350.0},
            {"id": 4, "amount": 400.0},
        ]

    def test_compare_records_basic(self):
        expected_discrepancies = {
            "missing_in_excel": [{"id": 4, "amount": 400.0}],
            "missing_in_qb": [{"id": 1, "amount": 100.0}],
            "amount_mismatches": [{"id": 3, "excel_amount": 300.0, "qb_amount": 350.0}],
        }
        result = compare_records(self.excel_data, self.qb_data)
        self.assertEqual(result, expected_discrepancies)

    def test_empty_datasets(self):
        result = compare_records([], [])
        expected = {
            "missing_in_excel": [],
            "missing_in_qb": [],
            "amount_mismatches": [],
        }
        self.assertEqual(result, expected)

    def test_identical_datasets(self):
        identical_data = [{"id": 1, "amount": 100.0}, {"id": 2, "amount": 200.0}]
        result = compare_records(identical_data, identical_data)
        expected = {
            "missing_in_excel": [],
            "missing_in_qb": [],
            "amount_mismatches": [],
        }
        self.assertEqual(result, expected)

    def test_completely_different_datasets(self):
        excel = [{"id": 1, "amount": 100.0}]
        qb = [{"id": 2, "amount": 200.0}]
        result = compare_records(excel, qb)
        expected = {
            "missing_in_excel": [{"id": 2, "amount": 200.0}],
            "missing_in_qb": [{"id": 1, "amount": 100.0}],
            "amount_mismatches": [],
        }
        self.assertEqual(result, expected)

    def test_multiple_amount_mismatches(self):
        excel = [{"id": 1, "amount": 100.0}, {"id": 2, "amount": 200.0}]
        qb = [{"id": 1, "amount": 150.0}, {"id": 2, "amount": 250.0}]
        result = compare_records(excel, qb)
        expected = {
            "missing_in_excel": [],
            "missing_in_qb": [],
            "amount_mismatches": [
                {"id": 1, "excel_amount": 100.0, "qb_amount": 150.0},
                {"id": 2, "excel_amount": 200.0, "qb_amount": 250.0},
            ],
        }
        self.assertEqual(result, expected)


if __name__ == "__main__":
    unittest.main()
