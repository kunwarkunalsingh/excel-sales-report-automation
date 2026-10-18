
import unittest

import pandas as pd

from main import (
    REQUIRED_COLUMNS,
    prepare_sales_data,
    create_report_tables,
)


class TestSalesAutomation(unittest.TestCase):

    def setUp(self):
        self.sample_data = pd.DataFrame(
            [
                [1, "2026-01-05", "Laptop", "Electronics", "North", 2, 100],
                [1, "2026-01-05", "Laptop", "Electronics", "North", 2, 100],
                [2, "2026-01-10", "Mouse", "Accessories", "South", 3, 20],
                [3, "2026-02-02", "Chair", "Furniture", "", 1, 50],
                [4, "2026-02-04", "Desk", "Furniture", "West", 0, 80],
            ],
            columns=REQUIRED_COLUMNS,
        )

    def test_exact_duplicates_are_removed(self):
        valid, invalid, duplicate_count = prepare_sales_data(
            self.sample_data
        )

        self.assertEqual(duplicate_count, 1)
        self.assertEqual(len(valid) + len(invalid), 4)

    def test_invalid_records_are_flagged(self):
        valid, invalid, _ = prepare_sales_data(self.sample_data)

        self.assertEqual(len(invalid), 2)
        self.assertIn("validation_issues", invalid.columns)

    def test_revenue_is_calculated_correctly(self):
        valid, _, _ = prepare_sales_data(self.sample_data)

        self.assertAlmostEqual(valid["revenue"].sum(), 260.0)

    def test_summary_reports_correct_record_counts(self):
        valid, invalid, duplicates = prepare_sales_data(
            self.sample_data
        )

        summary, _, _, _ = create_report_tables(
            valid, invalid, duplicates
        )

        metrics = dict(zip(summary["Metric"], summary["Value"]))

        self.assertEqual(metrics["Valid sales records"], 2)
        self.assertEqual(metrics["Records requiring review"], 2)
        self.assertEqual(metrics["Duplicate records removed"], 1)
        self.assertEqual(metrics["Total revenue"], 260.0)


if __name__ == "__main__":
    unittest.main()
