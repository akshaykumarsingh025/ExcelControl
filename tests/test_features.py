import unittest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.features import (
    build_gradebook_prompt,
    build_emi_prompt,
    build_test_case_prompt,
    build_reconciliation_prompt,
    build_inventory_dashboard_prompt,
    build_data_cleaning_prompt,
    build_payroll_prompt,
    build_gantt_prompt,
    build_medication_tracker_prompt,
    build_patient_schedule_prompt,
    build_rental_yield_prompt,
    build_property_comparison_prompt,
    build_clinical_cleaner_prompt,
    build_consolidator_prompt,
    build_shipping_tracker_prompt,
)


class TestFeaturePrompts(unittest.TestCase):
    def _assert_common_suffix(self, prompt):
        self.assertIn("Respond with ONLY Python code", prompt)
        self.assertIn("ws", prompt)
        self.assertIn("wb", prompt)

    def test_gradebook_prompt(self):
        prompt = build_gradebook_prompt(30, "HW:20, Mid:80", "Letter A-F", 60)
        self.assertIn("30", prompt)
        self.assertIn("Letter A-F", prompt)
        self._assert_common_suffix(prompt)

    def test_emi_prompt(self):
        prompt = build_emi_prompt(5000000, 8.5, 240, True, False)
        self.assertIn("5000000", prompt)
        self.assertIn("8.5", prompt)
        self.assertIn("amortization", prompt)
        self._assert_common_suffix(prompt)

    def test_emi_prompt_compare(self):
        prompt = build_emi_prompt(5000000, 8.5, 240, False, True)
        self.assertIn("comparison", prompt.lower())

    def test_test_case_prompt(self):
        prompt = build_test_case_prompt("User login", "High", True, True)
        self.assertIn("User login", prompt)
        self.assertIn("High", prompt)
        self.assertIn("negative", prompt.lower())
        self.assertIn("boundary", prompt.lower())
        self._assert_common_suffix(prompt)

    def test_reconciliation_prompt(self):
        prompt = build_reconciliation_prompt("Bank", "Ledger", "0.01")
        self.assertIn("Bank", prompt)
        self.assertIn("0.01", prompt)
        self._assert_common_suffix(prompt)

    def test_inventory_prompt(self):
        prompt = build_inventory_dashboard_prompt(
            "Product", "Stock", "Reorder", "Price", True, True
        )
        self.assertIn("Product", prompt)
        self.assertIn("ABC", prompt)
        self._assert_common_suffix(prompt)

    def test_data_cleaning_prompt(self):
        options = {
            "remove_duplicates": True,
            "trim_whitespace": True,
            "fix_dates": False,
        }
        prompt = build_data_cleaning_prompt("Sheet1", options)
        self.assertIn("Sheet1", prompt)
        self.assertIn("duplicates", prompt.lower())
        self._assert_common_suffix(prompt)

    def test_payroll_prompt(self):
        prompt = build_payroll_prompt("Basic", 40, 10, 12, 1.75, 10, False)
        self.assertIn("Basic", prompt)
        self.assertIn("40", prompt)
        self._assert_common_suffix(prompt)

    def test_medication_tracker_prompt(self):
        prompt = build_medication_tracker_prompt(5, 30, True, True, "Daily")
        self.assertIn("5", prompt)
        self.assertIn("30", prompt)
        self._assert_common_suffix(prompt)

    def test_patient_schedule_prompt(self):
        prompt = build_patient_schedule_prompt("09:00", "17:00", 15, "13:00", 60, 24, 0)
        self.assertIn("09:00", prompt)
        self._assert_common_suffix(prompt)

    def test_rental_yield_prompt(self):
        prompt = build_rental_yield_prompt("Price", "Rent", "Maintenance", 5)
        self.assertIn("Price", prompt)
        self.assertIn("5%", prompt)
        self._assert_common_suffix(prompt)

    def test_property_comparison_prompt(self):
        prompt = build_property_comparison_prompt(5, "Price, Location", True)
        self.assertIn("5", prompt)
        self.assertIn("ROI", prompt)
        self._assert_common_suffix(prompt)

    def test_clinical_cleaner_prompt(self):
        options = {"fix_dates": True, "normalize_units": True}
        prompt = build_clinical_cleaner_prompt(options)
        self.assertIn("ISO 8601", prompt)
        self._assert_common_suffix(prompt)

    def test_consolidator_prompt(self):
        prompt = build_consolidator_prompt(3, True, True, True)
        self.assertIn("3", prompt)
        self.assertIn("duplicate", prompt.lower())
        self._assert_common_suffix(prompt)

    def test_shipping_tracker_prompt(self):
        prompt = build_shipping_tracker_prompt(20, True, True, "Conditional")
        self.assertIn("20", prompt)
        self._assert_common_suffix(prompt)


if __name__ == "__main__":
    unittest.main()
