import unittest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from validators import (
    validate_ifsc,
    validate_account_number,
    match_bank_name,
    validate_serial_numbers,
    validate_row_completeness,
)


class TestValidateIFSC(unittest.TestCase):
    def test_valid_ifsc(self):
        is_valid, corrected = validate_ifsc("SBIN0001234")
        self.assertTrue(is_valid)
        self.assertEqual(corrected, "SBIN0001234")

    def test_lowercase_auto_upper(self):
        is_valid, corrected = validate_ifsc("sbin0001234")
        self.assertTrue(is_valid)
        self.assertEqual(corrected, "SBIN0001234")

    def test_ocr_correction_O_to_0(self):
        is_valid, corrected = validate_ifsc("SBINO001234")
        self.assertTrue(is_valid)
        self.assertEqual(corrected, "SBIN0001234")

    def test_digit_to_letter_correction(self):
        is_valid, corrected = validate_ifsc("2BIN0001234")
        self.assertTrue(is_valid)
        self.assertEqual(corrected, "SBIN0001234")

    def test_too_short(self):
        is_valid, corrected = validate_ifsc("SBIN0")
        self.assertFalse(is_valid)
        self.assertIsNone(corrected)

    def test_none_input(self):
        is_valid, corrected = validate_ifsc(None)
        self.assertFalse(is_valid)
        self.assertIsNone(corrected)

    def test_spaces_stripped(self):
        is_valid, corrected = validate_ifsc("SBIN 0 001234")
        self.assertTrue(is_valid)
        self.assertEqual(corrected, "SBIN0001234")

    def test_dots_stripped(self):
        is_valid, corrected = validate_ifsc("SBIN.0001234")
        self.assertTrue(is_valid)
        self.assertEqual(corrected, "SBIN0001234")


class TestValidateAccountNumber(unittest.TestCase):
    def test_valid_account(self):
        is_valid, cleaned = validate_account_number("1234567890")
        self.assertTrue(is_valid)
        self.assertEqual(cleaned, "1234567890")

    def test_leading_dots_stripped(self):
        is_valid, cleaned = validate_account_number("..1234567890")
        self.assertTrue(is_valid)
        self.assertEqual(cleaned, "1234567890")

    def test_spaces_removed(self):
        is_valid, cleaned = validate_account_number("1234 567 890")
        self.assertTrue(is_valid)
        self.assertEqual(cleaned, "1234567890")

    def test_too_short(self):
        is_valid, cleaned = validate_account_number("123")
        self.assertFalse(is_valid)

    def test_none_input(self):
        is_valid, cleaned = validate_account_number(None)
        self.assertFalse(is_valid)
        self.assertEqual(cleaned, "")

    def test_numeric_string(self):
        is_valid, cleaned = validate_account_number(1234567890)
        self.assertTrue(is_valid)
        self.assertEqual(cleaned, "1234567890")


class TestMatchBankName(unittest.TestCase):
    def test_exact_match(self):
        matched, confidence = match_bank_name("HDFC BANK")
        self.assertEqual(matched, "HDFC BANK")
        self.assertEqual(confidence, 1.0)

    def test_case_insensitive(self):
        matched, confidence = match_bank_name("hdfc bank")
        self.assertEqual(matched, "HDFC BANK")
        self.assertEqual(confidence, 1.0)

    def test_fuzzy_match(self):
        matched, confidence = match_bank_name("HDFC BAK")
        self.assertEqual(matched, "HDFC BANK")
        self.assertGreater(confidence, 0.5)

    def test_partial_containment(self):
        matched, confidence = match_bank_name("BANK OF INDIA")
        self.assertEqual(matched, "BANK OF INDIA")
        self.assertGreaterEqual(confidence, 0.8)

    def test_empty_string(self):
        matched, confidence = match_bank_name("")
        self.assertEqual(matched, "")
        self.assertEqual(confidence, 0.0)

    def test_doubled_words(self):
        matched, confidence = match_bank_name("HDFC BANK BANK")
        self.assertEqual(matched, "HDFC BANK")
        self.assertEqual(confidence, 1.0)


class TestValidateSerialNumbers(unittest.TestCase):
    def test_correct_serials(self):
        data = [["S.No", "Name"], [1, "Alice"], [2, "Bob"]]
        corrected, warnings = validate_serial_numbers(data)
        self.assertEqual(len(warnings), 0)

    def test_incorrect_serials(self):
        data = [["S.No", "Name"], [5, "Alice"], [7, "Bob"]]
        corrected, warnings = validate_serial_numbers(data)
        self.assertEqual(len(warnings), 2)
        self.assertEqual(corrected[1][0], 1)
        self.assertEqual(corrected[2][0], 2)

    def test_empty_data(self):
        corrected, warnings = validate_serial_numbers([])
        self.assertEqual(corrected, [])

    def test_header_only(self):
        data = [["S.No", "Name"]]
        corrected, warnings = validate_serial_numbers(data)
        self.assertEqual(len(warnings), 0)


class TestValidateRowCompleteness(unittest.TestCase):
    def test_full_row(self):
        score = validate_row_completeness(["a", "b", "c"], 3)
        self.assertEqual(score, 1.0)

    def test_half_row(self):
        score = validate_row_completeness(["a", None, "c"], 3)
        self.assertAlmostEqual(score, 2 / 3)

    def test_empty_row(self):
        score = validate_row_completeness([], 3)
        self.assertEqual(score, 0.0)

    def test_all_none(self):
        score = validate_row_completeness([None, None], 2)
        self.assertEqual(score, 0.0)


if __name__ == "__main__":
    unittest.main()
