import unittest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.dry_run import analyze_code, DryRunResult


class TestDryRun(unittest.TestCase):
    def test_detect_bold_formatting(self):
        code = 'ws["A1"].font.bold = True'
        result = analyze_code(code)
        self.assertEqual(len(result.formatting), 1)

    def test_detect_rename_sheet(self):
        code = 'ws.name = "My Sheet"'
        result = analyze_code(code)
        self.assertTrue(any("Rename" in s for s in result.sheet_ops))

    def test_detect_add_sheet(self):
        code = 'wb.sheets.add("NewSheet")'
        result = analyze_code(code)
        self.assertTrue(any("sheet" in s.lower() for s in result.sheet_ops))

    def test_detect_add_chart(self):
        code = 'ws.charts.add(100, 100, 400, 300)'
        result = analyze_code(code)
        self.assertTrue(any("chart" in s.lower() for s in result.sheet_ops))

    def test_syntax_error_produces_warning(self):
        code = "def ("
        result = analyze_code(code)
        self.assertTrue(len(result.warnings) > 0)
        self.assertTrue(any("syntax" in w.lower() for w in result.warnings))

    def test_empty_code_no_changes(self):
        code = ""
        result = analyze_code(code)
        self.assertEqual(len(result.cells_written), 0)
        self.assertEqual(len(result.formulas), 0)

    def test_summary_no_detectable(self):
        code = "x = 1"
        result = analyze_code(code)
        summary = result.summary()
        self.assertIn("DRY RUN", summary)
        self.assertIn("No detectable", summary)

    def test_detect_picture(self):
        code = 'ws.pictures.add("image.png")'
        result = analyze_code(code)
        self.assertTrue(any("picture" in s.lower() for s in result.sheet_ops))

    def test_detect_delete(self):
        code = 'ws.delete()'
        result = analyze_code(code)
        self.assertTrue(any("Delete" in s for s in result.sheet_ops))

    def test_result_types(self):
        result = DryRunResult([], [], [], [], [], [])
        self.assertIsInstance(result.cells_read, list)
        self.assertIsInstance(result.cells_written, list)
        self.assertIsInstance(result.formulas, list)
        self.assertIsInstance(result.formatting, list)
        self.assertIsInstance(result.sheet_ops, list)
        self.assertIsInstance(result.warnings, list)


if __name__ == "__main__":
    unittest.main()
