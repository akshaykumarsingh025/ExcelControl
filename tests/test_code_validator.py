import unittest
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from core.code_validator import validate_code, ValidationResult


class TestCodeValidator(unittest.TestCase):
    def test_safe_code_passes(self):
        code = 'ws["A1"].value = "Hello"'
        result = validate_code(code)
        self.assertTrue(result.is_safe)
        self.assertEqual(result.issues, [])

    def test_blocked_import_os(self):
        code = "import os\nos.listdir('.')"
        result = validate_code(code)
        self.assertFalse(result.is_safe)
        self.assertTrue(any("os" in i for i in result.issues))

    def test_blocked_import_subprocess(self):
        code = "import subprocess\nsubprocess.run(['rm', '-rf', '/'])"
        result = validate_code(code)
        self.assertFalse(result.is_safe)
        self.assertTrue(any("subprocess" in i for i in result.issues))

    def test_blocked_builtin_eval(self):
        code = "eval('1+1')"
        result = validate_code(code)
        self.assertFalse(result.is_safe)
        self.assertTrue(any("eval" in i for i in result.issues))

    def test_blocked_builtin_open(self):
        code = "open('/etc/passwd')"
        result = validate_code(code)
        self.assertFalse(result.is_safe)
        self.assertTrue(any("open" in i for i in result.issues))

    def test_blocked_dunder_class(self):
        code = "x.__class__"
        result = validate_code(code)
        self.assertFalse(result.is_safe)
        self.assertTrue(any("__class__" in i for i in result.issues))

    def test_syntax_error_fails(self):
        code = "def ("
        result = validate_code(code)
        self.assertFalse(result.is_safe)
        self.assertTrue(any("Syntax" in i for i in result.issues))

    def test_allowed_imports_pass(self):
        code = "import json\nimport math\nimport random"
        result = validate_code(code)
        self.assertTrue(result.is_safe)

    def test_xlwings_code_passes(self):
        code = (
            'ws.range("A1:C1").value = ["Name", "Age", "City"]\n'
            'ws.range("A1:C1").color = (144, 238, 144)\n'
            'ws.range("A1:C1").font.bold = True\n'
        )
        result = validate_code(code)
        self.assertTrue(result.is_safe)

    def test_blocked_import_from(self):
        code = "from pathlib import Path"
        result = validate_code(code)
        self.assertFalse(result.is_safe)
        self.assertTrue(any("pathlib" in i for i in result.issues))


if __name__ == "__main__":
    unittest.main()
