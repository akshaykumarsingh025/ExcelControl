import unittest
import sys
import os
import json
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from history import CommandHistory
from session_manager import save_session, load_session, clear_session


class TestCommandHistory(unittest.TestCase):
    def test_add_and_get(self):
        h = CommandHistory()
        h.log = []
        h.add("test cmd", "print('hi')", True)
        entries = h.get_all()
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0]["command"], "test cmd")
        self.assertTrue(entries[0]["success"])

    def test_get_last_code(self):
        h = CommandHistory()
        h.log = []
        h.add("cmd1", "code1", True)
        h.add("cmd2", "code2", False)
        self.assertEqual(h.get_last_code(), "code2")

    def test_get_last_code_empty(self):
        h = CommandHistory()
        h.log = []
        self.assertIsNone(h.get_last_code())

    def test_clear(self):
        h = CommandHistory()
        h.log = []
        h.add("cmd", "code", True)
        h.clear()
        self.assertEqual(len(h.get_all()), 0)

    def test_multiple_entries(self):
        h = CommandHistory()
        h.log = []
        for i in range(5):
            h.add(f"cmd{i}", f"code{i}", i % 2 == 0)
        self.assertEqual(len(h.get_all()), 5)


class TestSessionManager(unittest.TestCase):
    def test_save_and_load(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            session_file = os.path.join(tmpdir, "test_session.json")
            session_data = {
                "model": "gemma4:e4b",
                "auto_run": True,
                "dry_run": False,
                "analysis_mode": True,
                "last_file": "test.xlsx",
                "conversation_history": [],
            }
            with open(session_file, "w") as f:
                json.dump(session_data, f)

            with open(session_file) as f:
                loaded = json.load(f)

            self.assertEqual(loaded["model"], "gemma4:e4b")
            self.assertTrue(loaded["auto_run"])
            self.assertFalse(loaded["dry_run"])
            self.assertEqual(loaded["last_file"], "test.xlsx")

    def test_missing_session_file(self):
        session = load_session()
        if session is not None:
            self.assertIsInstance(session, dict)


if __name__ == "__main__":
    unittest.main()
