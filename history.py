# history.py
import json
import os


HISTORY_DIR = os.path.join(os.path.expanduser("~"), ".excelai")
HISTORY_FILE = os.path.join(HISTORY_DIR, "command_history.json")


class CommandHistory:
    def __init__(self):
        self.log = []
        self._load()

    def add(self, command: str, code: str, success: bool):
        self.log.append({"command": command, "code": code, "success": success})
        self._save()

    def get_last_code(self):
        if self.log:
            return self.log[-1]["code"]
        return None

    def get_all(self):
        return self.log

    def get_commands(self) -> list[str]:
        return [item["command"] for item in self.log if item.get("command")]

    def clear(self):
        self.log = []
        self._save()

    def _save(self):
        try:
            os.makedirs(HISTORY_DIR, exist_ok=True)
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self.log, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def _load(self):
        try:
            if os.path.exists(HISTORY_FILE):
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    self.log = json.load(f)
        except Exception:
            self.log = []
