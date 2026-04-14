# session_manager.py
import json
import os


SESSION_DIR = os.path.join(os.path.expanduser("~"), ".excelai")
SESSION_FILE = os.path.join(SESSION_DIR, "session.json")


def save_session(
    conversation_history: list,
    model: str,
    auto_run: bool,
    dry_run: bool,
    analysis_mode: bool,
    last_file: str = "",
) -> bool:
    try:
        os.makedirs(SESSION_DIR, exist_ok=True)
        data = {
            "conversation_history": conversation_history,
            "model": model,
            "auto_run": auto_run,
            "dry_run": dry_run,
            "analysis_mode": analysis_mode,
            "last_file": last_file,
        }
        with open(SESSION_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


def load_session() -> dict | None:
    try:
        if not os.path.exists(SESSION_FILE):
            return None
        with open(SESSION_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def clear_session() -> bool:
    try:
        if os.path.exists(SESSION_FILE):
            os.remove(SESSION_FILE)
        return True
    except Exception:
        return False
