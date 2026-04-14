# undo_manager.py
import os
import tempfile
import shutil
from datetime import datetime


class UndoManager:
    def __init__(self, max_snapshots=20):
        self.snapshots = []
        self.redo_stack = []
        self.max_snapshots = max_snapshots
        self.temp_dir = tempfile.mkdtemp(prefix="excelai_undo_")

    def save_snapshot(self, wb) -> bool:
        if wb is None:
            return False
        try:
            path = wb.fullname
            if not path or not os.path.exists(path):
                path = os.path.join(
                    self.temp_dir, f"snapshot_{len(self.snapshots)}.xlsx"
                )
                wb.save(path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            snap_path = os.path.join(
                self.temp_dir, f"snap_{timestamp}_{len(self.snapshots)}.xlsx"
            )
            shutil.copy2(path, snap_path)
            self.snapshots.append(
                {
                    "path": snap_path,
                    "timestamp": timestamp,
                    "original_path": wb.fullname,
                }
            )
            if len(self.snapshots) > self.max_snapshots:
                old = self.snapshots.pop(0)
                if os.path.exists(old["path"]):
                    os.remove(old["path"])
            self.redo_stack.clear()
            return True
        except Exception:
            return False

    def undo(self, wb) -> tuple[bool, str]:
        if not self.snapshots:
            return False, "No snapshots to undo."
        snap = self.snapshots.pop()
        try:
            if wb is not None:
                current_path = wb.fullname
                redo_path = os.path.join(
                    self.temp_dir,
                    f"redo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                )
                if os.path.exists(current_path):
                    shutil.copy2(current_path, redo_path)
                    self.redo_stack.append(
                        {
                            "path": redo_path,
                            "original_path": current_path,
                        }
                    )
            return True, snap["path"]
        except Exception as e:
            self.snapshots.append(snap)
            return False, f"Undo failed: {e}"

    def redo(self, wb) -> tuple[bool, str]:
        if not self.redo_stack:
            return False, "No redo snapshots available."
        snap = self.redo_stack.pop()
        try:
            if wb is not None:
                self.save_snapshot(wb)
            return True, snap["path"]
        except Exception as e:
            self.redo_stack.append(snap)
            return False, f"Redo failed: {e}"

    def can_undo(self):
        return len(self.snapshots) > 0

    def can_redo(self):
        return len(self.redo_stack) > 0

    def cleanup(self):
        try:
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        except Exception:
            pass
