# excel_controller.py
import xlwings as xw
from code_validator import validate_code
from dry_run import analyze_code
from undo_manager import UndoManager


class ExcelController:
    def __init__(self):
        self.app = None
        self.wb = None
        self.ws = None
        self.undo_manager = UndoManager()
        self.open_workbooks: dict[str, object] = {}

    def connect_or_open(self, filepath=None):
        try:
            if filepath:
                self.app = xw.App(visible=True)
                self.wb = self.app.books.open(filepath)
            elif xw.apps:
                self.app = xw.apps.active
                self.wb = self.app.books.active
            else:
                self.app = xw.App(visible=True)
                self.wb = self.app.books.add()

            self.ws = self.wb.sheets.active
            self.app.screen_updating = True
            self.open_workbooks[self.wb.name] = self.wb
            return True, f"Connected to: {self.wb.name}"
        except Exception as e:
            return False, str(e)

    def open_other_workbook(self, filepath: str):
        try:
            if not self.app:
                self.app = xw.App(visible=True)
            wb = self.app.books.open(filepath)
            self.open_workbooks[wb.name] = wb
            return True, f"Opened: {wb.name}"
        except Exception as e:
            return False, str(e)

    def get_open_workbook(self, name: str):
        if name in self.open_workbooks:
            return self.open_workbooks[name]
        try:
            if self.app:
                for b in self.app.books:
                    if b.name == name:
                        self.open_workbooks[name] = b
                        return b
        except Exception:
            pass
        return None

    def list_open_workbooks(self) -> list[str]:
        names = list(self.open_workbooks.keys())
        try:
            if self.app:
                for b in self.app.books:
                    if b.name not in names:
                        names.append(b.name)
                        self.open_workbooks[b.name] = b
        except Exception:
            pass
        return names

    def execute(self, code: str, dry_run_enabled: bool = False):
        if not self.wb:
            return False, "No Excel workbook connected."

        validation = validate_code(code)
        if not validation.is_safe:
            return False, f"Code blocked for safety:\n" + "\n".join(validation.issues)

        if dry_run_enabled:
            result = analyze_code(code)
            return True, f"DRY RUN:\n{result.summary()}"

        self.undo_manager.save_snapshot(self.wb)

        local_vars = {"ws": self.ws, "wb": self.wb, "xw": xw}

        try:
            exec(
                code,
                {
                    "__builtins__": {
                        "print": print,
                        "range": range,
                        "len": len,
                        "int": int,
                        "float": float,
                        "str": str,
                        "bool": bool,
                        "list": list,
                        "dict": dict,
                        "tuple": tuple,
                        "set": set,
                        "enumerate": enumerate,
                        "zip": zip,
                        "map": map,
                        "filter": filter,
                        "sorted": sorted,
                        "reversed": reversed,
                        "abs": abs,
                        "min": min,
                        "max": max,
                        "sum": sum,
                        "round": round,
                        "isinstance": isinstance,
                        "type": type,
                        "True": True,
                        "False": False,
                        "None": None,
                    }
                },
                local_vars,
            )
            self.wb.save()
            return True, "Done."
        except Exception as e:
            return False, f"Error: {str(e)}"

    def execute_analysis(self, code: str):
        if not self.wb:
            return False, "No Excel workbook connected."

        validation = validate_code(code)
        if not validation.is_safe:
            return False, f"Code blocked for safety:\n" + "\n".join(validation.issues)

        import io
        import contextlib

        local_vars = {
            "ws": self.ws,
            "wb": self.wb,
            "xw": xw,
            "statistics": __import__("statistics"),
            "math": __import__("math"),
            "collections": __import__("collections"),
            "datetime": __import__("datetime"),
        }

        output = io.StringIO()
        try:
            with contextlib.redirect_stdout(output):
                exec(
                    code,
                    {
                        "__builtins__": {
                            "print": print,
                            "range": range,
                            "len": len,
                            "int": int,
                            "float": float,
                            "str": str,
                            "bool": bool,
                            "list": list,
                            "dict": dict,
                            "tuple": tuple,
                            "set": set,
                            "enumerate": enumerate,
                            "zip": zip,
                            "map": map,
                            "filter": filter,
                            "sorted": sorted,
                            "reversed": reversed,
                            "abs": abs,
                            "min": min,
                            "max": max,
                            "sum": sum,
                            "round": round,
                            "isinstance": isinstance,
                            "type": type,
                            "True": True,
                            "False": False,
                            "None": None,
                        }
                    },
                    local_vars,
                )
            return True, output.getvalue()
        except Exception as e:
            return False, f"Analysis Error: {str(e)}"

    def undo(self):
        ok, snap_path = self.undo_manager.undo(self.wb)
        if ok and self.wb:
            try:
                current_path = self.wb.fullname
                self.wb.close()
                self.wb = self.app.books.open(snap_path)
                self.wb.save(current_path)
                self.ws = self.wb.sheets.active
                return True, "Undo applied."
            except Exception as e:
                return False, f"Undo restore failed: {e}"
        return ok, snap_path

    def redo(self):
        ok, snap_path = self.undo_manager.redo(self.wb)
        if ok and self.wb:
            try:
                current_path = self.wb.fullname
                self.wb.close()
                self.wb = self.app.books.open(snap_path)
                self.wb.save(current_path)
                self.ws = self.wb.sheets.active
                return True, "Redo applied."
            except Exception as e:
                return False, f"Redo restore failed: {e}"
        return ok, snap_path

    def can_undo(self):
        return self.undo_manager.can_undo()

    def can_redo(self):
        return self.undo_manager.can_redo()

    def get_sheet_names(self):
        if self.wb:
            return [s.name for s in self.wb.sheets]
        return []

    def switch_sheet(self, name):
        if self.wb:
            try:
                self.ws = self.wb.sheets[name]
                return True
            except Exception:
                return False
        return False

    def is_connected(self):
        return self.wb is not None

    def get_current_sheet_name(self) -> str:
        if self.ws:
            try:
                return self.ws.name
            except Exception:
                pass
        return ""

    def get_full_context(self) -> str:
        if not self.ws:
            return ""
        try:
            used = self.ws.used_range
            if used is None:
                return "Sheet is empty"
            last_row = min(used.last_cell.row, 50)
            last_col = min(used.last_cell.column, 20)
            if last_row <= 0 or last_col <= 0:
                return "Sheet is empty"
            col_letter = xw.utils.col_name(last_col)
            data = self.ws.range(f"A1:{col_letter}{last_row}").value
            if data is None:
                return "Sheet is empty"
            if not isinstance(data, list):
                data = [[data]]
            elif len(data) > 0 and not isinstance(data[0], list):
                data = [data]
            text_lines = []
            for row in data:
                clean_row = [str(cell) if cell is not None else "" for cell in row]
                if any(clean_row):
                    text_lines.append(" | ".join(clean_row))
            if not text_lines:
                return "Sheet is empty"
            return "\n".join(text_lines)
        except Exception as e:
            return f"Error reading context: {str(e)}"

    def get_sheet_context(self) -> str:
        if not self.ws:
            return ""
        try:
            data = self.ws.range("A1:G5").value
            if data is None:
                return "Sheet is empty"
            if not isinstance(data, list):
                data = [[data]]
            elif len(data) > 0 and not isinstance(data[0], list):
                data = [data]
            text_lines = []
            for row in data:
                clean_row = [str(cell) if cell is not None else "" for cell in row]
                if any(clean_row):
                    text_lines.append(" | ".join(clean_row))
            if not text_lines:
                return "Sheet is empty"
            return "\n".join(text_lines)
        except Exception as e:
            return f"Error reading context: {str(e)}"

    def cleanup(self):
        self.undo_manager.cleanup()
