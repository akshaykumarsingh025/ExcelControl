# excel_controller.py
import xlwings as xw


class ExcelController:
    def __init__(self):
        self.app = None
        self.wb = None
        self.ws = None

    def connect_or_open(self, filepath=None):
        """Connect to already open Excel, or open a file, or create new workbook."""
        try:
            if filepath:
                self.app = xw.App(visible=True)
                self.wb = self.app.books.open(filepath)
            elif xw.apps:
                # Attach to already running Excel
                self.app = xw.apps.active
                self.wb = self.app.books.active
            else:
                # Open fresh Excel with new workbook
                self.app = xw.App(visible=True)
                self.wb = self.app.books.add()

            self.ws = self.wb.sheets.active
            self.app.screen_updating = True  # Ensure changes are visible live
            return True, f"Connected to: {self.wb.name}"
        except Exception as e:
            return False, str(e)

    def execute(self, code: str):
        """Execute AI-generated xlwings code safely."""
        if not self.wb:
            return False, "No Excel workbook connected."

        # Inject ws and wb into execution context
        local_vars = {"ws": self.ws, "wb": self.wb, "xw": xw}

        try:
            exec(code, {}, local_vars)
            self.wb.save()
            return True, "Done."
        except Exception as e:
            return False, f"Error: {str(e)}"

    def get_sheet_names(self):
        if self.wb:
            return [s.name for s in self.wb.sheets]
        return []

    def switch_sheet(self, name):
        if self.wb:
            self.ws = self.wb.sheets[name]
            return True
        return False

    def is_connected(self):
        return self.wb is not None

    def get_sheet_context(self) -> str:
        """Reads the first few rows to provide context to the AI."""
        if not self.ws:
            return ""
        try:
            # Read a small safe block: A1 to G5
            data = self.ws.range("A1:G5").value

            if data is None:
                return "Sheet is empty"

            # xlwings returns a single value if it's 1 cell, a 1D list if it's 1 row/col, or a 2D list.
            if not isinstance(data, list):
                data = [[data]]
            elif len(data) > 0 and not isinstance(data[0], list):
                data = [data]

            text_lines = []
            for row in data:
                # Clean out None values and convert to string
                clean_row = [str(cell) if cell is not None else "" for cell in row]
                if any(clean_row):  # only include rows that have at least some data
                    text_lines.append(" | ".join(clean_row))

            if not text_lines:
                return "Sheet is empty"

            return "\n".join(text_lines)
        except Exception as e:
            return f"Error reading context: {str(e)}"
