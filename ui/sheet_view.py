from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QLabel,
    QCheckBox,
    QHeaderView,
    QComboBox,
)


class SheetView(QWidget):
    cell_edited = None

    def __init__(self, parent=None):
        super().__init__(parent)
        self._excel_controller = None
        self._auto_refresh = False
        self._timer = QTimer(self)
        self._timer.timeout.connect(self.refresh_data)
        self.cell_edited = None
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        self.refresh_btn = QPushButton("🔄 Refresh")
        self.refresh_btn.clicked.connect(self.refresh_data)
        toolbar.addWidget(self.refresh_btn)

        self.auto_refresh_cb = QCheckBox("Auto-refresh (5s)")
        self.auto_refresh_cb.toggled.connect(self._toggle_auto_refresh)
        toolbar.addWidget(self.auto_refresh_cb)

        toolbar.addStretch()

        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(160)
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        toolbar.addWidget(QLabel("Sheet:"))
        toolbar.addWidget(self.sheet_combo)

        self.cell_label = QLabel("Cell: —")
        self.cell_label.setStyleSheet(
            "color: #89b4fa; font-weight: bold; font-size: 12px; padding: 0 12px;"
        )
        toolbar.addWidget(self.cell_label)

        self.size_label = QLabel("Rows: 0 | Cols: 0")
        self.size_label.setStyleSheet("color: #6c7086; font-size: 12px;")
        toolbar.addWidget(self.size_label)

        layout.addLayout(toolbar)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.verticalHeader().setDefaultSectionSize(24)
        self.table.cellChanged.connect(self._on_cell_changed)
        self.table.currentCellChanged.connect(self._on_current_cell_changed)
        layout.addWidget(self.table, stretch=1)

    def set_excel_controller(self, controller):
        self._excel_controller = controller

    def refresh_data(self):
        if not self._excel_controller or not self._excel_controller.is_connected():
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.size_label.setText("Rows: 0 | Cols: 0")
            return

        try:
            data = self._excel_controller.get_all_data()
            sheet_names = self._excel_controller.get_sheet_names()

            self.sheet_combo.blockSignals(True)
            current_sheet = self._excel_controller.get_current_sheet_name()
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheet_names)
            idx = self.sheet_combo.findText(current_sheet)
            if idx >= 0:
                self.sheet_combo.setCurrentIndex(idx)
            self.sheet_combo.blockSignals(False)

            if not data:
                self.table.setRowCount(0)
                self.table.setColumnCount(0)
                self.size_label.setText("Rows: 0 | Cols: 0")
                return

            rows = len(data)
            cols = max(len(row) for row in data) if data else 0

            self.table.blockSignals(True)
            self.table.setRowCount(rows)
            self.table.setColumnCount(cols)

            col_headers = []
            for c in range(cols):
                col_letter = self._col_index_to_letter(c)
                col_headers.append(col_letter)
            self.table.setHorizontalHeaderLabels(col_headers)

            for r in range(rows):
                row_data = data[r] if r < len(data) else []
                for c in range(cols):
                    val = row_data[c] if c < len(row_data) else None
                    display = str(val) if val is not None else ""
                    item = QTableWidgetItem(display)
                    if val is None:
                        item.setForeground(QColor("#585b70"))
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                    self.table.setItem(r, c, item)

            self.table.blockSignals(False)
            self.size_label.setText(f"Rows: {rows} | Cols: {cols}")

        except Exception as e:
            self.size_label.setText(f"Error: {str(e)[:40]}")

    def _on_cell_changed(self, row, col):
        if not self._excel_controller or not self._excel_controller.is_connected():
            return
        item = self.table.item(row, col)
        if not item:
            return

        col_letter = self._col_index_to_letter(col)
        cell_ref = f"{col_letter}{row + 1}"
        new_value = item.text()

        try:
            if new_value == "":
                self._excel_controller.ws.range(cell_ref).value = None
            else:
                try:
                    parsed = float(new_value)
                    if parsed == int(parsed):
                        parsed = int(parsed)
                    self._excel_controller.ws.range(cell_ref).value = parsed
                except ValueError:
                    self._excel_controller.ws.range(cell_ref).value = new_value
        except Exception as e:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Write Error", f"Could not write to {cell_ref}:\n{e}")

    def _on_current_cell_changed(self, row, col, _prev_row, _prev_col):
        if row < 0 or col < 0:
            self.cell_label.setText("Cell: —")
            return
        col_letter = self._col_index_to_letter(col)
        self.cell_label.setText(f"Cell: {col_letter}{row + 1}")

    def _on_sheet_changed(self, name):
        if not self._excel_controller or not name:
            return
        self._excel_controller.switch_sheet(name)
        self.refresh_data()

    def _toggle_auto_refresh(self, enabled):
        if enabled:
            self._timer.start(5000)
        else:
            self._timer.stop()

    @staticmethod
    def _col_index_to_letter(index):
        result = ""
        n = index
        while True:
            result = chr(ord("A") + n % 26) + result
            n = n // 26 - 1
            if n < 0:
                break
        return result
