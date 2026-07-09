import csv
import os

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QLabel,
    QCheckBox,
    QComboBox,
    QLineEdit,
    QHeaderView,
    QFileDialog,
    QMessageBox,
    QGroupBox,
)

from ui.workflow_base import WorkflowBase
from ui.chat_panel import ChatPanel, AgentWorker
from ui.sheet_view import SheetView


class SheetViewPage(WorkflowBase):
    def get_workflow_name(self) -> str:
        return "Sheet View"

    def get_workflow_description(self) -> str:
        return "Live interactive view of your Excel workbook data with AI chat and formula builder."

    def setup_ui(self):
        content = self.get_content_layout()
        self._excel_controller = None
        self._agent = None
        self._worker = None

        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(160)
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        toolbar.addWidget(QLabel("Sheet:"))
        toolbar.addWidget(self.sheet_combo)

        self.refresh_btn = QPushButton("🔄 Refresh Data")
        self.refresh_btn.setObjectName("primaryBtn")
        self.refresh_btn.clicked.connect(self._refresh_data)
        toolbar.addWidget(self.refresh_btn)

        self.auto_refresh_cb = QCheckBox("Auto-Refresh (5s)")
        self.auto_refresh_cb.toggled.connect(self._toggle_auto_refresh)
        toolbar.addWidget(self.auto_refresh_cb)

        self.edit_mode_cb = QCheckBox("Edit Mode")
        self.edit_mode_cb.setToolTip("When enabled, cell edits write back to Excel via xlwings")
        toolbar.addWidget(self.edit_mode_cb)

        toolbar.addStretch()

        self.cell_label = QLabel("Cell: —")
        self.cell_label.setStyleSheet(
            "color: #89b4fa; font-weight: bold; font-size: 12px; padding: 0 12px;"
        )
        toolbar.addWidget(self.cell_label)

        self.size_label = QLabel("Rows: 0 | Cols: 0")
        self.size_label.setStyleSheet("color: #6c7086; font-size: 12px;")
        toolbar.addWidget(self.size_label)

        content.addLayout(toolbar)

        self._auto_timer = QTimer(self)
        self._auto_timer.timeout.connect(self._refresh_data)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.verticalHeader().setDefaultSectionSize(24)
        self.table.cellChanged.connect(self._on_cell_changed)
        self.table.currentCellChanged.connect(self._on_current_cell_changed)
        self.table.cellDoubleClicked.connect(self._on_cell_double_clicked)
        content.addWidget(self.table, stretch=1)

        export_row = QHBoxLayout()
        export_row.setSpacing(8)
        self.export_csv_btn = QPushButton("📁 Export to CSV")
        self.export_csv_btn.clicked.connect(self._export_csv)
        export_row.addWidget(self.export_csv_btn)
        export_row.addStretch()
        content.addLayout(export_row)

        chat_group = QGroupBox("Chat with This Data")
        chat_layout = QVBoxLayout(chat_group)
        chat_layout.setSpacing(6)

        chat_input_row = QHBoxLayout()
        self.chat_input = QLineEdit()
        self.chat_input.setPlaceholderText("Ask a question about the current sheet data...")
        self.chat_input.returnPressed.connect(self._ask_chat)
        chat_input_row.addWidget(self.chat_input, stretch=1)

        self.ask_btn = QPushButton("Ask")
        self.ask_btn.setObjectName("primaryBtn")
        self.ask_btn.clicked.connect(self._ask_chat)
        chat_input_row.addWidget(self.ask_btn)

        chat_layout.addLayout(chat_input_row)

        self.chat_panel = ChatPanel()
        self.chat_panel.setMinimumHeight(120)
        self.chat_panel.command_sent.connect(self.command_sent.emit)
        chat_layout.addWidget(self.chat_panel, stretch=1)

        content.addWidget(chat_group, stretch=1)

        formula_group = QGroupBox("Formula Builder")
        formula_layout = QVBoxLayout(formula_group)
        formula_layout.setSpacing(6)

        formula_input_row = QHBoxLayout()
        self.formula_input = QLineEdit()
        self.formula_input.setPlaceholderText(
            'e.g., "Sum of column B where column A contains Sales"'
        )
        formula_input_row.addWidget(self.formula_input, stretch=1)

        self.gen_formula_btn = QPushButton("Generate Formula")
        self.gen_formula_btn.setObjectName("primaryBtn")
        self.gen_formula_btn.clicked.connect(self._generate_formula)
        formula_input_row.addWidget(self.gen_formula_btn)

        formula_layout.addLayout(formula_input_row)

        formula_result_row = QHBoxLayout()
        self.formula_result_label = QLabel("Formula: —")
        self.formula_result_label.setStyleSheet(
            "color: #a6e3a1; font-family: 'Cascadia Code', 'Consolas', monospace; font-size: 13px;"
        )
        self.formula_result_label.setWordWrap(True)
        formula_result_row.addWidget(self.formula_result_label, stretch=1)

        self.insert_formula_btn = QPushButton("Insert Formula")
        self.insert_formula_btn.setObjectName("successBtn")
        self.insert_formula_btn.setEnabled(False)
        self.insert_formula_btn.clicked.connect(self._insert_formula)
        formula_result_row.addWidget(self.insert_formula_btn)

        formula_layout.addLayout(formula_result_row)

        content.addWidget(formula_group)

        self._last_generated_formula = ""

    def set_excel_controller(self, controller):
        self._excel_controller = controller
        if self.chat_panel:
            self.chat_panel.set_excel_controller(controller)

    def set_agent(self, agent):
        self._agent = agent
        if self.chat_panel:
            self.chat_panel.set_agent(agent)

    def refresh_data(self):
        self._refresh_data()

    def _refresh_data(self):
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
                    if not self.edit_mode_cb.isChecked():
                        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    else:
                        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                    self.table.setItem(r, c, item)

            self.table.blockSignals(False)
            self.size_label.setText(f"Rows: {rows} | Cols: {cols}")

        except Exception as e:
            self.size_label.setText(f"Error: {str(e)[:40]}")

    def _on_cell_changed(self, row, col):
        if not self.edit_mode_cb.isChecked():
            return
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
            QMessageBox.warning(self, "Write Error", f"Could not write to {cell_ref}:\n{e}")

    def _on_cell_double_clicked(self, row, col):
        if self.edit_mode_cb.isChecked():
            self.table.editItem(self.table.item(row, col))

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
        self._refresh_data()

    def _toggle_auto_refresh(self, enabled):
        if enabled:
            self._auto_timer.start(5000)
        else:
            self._auto_timer.stop()

    def _get_sheet_data(self) -> list[list]:
        if not self._excel_controller or not self._excel_controller.is_connected():
            return []
        return self._excel_controller.get_all_data()

    def _ask_chat(self):
        question = self.chat_input.text().strip()
        if not question:
            return
        self.chat_input.clear()
        sheet_data = self._get_sheet_data()

        if self._agent and sheet_data:
            self.chat_panel.add_message(f"You: {question}", "user")
            self.chat_panel.add_message("Analyzing data...", "system")
            self.ask_btn.setEnabled(False)
            self._worker = AgentWorker(self._agent, "__chat_data__", context="")
            self._worker.response_ready.connect(self._on_chat_response)
            self._worker.error_occurred.connect(self._on_chat_error)
            self._worker.start()
            self._pending_chat_question = question
            self._pending_sheet_data = sheet_data
        elif self._agent:
            self.chat_panel.add_message(f"You: {question}", "user")
            self.chat_panel.command_sent.emit(question)
        else:
            self.chat_panel.add_message("No AI agent configured.", "error")

    def _on_chat_response(self, response):
        self.ask_btn.setEnabled(True)
        question = getattr(self, "_pending_chat_question", "")
        sheet_data = getattr(self, "_pending_sheet_data", [])
        if question and sheet_data and self._agent:
            try:
                result = self._agent.chat_with_data(question, sheet_data)
                self.chat_panel.add_message(f"AI:\n{result}", "ai")
            except Exception as e:
                self.chat_panel.add_message(f"AI:\n{response}", "ai")
        else:
            self.chat_panel.add_message(f"AI:\n{response}", "ai")

    def _on_chat_error(self, error_msg):
        self.ask_btn.setEnabled(True)
        self.chat_panel.add_message(f"Error: {error_msg}", "error")

    def _generate_formula(self):
        description = self.formula_input.text().strip()
        if not description:
            self.formula_result_label.setText("Formula: —")
            self.insert_formula_btn.setEnabled(False)
            return

        if not self._agent:
            self.formula_result_label.setText("Formula: No AI agent configured")
            self.insert_formula_btn.setEnabled(False)
            return

        context = ""
        if self._excel_controller and self._excel_controller.is_connected():
            context = self._excel_controller.get_sheet_context()

        self.gen_formula_btn.setEnabled(False)
        self.formula_result_label.setText("Formula: Generating...")

        self._formula_worker = AgentWorker(
            self._agent,
            f"__generate_formula__{description}",
            context=context,
        )
        self._formula_worker.response_ready.connect(self._on_formula_response)
        self._formula_worker.error_occurred.connect(self._on_formula_error)
        self._formula_worker.start()

    def _on_formula_response(self, response):
        self.gen_formula_btn.setEnabled(True)
        if self._agent:
            description = self.formula_input.text().strip()
            context = ""
            if self._excel_controller and self._excel_controller.is_connected():
                context = self._excel_controller.get_sheet_context()
            try:
                formula = self._agent.generate_formula(description, context)
                self._last_generated_formula = formula
                self.formula_result_label.setText(f"Formula: {formula}")
                self.insert_formula_btn.setEnabled(True)
                return
            except Exception:
                pass
        self.formula_result_label.setText(f"Formula: {response}")
        self._last_generated_formula = response
        self.insert_formula_btn.setEnabled(True)

    def _on_formula_error(self, error_msg):
        self.gen_formula_btn.setEnabled(True)
        self.formula_result_label.setText(f"Formula: Error - {error_msg}")
        self.insert_formula_btn.setEnabled(False)

    def _insert_formula(self):
        if not self._last_generated_formula:
            return
        if not self._excel_controller or not self._excel_controller.is_connected():
            QMessageBox.warning(self, "Not Connected", "No Excel workbook connected.")
            return

        current = self.table.currentItem()
        if not current:
            QMessageBox.information(
                self, "Select Cell", "Select a cell in the table first."
            )
            return

        row = current.row()
        col = current.column()
        col_letter = self._col_index_to_letter(col)
        cell_ref = f"{col_letter}{row + 1}"

        try:
            formula = self._last_generated_formula.strip()
            if not formula.startswith("="):
                formula = "=" + formula
            self._excel_controller.ws.range(cell_ref).formula = formula
            QMessageBox.information(
                self, "Inserted", f"Formula inserted into {cell_ref}:\n{formula}"
            )
            self._refresh_data()
        except Exception as e:
            QMessageBox.warning(
                self, "Insert Error", f"Could not insert formula into {cell_ref}:\n{e}"
            )

    def _export_csv(self):
        if not self._excel_controller or not self._excel_controller.is_connected():
            QMessageBox.warning(self, "Not Connected", "No Excel workbook connected.")
            return

        data = self._get_sheet_data()
        if not data:
            QMessageBox.information(self, "No Data", "The current sheet is empty.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Export to CSV", "", "CSV Files (*.csv)"
        )
        if not path:
            return

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                for row in data:
                    writer.writerow([cell if cell is not None else "" for cell in row])
            QMessageBox.information(self, "Exported", f"Data exported to:\n{path}")
        except Exception as e:
            QMessageBox.warning(self, "Export Error", f"Could not export CSV:\n{e}")

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
