import csv
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QTextEdit,
    QComboBox,
    QCheckBox,
    QGroupBox,
    QFileDialog,
    QMessageBox,
    QScrollArea,
    QWidget,
)

from ui.workflow_base import WorkflowBase
from ui.chat_panel import ChatPanel, AgentWorker
from core.features import (
    build_data_cleaning_prompt,
    build_consolidator_prompt,
    build_invoice_extract_prompt,
    build_email_cleaner_prompt,
)
from image.ocr_pipeline import OcrPipeline
from image.preprocessor import ImagePreprocessor


class DataToolsPage(WorkflowBase):
    def get_workflow_name(self) -> str:
        return "Data Tools"

    def get_workflow_description(self) -> str:
        return "Data cleaning, multi-file intelligence, image-to-Excel, email export, and session macros."

    def setup_ui(self):
        content = self.get_content_layout()
        self._excel_controller = None
        self._agent = None
        self._worker = None
        self._selected_files = []
        self._selected_image_path = ""

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setContentsMargins(0, 0, 8, 0)
        scroll_layout.setSpacing(10)

        self._build_cleaning_section(scroll_layout)
        self._build_multi_file_section(scroll_layout)
        self._build_image_section(scroll_layout)
        self._build_email_section(scroll_layout)
        self._build_macro_section(scroll_layout)

        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        content.addWidget(scroll, stretch=2)

        self.chat_panel = ChatPanel()
        self.chat_panel.setMinimumHeight(140)
        self.chat_panel.command_sent.connect(self.command_sent.emit)
        content.addWidget(self.chat_panel, stretch=1)

    def set_excel_controller(self, controller):
        self._excel_controller = controller
        if self.chat_panel:
            self.chat_panel.set_excel_controller(controller)

    def set_agent(self, agent):
        self._agent = agent
        if self.chat_panel:
            self.chat_panel.set_agent(agent)

    def _build_cleaning_section(self, parent_layout):
        group = QGroupBox("Data Cleaning Copilot")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        sheet_row = QHBoxLayout()
        sheet_row.addWidget(QLabel("Data Sheet:"))
        self.clean_sheet_combo = QComboBox()
        self.clean_sheet_combo.setMinimumWidth(200)
        sheet_row.addWidget(self.clean_sheet_combo, stretch=1)
        layout.addLayout(sheet_row)

        options_row = QHBoxLayout()
        options_row.setSpacing(12)

        left_col = QVBoxLayout()
        left_col.setSpacing(4)
        self.cb_remove_dupes = QCheckBox("Remove duplicates")
        self.cb_trim_ws = QCheckBox("Trim whitespace")
        self.cb_fix_dates = QCheckBox("Fix date formats (DD/MM vs MM/DD)")
        self.cb_standardize_phones = QCheckBox("Standardize phone numbers")
        left_col.addWidget(self.cb_remove_dupes)
        left_col.addWidget(self.cb_trim_ws)
        left_col.addWidget(self.cb_fix_dates)
        left_col.addWidget(self.cb_standardize_phones)
        options_row.addLayout(left_col)

        right_col = QVBoxLayout()
        right_col.setSpacing(4)
        self.cb_normalize_text = QCheckBox("Normalize text (Title Case, etc.)")
        self.cb_remove_empty = QCheckBox("Remove empty rows")
        self.cb_fill_down = QCheckBox("Fill down missing values")
        right_col.addWidget(self.cb_normalize_text)
        right_col.addWidget(self.cb_remove_empty)
        right_col.addWidget(self.cb_fill_down)
        options_row.addLayout(right_col)

        layout.addLayout(options_row)

        self.clean_btn = QPushButton("🧹 Clean My Data")
        self.clean_btn.setObjectName("primaryBtn")
        self.clean_btn.clicked.connect(self._clean_data)
        layout.addWidget(self.clean_btn)

        parent_layout.addWidget(group)

    def _build_multi_file_section(self, parent_layout):
        group = QGroupBox("Multi-File Intelligence")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        file_row = QHBoxLayout()
        self.select_files_btn = QPushButton("📂 Select Files to Compare")
        self.select_files_btn.clicked.connect(self._select_files)
        file_row.addWidget(self.select_files_btn)

        self.selected_files_label = QLabel("No files selected")
        self.selected_files_label.setStyleSheet("color: #a6adc8; font-size: 12px;")
        self.selected_files_label.setWordWrap(True)
        file_row.addWidget(self.selected_files_label, stretch=1)
        layout.addLayout(file_row)

        op_row = QHBoxLayout()
        op_row.addWidget(QLabel("Operation:"))
        self.multi_op_combo = QComboBox()
        self.multi_op_combo.addItems(["Compare", "Merge", "VLOOKUP across files", "Diff"])
        op_row.addWidget(self.multi_op_combo, stretch=1)

        self.process_files_btn = QPushButton("Process Files")
        self.process_files_btn.setObjectName("primaryBtn")
        self.process_files_btn.clicked.connect(self._process_files)
        op_row.addWidget(self.process_files_btn)
        layout.addLayout(op_row)

        parent_layout.addWidget(group)

    def _build_image_section(self, parent_layout):
        group = QGroupBox("Image / Table → Excel")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        img_row = QHBoxLayout()
        self.select_image_btn = QPushButton("🖼️ Select Image")
        self.select_image_btn.clicked.connect(self._select_image)
        img_row.addWidget(self.select_image_btn)

        self.image_path_label = QLabel("No image selected")
        self.image_path_label.setStyleSheet("color: #a6adc8; font-size: 12px;")
        self.image_path_label.setWordWrap(True)
        img_row.addWidget(self.image_path_label, stretch=1)
        layout.addLayout(img_row)

        type_row = QHBoxLayout()
        type_row.addWidget(QLabel("Table Type:"))
        self.table_type_combo = QComboBox()
        self.table_type_combo.addItems([
            "Single Page",
            "Two-Page Spread",
            "Receipt",
            "Invoice",
            "Handwritten Notes",
        ])
        type_row.addWidget(self.table_type_combo, stretch=1)
        layout.addLayout(type_row)

        check_row = QHBoxLayout()
        self.cb_auto_validate = QCheckBox("Auto-validate extracted data")
        self.cb_smart_merge = QCheckBox("Smart merge multi-pass (more accurate, slower)")
        check_row.addWidget(self.cb_auto_validate)
        check_row.addWidget(self.cb_smart_merge)
        layout.addLayout(check_row)

        self.extract_btn = QPushButton("📄 Extract to Sheet")
        self.extract_btn.setObjectName("primaryBtn")
        self.extract_btn.clicked.connect(self._extract_image)
        layout.addWidget(self.extract_btn)

        parent_layout.addWidget(group)

    def _build_email_section(self, parent_layout):
        group = QGroupBox("Email This Sheet")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        email_row = QHBoxLayout()
        email_row.addWidget(QLabel("To:"))
        self.email_to = QLineEdit()
        self.email_to.setPlaceholderText("recipient@example.com")
        email_row.addWidget(self.email_to, stretch=1)
        layout.addLayout(email_row)

        subj_row = QHBoxLayout()
        subj_row.addWidget(QLabel("Subject:"))
        self.email_subject = QLineEdit()
        self.email_subject.setPlaceholderText("Spreadsheet data")
        subj_row.addWidget(self.email_subject, stretch=1)
        layout.addLayout(subj_row)

        self.email_body = QTextEdit()
        self.email_body.setMaximumHeight(60)
        self.email_body.setPlaceholderText("Optional message body...")
        layout.addWidget(self.email_body)

        fmt_row = QHBoxLayout()
        fmt_row.addWidget(QLabel("Format:"))
        self.email_format_combo = QComboBox()
        self.email_format_combo.addItems(["HTML Table", "CSV Attachment", "Both"])
        fmt_row.addWidget(self.email_format_combo, stretch=1)

        self.send_email_btn = QPushButton("📧 Send Sheet")
        self.send_email_btn.setObjectName("primaryBtn")
        self.send_email_btn.clicked.connect(self._send_email)
        fmt_row.addWidget(self.send_email_btn)
        layout.addLayout(fmt_row)

        parent_layout.addWidget(group)

    def _build_macro_section(self, parent_layout):
        group = QGroupBox("Export Macro")
        layout = QVBoxLayout(group)
        layout.setSpacing(6)

        name_row = QHBoxLayout()
        name_row.addWidget(QLabel("Script Name:"))
        self.macro_name_input = QLineEdit()
        self.macro_name_input.setPlaceholderText("my_automation")
        name_row.addWidget(self.macro_name_input, stretch=1)
        layout.addLayout(name_row)

        btn_row = QHBoxLayout()
        self.export_session_btn = QPushButton("📤 Export Current Session as Python Script")
        self.export_session_btn.setObjectName("primaryBtn")
        self.export_session_btn.clicked.connect(self._export_macro)
        btn_row.addWidget(self.export_session_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        parent_layout.addWidget(group)

    def refresh_sheet_list(self):
        self.clean_sheet_combo.clear()
        if self._excel_controller and self._excel_controller.is_connected():
            names = self._excel_controller.get_sheet_names()
            self.clean_sheet_combo.addItems(names)

    def _clean_data(self):
        if not self._agent:
            self.chat_panel.add_message("No AI agent configured.", "error")
            return

        sheet_name = self.clean_sheet_combo.currentText()
        if not sheet_name and self._excel_controller and self._excel_controller.is_connected():
            sheet_name = self._excel_controller.get_current_sheet_name()

        options = {
            "remove_duplicates": self.cb_remove_dupes.isChecked(),
            "trim_whitespace": self.cb_trim_ws.isChecked(),
            "fix_dates": self.cb_fix_dates.isChecked(),
            "standardize_phones": self.cb_standardize_phones.isChecked(),
            "normalize_text": self.cb_normalize_text.isChecked(),
            "remove_empty_rows": self.cb_remove_empty.isChecked(),
            "fill_down": self.cb_fill_down.isChecked(),
        }

        if not any(options.values()):
            self.chat_panel.add_message("Select at least one cleaning operation.", "error")
            return

        prompt = build_data_cleaning_prompt(sheet_name, options)
        self.chat_panel.add_message("Generating cleaning code...", "system")
        self.clean_btn.setEnabled(False)
        self._worker = AgentWorker(self._agent, prompt)
        self._worker.response_ready.connect(self._on_clean_response)
        self._worker.error_occurred.connect(self._on_worker_error)
        self._worker.start()

    def _on_clean_response(self, response):
        self.clean_btn.setEnabled(True)
        self.chat_panel.add_message(f"Cleaning Code:\n{response}", "code")

    def _select_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel Files", "", "Excel Files (*.xlsx *.xlsm *.xlsb *.xls *.csv)"
        )
        if paths:
            self._selected_files = paths
            display = "; ".join(os.path.basename(p) for p in paths)
            self.selected_files_label.setText(display)

    def _process_files(self):
        if not self._selected_files:
            self.chat_panel.add_message("Select files first.", "error")
            return
        if not self._agent:
            self.chat_panel.add_message("No AI agent configured.", "error")
            return

        operation = self.multi_op_combo.currentText()
        file_list = "\n".join(f"  - {f}" for f in self._selected_files)
        file_count = len(self._selected_files)

        if operation == "Compare":
            prompt = (
                f"Compare the following {file_count} Excel files and highlight differences in data, "
                f"formatting, and structure.\n\nFiles:\n{file_list}\n\n"
                f"Write Python xlwings code to open each file, read the data, and create a comparison "
                f"report in a new sheet. The sheet object is 'ws' and workbook is 'wb'."
            )
        elif operation == "Merge":
            prompt = build_consolidator_prompt(
                file_count, remove_dupes=True, add_source=True, align_columns=True
            )
            prompt += f"\n\nFiles to merge:\n{file_list}"
        elif operation == "VLOOKUP across files":
            prompt = (
                f"Perform a VLOOKUP-style operation across these {file_count} files.\n\nFiles:\n{file_list}\n\n"
                f"Write Python xlwings code to open each file, read a key column, and bring matching data "
                f"into the current workbook. The sheet object is 'ws' and workbook is 'wb'."
            )
        elif operation == "Diff":
            prompt = (
                f"Compute a detailed diff between these {file_count} files, showing added rows, "
                f"removed rows, and changed cells.\n\nFiles:\n{file_list}\n\n"
                f"Write Python xlwings code to produce a diff report sheet. "
                f"The sheet object is 'ws' and workbook is 'wb'."
            )
        else:
            prompt = f"Process these files: {file_list}"

        self.chat_panel.add_message(f"Processing {file_count} file(s) — {operation}...", "system")
        self.process_files_btn.setEnabled(False)
        self._worker = AgentWorker(self._agent, prompt)
        self._worker.response_ready.connect(self._on_multi_file_response)
        self._worker.error_occurred.connect(self._on_worker_error)
        self._worker.start()

    def _on_multi_file_response(self, response):
        self.process_files_btn.setEnabled(True)
        self.chat_panel.add_message(f"Result:\n{response}", "code")

    def _select_image(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "", "Images (*.png *.jpg *.jpeg *.bmp *.tiff)"
        )
        if path:
            self._selected_image_path = path
            self.image_path_label.setText(os.path.basename(path))

    def _extract_image(self):
        if not self._selected_image_path:
            self.chat_panel.add_message("Select an image first.", "error")
            return
        if not self._agent:
            self.chat_panel.add_message("No AI agent configured.", "error")
            return
        if not self._excel_controller or not self._excel_controller.is_connected():
            self.chat_panel.add_message("No Excel workbook connected.", "error")
            return

        self.chat_panel.add_message("Extracting data from image...", "system")
        self.extract_btn.setEnabled(False)

        try:
            preprocessor = ImagePreprocessor()
            pipeline = OcrPipeline(self._agent, preprocessor)
            result = pipeline.ask_with_image_json(self._selected_image_path)
            self.extract_btn.setEnabled(True)
            self.chat_panel.add_message(f"Extraction Result:\n{result}", "code")
        except Exception as e:
            self.extract_btn.setEnabled(True)
            self.chat_panel.add_message(f"Extraction Error: {e}", "error")

    def _send_email(self):
        to_addr = self.email_to.text().strip()
        if not to_addr:
            self.chat_panel.add_message("Please enter a recipient email address.", "error")
            return
        if not self._excel_controller or not self._excel_controller.is_connected():
            self.chat_panel.add_message("No Excel workbook connected.", "error")
            return

        data = self._excel_controller.get_all_data()
        if not data:
            self.chat_panel.add_message("The current sheet is empty.", "error")
            return

        subject = self.email_subject.text().strip() or "Spreadsheet Data"
        body = self.email_body.toPlainText().strip()
        fmt = self.email_format_combo.currentText()

        try:
            msg = MIMEMultipart()
            msg["From"] = "excelai@local"
            msg["To"] = to_addr
            msg["Subject"] = subject

            if body:
                msg.attach(MIMEText(body, "plain"))

            if fmt in ("HTML Table", "Both"):
                html_table = self._data_to_html_table(data)
                html_body = body + "<br><br>" + html_table if body else html_table
                msg.attach(MIMEText(html_body, "html"))

            if fmt in ("CSV Attachment", "Both"):
                csv_path = os.path.join(
                    os.environ.get("TEMP", os.path.expanduser("~")),
                    "excelai_export_temp.csv",
                )
                with open(csv_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    for row in data:
                        writer.writerow([cell if cell is not None else "" for cell in row])

                with open(csv_path, "rb") as f:
                    part = MIMEBase("text", "csv")
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition",
                        "attachment",
                        filename="spreadsheet_data.csv",
                    )
                    msg.attach(part)

                try:
                    os.remove(csv_path)
                except OSError:
                    pass

            with smtplib.SMTP("localhost", 25, timeout=10) as server:
                server.sendmail(msg["From"], [to_addr], msg.as_string())

            self.chat_panel.add_message(
                f"Email sent to {to_addr} (format: {fmt})", "success"
            )
        except ConnectionRefusedError:
            self.chat_panel.add_message(
                "Could not connect to SMTP server (localhost:25). "
                "Please configure an SMTP server or use a local relay.",
                "error",
            )
        except Exception as e:
            self.chat_panel.add_message(f"Email error: {e}", "error")

    def _export_macro(self):
        name = self.macro_name_input.text().strip() or "excelai_session"
        if not name.endswith(".py"):
            name += ".py"

        try:
            from history import CommandHistory

            history = CommandHistory()
            all_entries = history.get_all()
            successful = [e for e in all_entries if e.get("success")]

            if not successful:
                self.chat_panel.add_message("No successful commands in history to export.", "error")
                return

            path, _ = QFileDialog.getSaveFileName(
                self, "Export Macro", name, "Python Scripts (*.py)"
            )
            if not path:
                return

            lines = [
                '"""ExcelAI Session Macro — Auto-generated."""',
                "import xlwings as xw",
                "",
                "app = xw.App(visible=True)",
                "wb = app.books.add()",
                "ws = wb.sheets.active",
                "",
            ]

            for i, entry in enumerate(successful, 1):
                code = entry.get("code", "")
                cmd = entry.get("command", "")
                lines.append(f"# Command {i}: {cmd}")
                lines.append(code.strip())
                lines.append("")

            lines.append("wb.save()")
            lines.append('print("Macro complete.")')

            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))

            self.chat_panel.add_message(
                f"Exported {len(successful)} commands to:\n{path}", "success"
            )
        except Exception as e:
            self.chat_panel.add_message(f"Export error: {e}", "error")

    def _on_worker_error(self, error_msg):
        self.clean_btn.setEnabled(True)
        self.process_files_btn.setEnabled(True)
        self.extract_btn.setEnabled(True)
        self.chat_panel.add_message(f"Error: {error_msg}", "error")

    @staticmethod
    def _data_to_html_table(data: list[list]) -> str:
        rows_html = []
        for i, row in enumerate(data):
            tag = "th" if i == 0 else "td"
            cells = []
            for cell in row:
                val = str(cell) if cell is not None else ""
                cells.append(f"<{tag}>{val}</{tag}>")
            rows_html.append(f"<tr>{''.join(cells)}</tr>")

        return (
            '<table border="1" cellpadding="4" cellspacing="0" '
            'style="border-collapse: collapse; font-family: Arial, sans-serif;">'
            + "".join(rows_html)
            + "</table>"
        )
