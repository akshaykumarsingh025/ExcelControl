from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QTextEdit,
    QGroupBox,
    QMessageBox,
)

from ui.workflow_base import WorkflowBase
from ui.code_editor import CodeEditor
from ui.chat_panel import ChatPanel, AgentWorker
from core.code_validator import validate_code
from core.dry_run import analyze_code
from core.features import build_health_check_prompt, build_vba_to_python_prompt


class CodeEditorPage(WorkflowBase):
    def get_workflow_name(self) -> str:
        return "Code Editor"

    def get_workflow_description(self) -> str:
        return "Write, execute, and debug Python xlwings code with AI-powered error fixing and formula explanation."

    def setup_ui(self):
        content = self.get_content_layout()
        self._excel_controller = None
        self._agent = None
        self._worker = None

        self.code_editor = CodeEditor()
        self.code_editor.execute_requested = self._execute_code
        self.code_editor.validate_requested = self._validate_code
        self.code_editor.dry_run_requested = self._dry_run
        content.addWidget(self.code_editor, stretch=1)

        error_group = QGroupBox("Smart Error Fix")
        error_layout = QVBoxLayout(error_group)
        error_layout.setSpacing(6)

        self.error_input = QTextEdit()
        self.error_input.setMaximumHeight(80)
        self.error_input.setPlaceholderText("Paste the error message here...")
        error_layout.addWidget(self.error_input)

        error_btn_row = QHBoxLayout()
        self.analyze_error_btn = QPushButton("Analyze Error")
        self.analyze_error_btn.setObjectName("primaryBtn")
        self.analyze_error_btn.clicked.connect(self._analyze_error)
        error_btn_row.addWidget(self.analyze_error_btn)
        error_btn_row.addStretch()
        error_layout.addLayout(error_btn_row)

        content.addWidget(error_group)

        tools_row = QHBoxLayout()
        tools_row.setSpacing(8)

        formula_group = QGroupBox("Formula Explainer")
        formula_layout = QVBoxLayout(formula_group)
        formula_layout.setSpacing(4)

        self.formula_input = QLineEdit()
        self.formula_input.setPlaceholderText("e.g., =IFERROR(XLOOKUP(A2,Sheet2!A:C,3,0),0)")
        formula_layout.addWidget(self.formula_input)

        self.explain_formula_btn = QPushButton("Explain Formula")
        self.explain_formula_btn.setObjectName("primaryBtn")
        self.explain_formula_btn.clicked.connect(self._explain_formula)
        formula_layout.addWidget(self.explain_formula_btn)

        tools_row.addWidget(formula_group, stretch=1)

        vba_group = QGroupBox("VBA to Python")
        vba_layout = QVBoxLayout(vba_group)
        vba_layout.setSpacing(4)

        self.vba_input = QTextEdit()
        self.vba_input.setMaximumHeight(80)
        self.vba_input.setPlaceholderText("Paste VBA code here...")
        vba_layout.addWidget(self.vba_input)

        self.convert_vba_btn = QPushButton("Convert to Python")
        self.convert_vba_btn.setObjectName("primaryBtn")
        self.convert_vba_btn.clicked.connect(self._convert_vba)
        vba_layout.addWidget(self.convert_vba_btn)

        tools_row.addWidget(vba_group, stretch=1)

        content.addLayout(tools_row)

        health_row = QHBoxLayout()
        self.health_check_btn = QPushButton("🩺 Spreadsheet Health Check")
        self.health_check_btn.setObjectName("primaryBtn")
        self.health_check_btn.clicked.connect(self._health_check)
        health_row.addWidget(self.health_check_btn)
        health_row.addStretch()
        content.addLayout(health_row)

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

    def set_code(self, code: str):
        self.code_editor.set_code(code)

    def get_code(self) -> str:
        return self.code_editor.get_code()

    def _execute_code(self):
        code = self.code_editor.get_code()
        if not code.strip():
            self.code_editor.set_status("No code to execute", success=False)
            return
        if not self._excel_controller or not self._excel_controller.is_connected():
            self.code_editor.set_status("Not connected to Excel", success=False)
            return

        success, message = self._excel_controller.execute(code)
        if success:
            self.code_editor.set_status("Done", success=True)
        else:
            self.code_editor.set_status(message[:60], success=False)

        self.chat_panel.add_message(f"Executed code. Result: {message}", "system")

    def _validate_code(self):
        code = self.code_editor.get_code()
        if not code.strip():
            self.code_editor.set_status("No code to validate", success=False)
            return

        result = validate_code(code)
        if result.is_safe:
            self.code_editor.set_status("Code is safe", success=True)
            self.chat_panel.add_message("Validation: Code passed all safety checks.", "success")
        else:
            self.code_editor.set_status(f"Blocked: {len(result.issues)} issues", success=False)
            issues_text = "\n".join(f"  - {issue}" for issue in result.issues)
            self.chat_panel.add_message(f"Validation Issues:\n{issues_text}", "error")

    def _dry_run(self):
        code = self.code_editor.get_code()
        if not code.strip():
            self.code_editor.set_status("No code to analyze", success=False)
            return

        result = analyze_code(code)
        summary = result.summary()
        self.code_editor.set_status("Dry run complete", success=True)
        self.chat_panel.add_message(summary, "dryrun")

    def _analyze_error(self):
        error_text = self.error_input.toPlainText().strip()
        if not error_text:
            self.chat_panel.add_message("Please paste an error message first.", "error")
            return
        if not self._agent:
            self.chat_panel.add_message("No AI agent configured.", "error")
            return

        code_in_editor = self.code_editor.get_code()
        prompt = (
            f"I got this error in my Excel automation code: {error_text}\n\n"
            f"The code was:\n{code_in_editor}\n\n"
            f"Suggest a fix. Provide corrected Python xlwings code. "
            f"The sheet object is 'ws' and workbook is 'wb'."
        )

        self.chat_panel.add_message("Analyzing error...", "system")
        self.analyze_error_btn.setEnabled(False)
        self._worker = AgentWorker(self._agent, prompt)
        self._worker.response_ready.connect(self._on_error_analysis)
        self._worker.error_occurred.connect(self._on_worker_error)
        self._worker.start()

    def _explain_formula(self):
        formula = self.formula_input.text().strip()
        if not formula:
            self.chat_panel.add_message("Please enter a formula to explain.", "error")
            return
        if not self._agent:
            self.chat_panel.add_message("No AI agent configured.", "error")
            return

        prompt = f"Explain this Excel formula in plain English: {formula}"
        self.chat_panel.add_message("Explaining formula...", "system")
        self.explain_formula_btn.setEnabled(False)
        self._worker = AgentWorker(self._agent, prompt)
        self._worker.response_ready.connect(self._on_formula_explained)
        self._worker.error_occurred.connect(self._on_worker_error)
        self._worker.start()

    def _convert_vba(self):
        vba_code = self.vba_input.toPlainText().strip()
        if not vba_code:
            self.chat_panel.add_message("Please paste VBA code first.", "error")
            return
        if not self._agent:
            self.chat_panel.add_message("No AI agent configured.", "error")
            return

        prompt = build_vba_to_python_prompt(vba_code)
        self.chat_panel.add_message("Converting VBA to Python...", "system")
        self.convert_vba_btn.setEnabled(False)
        self._worker = AgentWorker(self._agent, prompt)
        self._worker.response_ready.connect(self._on_vba_converted)
        self._worker.error_occurred.connect(self._on_worker_error)
        self._worker.start()

    def _health_check(self):
        if not self._excel_controller or not self._excel_controller.is_connected():
            self.chat_panel.add_message("No Excel workbook connected.", "error")
            return
        if not self._agent:
            self.chat_panel.add_message("No AI agent configured.", "error")
            return

        data = self._excel_controller.get_all_data()
        prompt = build_health_check_prompt(data)
        self.chat_panel.add_message("Running health check...", "system")
        self.health_check_btn.setEnabled(False)
        self._worker = AgentWorker(self._agent, prompt, analysis_mode=True)
        self._worker.response_ready.connect(self._on_health_check)
        self._worker.error_occurred.connect(self._on_worker_error)
        self._worker.start()

    def _on_error_analysis(self, response):
        self.analyze_error_btn.setEnabled(True)
        self.chat_panel.add_message(f"Error Analysis:\n{response}", "ai")

    def _on_formula_explained(self, response):
        self.explain_formula_btn.setEnabled(True)
        self.chat_panel.add_message(f"Formula Explanation:\n{response}", "ai")

    def _on_vba_converted(self, response):
        self.convert_vba_btn.setEnabled(True)
        self.chat_panel.add_message(f"VBA → Python:\n{response}", "code")

    def _on_health_check(self, response):
        self.health_check_btn.setEnabled(True)
        self.chat_panel.add_message(f"Health Check:\n{response}", "analysis")

    def _on_worker_error(self, error_msg):
        self.analyze_error_btn.setEnabled(True)
        self.explain_formula_btn.setEnabled(True)
        self.convert_vba_btn.setEnabled(True)
        self.health_check_btn.setEnabled(True)
        self.chat_panel.add_message(f"Error: {error_msg}", "error")
