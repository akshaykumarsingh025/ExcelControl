from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QTextEdit, QLineEdit,
    QComboBox, QCheckBox, QSpinBox, QPushButton, QScrollArea, QWidget,
    QFormLayout, QSizePolicy,
)
from ui.workflow_base import WorkflowBase
from core.features import (
    build_test_case_prompt,
    build_bug_report_prompt,
    build_traceability_prompt,
    build_test_data_prompt,
)


class QAPage(WorkflowBase):
    def get_workflow_name(self):
        return "QA & Software Testing"

    def get_workflow_description(self):
        return "Generate test cases, bug reports, traceability matrices, and test data for software QA."

    def setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(8, 8, 8, 8)

        layout.addWidget(self._build_test_case_gen())
        layout.addWidget(self._build_bug_report_gen())
        layout.addWidget(self._build_test_case_reviewer())
        layout.addWidget(self._build_traceability_matrix())
        layout.addWidget(self._build_test_data_gen())
        layout.addStretch()

        scroll.setWidget(container)
        self.get_content_layout().addWidget(scroll)

    def _build_test_case_gen(self):
        group = QGroupBox("A1: AI Test Case Generator")
        form = QFormLayout()
        form.setSpacing(8)

        self.tc_requirement = QTextEdit()
        self.tc_requirement.setPlaceholderText(
            "As a user, I want to reset my password via email link so I can regain access..."
        )
        self.tc_requirement.setMaximumHeight(80)
        form.addRow("Requirement / User Story:", self.tc_requirement)

        self.tc_priority = QComboBox()
        self.tc_priority.addItems(["All", "Critical", "High", "Medium", "Low"])
        form.addRow("Priority Level:", self.tc_priority)

        self.tc_negative = QCheckBox("Include negative test cases")
        self.tc_negative.setChecked(True)
        form.addRow("", self.tc_negative)

        self.tc_boundary = QCheckBox("Include boundary value tests")
        self.tc_boundary.setChecked(True)
        form.addRow("", self.tc_boundary)

        btn = QPushButton("Generate Test Cases")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_generate_test_cases)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_bug_report_gen(self):
        group = QGroupBox("A2: Bug Report Generator")
        form = QFormLayout()
        form.setSpacing(8)

        self.bug_description = QTextEdit()
        self.bug_description.setPlaceholderText(
            "Login button doesn't respond when clicked on mobile Safari..."
        )
        self.bug_description.setMaximumHeight(70)
        form.addRow("Bug Description:", self.bug_description)

        self.bug_module = QLineEdit()
        self.bug_module.setPlaceholderText("e.g., Authentication Module")
        form.addRow("Affected Feature/Module:", self.bug_module)

        self.bug_severity = QComboBox()
        self.bug_severity.addItems(["Blocker", "Critical", "Major", "Minor", "Trivial"])
        self.bug_severity.setCurrentIndex(2)
        form.addRow("Severity:", self.bug_severity)

        btn = QPushButton("Generate Bug Report")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_generate_bug_report)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_test_case_reviewer(self):
        group = QGroupBox("A3: Test Case Reviewer")
        form = QFormLayout()
        form.setSpacing(8)

        info = QLabel(
            "Reads test cases from the current sheet and identifies gaps:\n"
            "missing negative scenarios, untested boundaries, and duplicate coverage."
        )
        info.setWordWrap(True)
        info.setObjectName("subheadingLabel")
        form.addRow(info)

        btn = QPushButton("Review Current Test Cases")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_review_test_cases)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_traceability_matrix(self):
        group = QGroupBox("A4: Requirement Traceability Matrix")
        form = QFormLayout()
        form.setSpacing(8)

        self.trace_requirements = QTextEdit()
        self.trace_requirements.setPlaceholderText(
            "REQ-001: User login\nREQ-002: Password reset\nREQ-003: Session timeout"
        )
        self.trace_requirements.setMaximumHeight(80)
        form.addRow("Requirements (one per line):", self.trace_requirements)

        self.trace_test_ids = QTextEdit()
        self.trace_test_ids.setPlaceholderText(
            "TC-001, TC-002, TC-003 (leave empty to auto-generate)"
        )
        self.trace_test_ids.setMaximumHeight(60)
        form.addRow("Test Case IDs (one per line):", self.trace_test_ids)

        btn = QPushButton("Generate Traceability Matrix")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_generate_traceability)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_test_data_gen(self):
        group = QGroupBox("A5: Test Data Generator")
        form = QFormLayout()
        form.setSpacing(8)

        self.td_type = QLineEdit()
        self.td_type.setPlaceholderText("emails, phone numbers, addresses, names, credit cards")
        form.addRow("Data Type:", self.td_type)

        self.td_count = QSpinBox()
        self.td_count.setRange(1, 10000)
        self.td_count.setValue(100)
        form.addRow("Number of Records:", self.td_count)

        self.td_invalid = QCheckBox("Include invalid/edge cases")
        self.td_invalid.setChecked(True)
        form.addRow("", self.td_invalid)

        self.td_international = QCheckBox("International formats")
        self.td_international.setChecked(False)
        form.addRow("", self.td_international)

        btn = QPushButton("Generate Test Data")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_generate_test_data)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _on_generate_test_cases(self):
        req = self.tc_requirement.toPlainText().strip()
        if not req:
            return
        prompt = build_test_case_prompt(
            requirement=req,
            priority=self.tc_priority.currentText(),
            include_negative=self.tc_negative.isChecked(),
            include_boundary=self.tc_boundary.isChecked(),
        )
        self.command_sent.emit(prompt)

    def _on_generate_bug_report(self):
        desc = self.bug_description.toPlainText().strip()
        if not desc:
            return
        prompt = build_bug_report_prompt(
            description=desc,
            module=self.bug_module.text().strip() or "General",
            severity=self.bug_severity.currentText(),
        )
        self.command_sent.emit(prompt)

    def _on_review_test_cases(self):
        sheet_data = []
        try:
            if hasattr(self, 'excel') and self.excel and self.excel.is_connected():
                sheet_data = self.excel.get_all_data()
        except Exception:
            pass

        if sheet_data and len(sheet_data) > 1:
            import json
            preview = sheet_data[:30]
            data_str = json.dumps(preview, indent=2, ensure_ascii=False)
            prompt = (
                "Review the following test cases and identify:\n"
                "1. Missing negative scenarios\n"
                "2. Untested boundary conditions\n"
                "3. Duplicate or overlapping test cases\n"
                "4. Missing edge cases\n"
                "5. Coverage gaps by feature area\n\n"
                f"Test case data:\n{data_str}\n\n"
                "Write Python code that creates a 'Test Review' sheet listing all gaps found, "
                "with columns: Gap Type, Description, Suggested Test Case, Priority. "
                "Use ws for the active sheet and wb for the workbook. "
                "Respond with ONLY Python code. No explanation, no markdown, no backticks."
            )
        else:
            prompt = (
                "Generate a comprehensive test case review template. "
                "Create a sheet with columns: Review ID, Gap Type (Missing Negative/Missing Boundary/"
                "Duplicate/Edge Case/Coverage Gap), Description, Suggested Test Case, Priority. "
                "Add sample review findings for a typical web application. "
                "Use ws for the active sheet and wb for the workbook. "
                "Respond with ONLY Python code. No explanation, no markdown, no backticks."
            )
        self.command_sent.emit(prompt)

    def _on_generate_traceability(self):
        reqs = self.trace_requirements.toPlainText().strip()
        if not reqs:
            return
        tc_ids = self.trace_test_ids.toPlainText().strip()
        prompt = build_traceability_prompt(
            requirements=reqs,
            test_case_ids=tc_ids or "Auto-generate",
        )
        self.command_sent.emit(prompt)

    def _on_generate_test_data(self):
        data_type = self.td_type.text().strip()
        if not data_type:
            return
        prompt = build_test_data_prompt(
            data_type=data_type,
            count=self.td_count.value(),
            include_invalid=self.td_invalid.isChecked(),
            international=self.td_international.isChecked(),
        )
        self.command_sent.emit(prompt)
