from PyQt6.QtCore import Qt, QTime
from PyQt6.QtWidgets import (
    QVBoxLayout, QGroupBox, QLabel, QTextEdit, QLineEdit,
    QComboBox, QCheckBox, QSpinBox, QDoubleSpinBox, QPushButton,
    QScrollArea, QWidget, QFormLayout, QFileDialog, QTimeEdit,
)
from ui.workflow_base import WorkflowBase
from core.features import (
    build_resume_parse_prompt,
    build_attendance_report_prompt,
    build_payroll_prompt,
    build_directory_prompt,
    build_onboarding_prompt,
)


class HRPage(WorkflowBase):
    def get_workflow_name(self):
        return "Human Resources"

    def get_workflow_description(self):
        return "Resume parsing, attendance reports, payroll calculation, employee directories, and onboarding checklists."

    def setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(8, 8, 8, 8)

        layout.addWidget(self._build_resume_parser())
        layout.addWidget(self._build_attendance())
        layout.addWidget(self._build_payroll())
        layout.addWidget(self._build_directory())
        layout.addWidget(self._build_onboarding())
        layout.addStretch()

        scroll.setWidget(container)
        self.get_content_layout().addWidget(scroll)

    def _build_resume_parser(self):
        group = QGroupBox("C1: Resume Parser")
        form = QFormLayout()
        form.setSpacing(8)

        self.resume_files_label = QLabel("No files selected")
        self.resume_files_label.setObjectName("subheadingLabel")
        form.addRow("Selected Files:", self.resume_files_label)

        btn_select = QPushButton("Select Resume Files")
        btn_select.clicked.connect(self._on_select_resumes)
        form.addRow("", btn_select)

        self.resume_extract_skills = QCheckBox("Extract skills from descriptions")
        self.resume_extract_skills.setChecked(True)
        form.addRow("", self.resume_extract_skills)

        self.resume_rank = QCheckBox("Rank by keyword match")
        self.resume_rank.setChecked(True)
        form.addRow("", self.resume_rank)

        self.resume_keywords = QLineEdit()
        self.resume_keywords.setPlaceholderText("Python, SQL, Machine Learning, React")
        form.addRow("Required Keywords:", self.resume_keywords)

        btn = QPushButton("Parse Resumes")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_parse_resumes)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_attendance(self):
        group = QGroupBox("C2: Attendance Report Generator")
        form = QFormLayout()
        form.setSpacing(8)

        self.att_name_col = QLineEdit("Employee Name")
        form.addRow("Employee Name Column:", self.att_name_col)

        self.att_date_col = QLineEdit("Date")
        form.addRow("Date Column:", self.att_date_col)

        self.att_status_col = QLineEdit("Status")
        form.addRow("Status Column:", self.att_status_col)

        self.att_work_start = QTimeEdit()
        self.att_work_start.setTime(QTime(9, 0))
        form.addRow("Work Start Time:", self.att_work_start)

        self.att_work_end = QTimeEdit()
        self.att_work_end.setTime(QTime(18, 0))
        form.addRow("Work End Time:", self.att_work_end)

        self.att_late_threshold = QSpinBox()
        self.att_late_threshold.setRange(1, 120)
        self.att_late_threshold.setValue(15)
        self.att_late_threshold.setSuffix(" min")
        form.addRow("Late Threshold:", self.att_late_threshold)

        btn = QPushButton("Generate Attendance Report")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_attendance_report)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_payroll(self):
        group = QGroupBox("C3: Payroll Calculator")
        form = QFormLayout()
        form.setSpacing(8)

        self.pay_basic_col = QLineEdit("Basic Salary")
        form.addRow("Basic Salary Column:", self.pay_basic_col)

        self.pay_hra = QDoubleSpinBox()
        self.pay_hra.setRange(0, 100)
        self.pay_hra.setValue(40)
        self.pay_hra.setSuffix(" %")
        form.addRow("HRA %:", self.pay_hra)

        self.pay_da = QDoubleSpinBox()
        self.pay_da.setRange(0, 100)
        self.pay_da.setValue(10)
        self.pay_da.setSuffix(" %")
        form.addRow("DA %:", self.pay_da)

        self.pay_pf = QDoubleSpinBox()
        self.pay_pf.setRange(0, 100)
        self.pay_pf.setValue(12)
        self.pay_pf.setSuffix(" %")
        form.addRow("PF %:", self.pay_pf)

        self.pay_esi = QDoubleSpinBox()
        self.pay_esi.setRange(0, 100)
        self.pay_esi.setValue(1.75)
        self.pay_esi.setSuffix(" %")
        form.addRow("ESI %:", self.pay_esi)

        self.pay_tds = QDoubleSpinBox()
        self.pay_tds.setRange(0, 100)
        self.pay_tds.setValue(10)
        self.pay_tds.setSuffix(" %")
        form.addRow("TDS %:", self.pay_tds)

        self.pay_overtime = QCheckBox("Include Overtime")
        self.pay_overtime.setChecked(False)
        form.addRow("", self.pay_overtime)

        btn = QPushButton("Calculate Payroll")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_payroll)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_directory(self):
        group = QGroupBox("C4: Employee Directory Builder")
        form = QFormLayout()
        form.setSpacing(8)

        self.dir_departments = QLineEdit()
        self.dir_departments.setPlaceholderText("Engineering, Marketing, Sales, HR, Finance")
        form.addRow("Department Names:", self.dir_departments)

        self.dir_photo = QCheckBox("Include photo placeholder column")
        self.dir_photo.setChecked(False)
        form.addRow("", self.dir_photo)

        self.dir_filter = QCheckBox("Add department filter dropdown")
        self.dir_filter.setChecked(True)
        form.addRow("", self.dir_filter)

        btn = QPushButton("Generate Directory")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_directory)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_onboarding(self):
        group = QGroupBox("C5: Onboarding Checklist")
        form = QFormLayout()
        form.setSpacing(8)

        self.onb_role = QComboBox()
        self.onb_role.addItems([
            "Software Engineer", "Designer", "Manager", "Sales", "Intern",
            "Data Analyst", "DevOps Engineer", "Product Manager",
        ])
        form.addRow("Role Type:", self.onb_role)

        self.onb_department = QComboBox()
        self.onb_department.addItems([
            "Engineering", "Design", "Product", "Sales", "Marketing", "HR", "Finance", "Operations",
        ])
        form.addRow("Department:", self.onb_department)

        self.onb_duration = QSpinBox()
        self.onb_duration.setRange(7, 180)
        self.onb_duration.setValue(90)
        self.onb_duration.setSuffix(" days")
        form.addRow("Duration:", self.onb_duration)

        btn = QPushButton("Generate Onboarding Checklist")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_onboarding)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _on_select_resumes(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Resume Files", "",
            "Documents (*.pdf *.docx *.txt);;All Files (*)"
        )
        if files:
            self.resume_files_label.setText(f"{len(files)} file(s) selected")
            self._resume_files = files

    def _on_parse_resumes(self):
        prompt = build_resume_parse_prompt(
            keywords=self.resume_keywords.text().strip(),
            rank=self.resume_rank.isChecked(),
            extract_skills=self.resume_extract_skills.isChecked(),
        )
        self.command_sent.emit(prompt)

    def _on_attendance_report(self):
        prompt = build_attendance_report_prompt(
            name_col=self.att_name_col.text().strip(),
            date_col=self.att_date_col.text().strip(),
            status_col=self.att_status_col.text().strip(),
            work_start=self.att_work_start.time().toString("HH:mm"),
            work_end=self.att_work_end.time().toString("HH:mm"),
            late_threshold=self.att_late_threshold.value(),
        )
        self.command_sent.emit(prompt)

    def _on_payroll(self):
        prompt = build_payroll_prompt(
            basic_col=self.pay_basic_col.text().strip(),
            hra_pct=self.pay_hra.value(),
            da_pct=self.pay_da.value(),
            pf_pct=self.pay_pf.value(),
            esi_pct=self.pay_esi.value(),
            tds_pct=self.pay_tds.value(),
            include_overtime=self.pay_overtime.isChecked(),
        )
        self.command_sent.emit(prompt)

    def _on_directory(self):
        prompt = build_directory_prompt(
            departments=self.dir_departments.text().strip(),
            photo_placeholder=self.dir_photo.isChecked(),
            filter_dropdown=self.dir_filter.isChecked(),
        )
        self.command_sent.emit(prompt)

    def _on_onboarding(self):
        prompt = build_onboarding_prompt(
            role=self.onb_role.currentText(),
            department=self.onb_department.currentText(),
            duration=self.onb_duration.value(),
        )
        self.command_sent.emit(prompt)
