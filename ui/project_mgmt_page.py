from PyQt6.QtCore import Qt, QDate
from PyQt6.QtWidgets import (
    QVBoxLayout, QGroupBox, QLabel, QTextEdit, QLineEdit,
    QComboBox, QCheckBox, QSpinBox, QPushButton, QScrollArea,
    QWidget, QFormLayout, QDateEdit,
)
from ui.workflow_base import WorkflowBase
from core.features import (
    build_gantt_prompt,
    build_sprint_backlog_prompt,
    build_risk_register_prompt,
    build_raci_prompt,
    build_status_report_prompt,
)


class ProjectMgmtPage(WorkflowBase):
    def get_workflow_name(self):
        return "Project Management"

    def get_workflow_description(self):
        return "Gantt charts, sprint backlogs, risk registers, RACI matrices, and status reports."

    def setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(8, 8, 8, 8)

        layout.addWidget(self._build_gantt())
        layout.addWidget(self._build_sprint())
        layout.addWidget(self._build_risk())
        layout.addWidget(self._build_raci())
        layout.addWidget(self._build_status())
        layout.addStretch()

        scroll.setWidget(container)
        self.get_content_layout().addWidget(scroll)

    def _build_gantt(self):
        group = QGroupBox("E1: Gantt Chart Builder")
        form = QFormLayout()
        form.setSpacing(8)

        self.gantt_tasks = QTextEdit()
        self.gantt_tasks.setPlaceholderText(
            "Requirements Gathering, 5\n"
            "Design, 7\n"
            "Development, 14\n"
            "Testing, 7\n"
            "Deployment, 3"
        )
        self.gantt_tasks.setMaximumHeight(100)
        form.addRow("Tasks (Name, Duration days):", self.gantt_tasks)

        self.gantt_start = QDateEdit()
        self.gantt_start.setCalendarPopup(True)
        self.gantt_start.setDate(QDate.currentDate())
        form.addRow("Project Start Date:", self.gantt_start)

        btn = QPushButton("Create Gantt Chart")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_gantt)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_sprint(self):
        group = QGroupBox("E2: Sprint Backlog Template")
        form = QFormLayout()
        form.setSpacing(8)

        self.sprint_name = QLineEdit("Sprint 1")
        form.addRow("Sprint Name:", self.sprint_name)

        self.sprint_weeks = QSpinBox()
        self.sprint_weeks.setRange(1, 8)
        self.sprint_weeks.setValue(2)
        self.sprint_weeks.setSuffix(" weeks")
        form.addRow("Sprint Duration:", self.sprint_weeks)

        self.sprint_velocity = QSpinBox()
        self.sprint_velocity.setRange(5, 200)
        self.sprint_velocity.setValue(30)
        self.sprint_velocity.setSuffix(" pts")
        form.addRow("Team Velocity:", self.sprint_velocity)

        self.sprint_stories = QTextEdit()
        self.sprint_stories.setPlaceholderText(
            "As a user, I want to login, 5\n"
            "As a user, I want to reset password, 3\n"
            "As a user, I want to view dashboard, 8"
        )
        self.sprint_stories.setMaximumHeight(80)
        form.addRow("User Stories (Name, Points):", self.sprint_stories)

        btn = QPushButton("Generate Sprint Backlog")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_sprint)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_risk(self):
        group = QGroupBox("E3: Risk Register")
        form = QFormLayout()
        form.setSpacing(8)

        self.risk_project = QTextEdit()
        self.risk_project.setPlaceholderText("Project name and brief description...")
        self.risk_project.setMaximumHeight(60)
        form.addRow("Project Description:", self.risk_project)

        self.risk_categories = QSpinBox()
        self.risk_categories.setRange(2, 15)
        self.risk_categories.setValue(5)
        form.addRow("Number of Risk Categories:", self.risk_categories)

        btn = QPushButton("Generate Risk Register")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_risk)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_raci(self):
        group = QGroupBox("E4: RACI Matrix")
        form = QFormLayout()
        form.setSpacing(8)

        self.raci_phases = QTextEdit()
        self.raci_phases.setPlaceholderText(
            "Requirements\nDesign\nDevelopment\nTesting\nDeployment"
        )
        self.raci_phases.setMaximumHeight(80)
        form.addRow("Project Phases (one per line):", self.raci_phases)

        self.raci_team = QLineEdit()
        self.raci_team.setPlaceholderText("Alice, Bob, Charlie, Diana, Edward")
        form.addRow("Team Members (comma-separated):", self.raci_team)

        btn = QPushButton("Generate RACI Matrix")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_raci)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_status(self):
        group = QGroupBox("E5: Status Report Generator")
        form = QFormLayout()
        form.setSpacing(8)

        self.status_project = QLineEdit()
        self.status_project.setPlaceholderText("Project Alpha")
        form.addRow("Project Name:", self.status_project)

        self.status_period = QLineEdit()
        self.status_period.setPlaceholderText("Week of July 1, 2026")
        form.addRow("Reporting Period:", self.status_period)

        self.status_risks = QCheckBox("Include risk summary")
        self.status_risks.setChecked(True)
        form.addRow("", self.status_risks)

        btn = QPushButton("Generate Status Report")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_status)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _on_gantt(self):
        prompt = build_gantt_prompt(
            tasks_text=self.gantt_tasks.toPlainText().strip(),
            start_date=self.gantt_start.date().toString("yyyy-MM-dd"),
        )
        self.command_sent.emit(prompt)

    def _on_sprint(self):
        prompt = build_sprint_backlog_prompt(
            sprint_name=self.sprint_name.text().strip(),
            duration_weeks=self.sprint_weeks.value(),
            velocity=self.sprint_velocity.value(),
            stories=self.sprint_stories.toPlainText().strip(),
        )
        self.command_sent.emit(prompt)

    def _on_risk(self):
        prompt = build_risk_register_prompt(
            project_name=self.risk_project.toPlainText().strip(),
            num_categories=self.risk_categories.value(),
        )
        self.command_sent.emit(prompt)

    def _on_raci(self):
        prompt = build_raci_prompt(
            phases=self.raci_phases.toPlainText().strip(),
            team_members=self.raci_team.text().strip(),
        )
        self.command_sent.emit(prompt)

    def _on_status(self):
        prompt = build_status_report_prompt(
            task_sheet="Current Sheet",
            project_name=self.status_project.text().strip() or "Project",
            period=self.status_period.text().strip() or "Current Period",
            include_risks=self.status_risks.isChecked(),
        )
        self.command_sent.emit(prompt)
