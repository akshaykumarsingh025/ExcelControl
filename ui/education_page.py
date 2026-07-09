from PyQt6.QtCore import Qt, QDate
from PyQt6.QtWidgets import (
    QVBoxLayout, QGroupBox, QLabel, QTextEdit, QLineEdit,
    QComboBox, QCheckBox, QSpinBox, QDoubleSpinBox, QPushButton,
    QScrollArea, QWidget, QFormLayout, QDateEdit,
)
from ui.workflow_base import WorkflowBase
from core.features import (
    build_gradebook_prompt,
    build_study_planner_prompt,
    build_flashcard_prompt,
)


class EducationPage(WorkflowBase):
    def get_workflow_name(self):
        return "Education"

    def get_workflow_description(self):
        return "Gradebooks, attendance trackers, rubrics, study planners, and flashcard generators."

    def setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(8, 8, 8, 8)

        layout.addWidget(self._build_gradebook())
        layout.addWidget(self._build_attendance())
        layout.addWidget(self._build_rubric())
        layout.addWidget(self._build_study_planner())
        layout.addWidget(self._build_flashcard())
        layout.addStretch()

        scroll.setWidget(container)
        self.get_content_layout().addWidget(scroll)

    def _build_gradebook(self):
        group = QGroupBox("F1: Gradebook Builder")
        form = QFormLayout()
        form.setSpacing(8)

        self.gb_students = QSpinBox()
        self.gb_students.setRange(5, 200)
        self.gb_students.setValue(30)
        form.addRow("Number of Students:", self.gb_students)

        self.gb_assignments = QLineEdit()
        self.gb_assignments.setPlaceholderText("Homework:20, Midterm:30, Final:50")
        form.addRow("Assignments & Weights:", self.gb_assignments)

        self.gb_scale = QComboBox()
        self.gb_scale.addItems(["Letter A-F", "Percentage", "GPA 4.0"])
        form.addRow("Grading Scale:", self.gb_scale)

        self.gb_attendance = QCheckBox("Include attendance column")
        self.gb_attendance.setChecked(False)
        form.addRow("", self.gb_attendance)

        self.gb_passing = QDoubleSpinBox()
        self.gb_passing.setRange(0, 100)
        self.gb_passing.setValue(60)
        self.gb_passing.setSuffix(" %")
        form.addRow("Passing Grade:", self.gb_passing)

        btn = QPushButton("Create Gradebook")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_gradebook)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_attendance(self):
        group = QGroupBox("F2: Attendance Tracker")
        form = QFormLayout()
        form.setSpacing(8)

        self.att_students = QSpinBox()
        self.att_students.setRange(5, 200)
        self.att_students.setValue(30)
        form.addRow("Number of Students:", self.att_students)

        self.att_start = QDateEdit()
        self.att_start.setCalendarPopup(True)
        self.att_start.setDate(QDate.currentDate())
        form.addRow("Start Date:", self.att_start)

        self.att_end = QDateEdit()
        self.att_end.setCalendarPopup(True)
        self.att_end.setDate(QDate.currentDate().addMonths(1))
        form.addRow("End Date:", self.att_end)

        self.att_type = QComboBox()
        self.att_type.addItems(["Daily", "Weekly", "Monthly"])
        form.addRow("Tracking Type:", self.att_type)

        self.att_min = QDoubleSpinBox()
        self.att_min.setRange(0, 100)
        self.att_min.setValue(75)
        self.att_min.setSuffix(" %")
        form.addRow("Minimum Attendance %:", self.att_min)

        btn = QPushButton("Create Attendance Tracker")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_attendance)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_rubric(self):
        group = QGroupBox("F3: Rubric Generator")
        form = QFormLayout()
        form.setSpacing(8)

        self.rubric_type = QLineEdit("Research Paper")
        form.addRow("Assignment Type:", self.rubric_type)

        self.rubric_criteria = QTextEdit()
        self.rubric_criteria.setPlaceholderText(
            "Content Knowledge\nOrganization\nGrammar & Mechanics\n"
            "Citations & References\nCritical Thinking"
        )
        self.rubric_criteria.setMaximumHeight(80)
        form.addRow("Criteria (one per line):", self.rubric_criteria)

        self.rubric_levels = QSpinBox()
        self.rubric_levels.setRange(2, 6)
        self.rubric_levels.setValue(4)
        form.addRow("Number of Levels:", self.rubric_levels)

        self.rubric_points = QSpinBox()
        self.rubric_points.setRange(10, 1000)
        self.rubric_points.setValue(100)
        form.addRow("Max Points:", self.rubric_points)

        btn = QPushButton("Generate Rubric")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_rubric)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_study_planner(self):
        group = QGroupBox("F4: Study Planner")
        form = QFormLayout()
        form.setSpacing(8)

        self.study_subjects = QTextEdit()
        self.study_subjects.setPlaceholderText(
            "Math: Chapters 5-8\nScience: Units 3-4\nHistory: Chapters 7-9"
        )
        self.study_subjects.setMaximumHeight(70)
        form.addRow("Subjects & Topics:", self.study_subjects)

        self.study_exam = QDateEdit()
        self.study_exam.setCalendarPopup(True)
        self.study_exam.setDate(QDate.currentDate().addDays(28))
        form.addRow("Exam Date:", self.study_exam)

        self.study_days = QSpinBox()
        self.study_days.setRange(7, 365)
        self.study_days.setValue(28)
        form.addRow("Study Days Available:", self.study_days)

        self.study_style = QComboBox()
        self.study_style.addItems(["Balanced", "Spaced Repetition", "Block Schedule", "Pomodoro"])
        form.addRow("Study Style:", self.study_style)

        btn = QPushButton("Generate Study Plan")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_study_planner)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_flashcard(self):
        group = QGroupBox("F5: Flashcard Generator")
        form = QFormLayout()
        form.setSpacing(8)

        self.fc_terms = QTextEdit()
        self.fc_terms.setPlaceholderText(
            "Photosynthesis = Process by which plants convert light to energy\n"
            "Mitosis = Cell division producing two identical daughter cells\n"
            "DNA = Deoxyribonucleic acid, carries genetic information"
        )
        self.fc_terms.setMaximumHeight(90)
        form.addRow("Terms & Definitions (term = definition):", self.fc_terms)

        self.fc_random = QCheckBox("Randomize order")
        self.fc_random.setChecked(True)
        form.addRow("", self.fc_random)

        self.fc_quiz = QCheckBox("Create quiz mode (answer hidden)")
        self.fc_quiz.setChecked(False)
        form.addRow("", self.fc_quiz)

        btn = QPushButton("Generate Flashcards")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_flashcard)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _on_gradebook(self):
        prompt = build_gradebook_prompt(
            num_students=self.gb_students.value(),
            assignments=self.gb_assignments.text().strip() or "Homework:20, Midterm:30, Final:50",
            grading_scale=self.gb_scale.currentText(),
            passing_grade=self.gb_passing.value(),
        )
        self.command_sent.emit(prompt)

    def _on_attendance(self):
        prompt = (
            f"Create an attendance tracker for {self.att_students.value()} students "
            f"from {self.att_start.date().toString('yyyy-MM-dd')} "
            f"to {self.att_end.date().toString('yyyy-MM-dd')}. "
            f"Tracking type: {self.att_type.currentText()}. "
            f"Minimum attendance: {self.att_min.value()}%. "
            "Columns: Student Name, Student ID, date columns (P/A/L), Total Present, "
            "Total Absent, Late Count, Attendance %. Use COUNTIF for totals. "
            "Conditional formatting: red <75%, yellow 75-90%, green >90%. "
            "Add data validation dropdown (P/A/L). Bold headers. Apply borders. "
            "Start the data at A1. Use ws for the active sheet. "
            "Respond with ONLY Python code. No explanation, no markdown, no backticks."
        )
        self.command_sent.emit(prompt)

    def _on_rubric(self):
        criteria = self.rubric_criteria.toPlainText().strip()
        prompt = (
            f"Create a grading rubric for: {self.rubric_type.text().strip()}. "
            f"Criteria: {criteria}. "
            f"Number of performance levels: {self.rubric_levels.value()}. "
            f"Maximum points: {self.rubric_points.value()}. "
            "Create a rubric table: rows = criteria, columns = performance levels "
            "(Exemplary/Proficient/Basic/Below Basic). Each cell has description and point range. "
            "Bold headers and criteria. Center-align. Apply borders. Alternate row colors. "
            "Start the data at A1. Use ws for the active sheet. "
            "Respond with ONLY Python code. No explanation, no markdown, no backticks."
        )
        self.command_sent.emit(prompt)

    def _on_study_planner(self):
        prompt = build_study_planner_prompt(
            subjects=self.study_subjects.toPlainText().strip(),
            exam_date=self.study_exam.date().toString("yyyy-MM-dd"),
            study_days=self.study_days.value(),
            style=self.study_style.currentText(),
        )
        self.command_sent.emit(prompt)

    def _on_flashcard(self):
        prompt = build_flashcard_prompt(
            terms_text=self.fc_terms.toPlainText().strip(),
            randomize=self.fc_random.isChecked(),
            quiz_mode=self.fc_quiz.isChecked(),
        )
        self.command_sent.emit(prompt)
