from PyQt6.QtCore import Qt, QTime
from PyQt6.QtWidgets import (
    QVBoxLayout, QGroupBox, QLabel, QLineEdit,
    QComboBox, QCheckBox, QSpinBox, QPushButton,
    QScrollArea, QWidget, QFormLayout, QTimeEdit,
)
from ui.workflow_base import WorkflowBase
from core.features import (
    build_patient_schedule_prompt,
    build_medication_tracker_prompt,
    build_clinical_cleaner_prompt,
)


class HealthcarePage(WorkflowBase):
    def get_workflow_name(self):
        return "Healthcare"

    def get_workflow_description(self):
        return "Patient scheduling, medication tracking, and clinical data cleaning."

    def setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(8, 8, 8, 8)

        layout.addWidget(self._build_patient_schedule())
        layout.addWidget(self._build_medication_tracker())
        layout.addWidget(self._build_clinical_cleaner())
        layout.addStretch()

        scroll.setWidget(container)
        self.get_content_layout().addWidget(scroll)

    def _build_patient_schedule(self):
        group = QGroupBox("I1: Patient Schedule Builder")
        form = QFormLayout()
        form.setSpacing(8)

        self.sched_start = QTimeEdit()
        self.sched_start.setTime(QTime(9, 0))
        form.addRow("Clinic Start Time:", self.sched_start)

        self.sched_end = QTimeEdit()
        self.sched_end.setTime(QTime(17, 0))
        form.addRow("Clinic End Time:", self.sched_end)

        self.sched_duration = QSpinBox()
        self.sched_duration.setRange(5, 60)
        self.sched_duration.setValue(15)
        self.sched_duration.setSuffix(" min")
        form.addRow("Appointment Duration:", self.sched_duration)

        self.sched_lunch_start = QTimeEdit()
        self.sched_lunch_start.setTime(QTime(13, 0))
        form.addRow("Lunch Break Start:", self.sched_lunch_start)

        self.sched_lunch_duration = QSpinBox()
        self.sched_lunch_duration.setRange(15, 120)
        self.sched_lunch_duration.setValue(60)
        self.sched_lunch_duration.setSuffix(" min")
        form.addRow("Lunch Duration:", self.sched_lunch_duration)

        self.sched_max = QSpinBox()
        self.sched_max.setRange(1, 100)
        self.sched_max.setValue(24)
        form.addRow("Max Patients per Day:", self.sched_max)

        self.sched_buffer = QCheckBox("Include buffer time between appointments")
        self.sched_buffer.setChecked(False)
        form.addRow("", self.sched_buffer)

        btn = QPushButton("Create Schedule Template")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_patient_schedule)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_medication_tracker(self):
        group = QGroupBox("I2: Medication Tracker")
        form = QFormLayout()
        form.setSpacing(8)

        self.med_count = QSpinBox()
        self.med_count.setRange(1, 20)
        self.med_count.setValue(5)
        form.addRow("Number of Medications:", self.med_count)

        self.med_duration = QSpinBox()
        self.med_duration.setRange(7, 365)
        self.med_duration.setValue(30)
        self.med_duration.setSuffix(" days")
        form.addRow("Tracking Duration:", self.med_duration)

        self.med_side_effects = QCheckBox("Include side effects log")
        self.med_side_effects.setChecked(True)
        form.addRow("", self.med_side_effects)

        self.med_refill = QCheckBox("Include refill reminder column")
        self.med_refill.setChecked(True)
        form.addRow("", self.med_refill)

        self.med_time_of_day = QCheckBox("Time-of-day tracking (Morning/Afternoon/Evening/Night)")
        self.med_time_of_day.setChecked(True)
        form.addRow("", self.med_time_of_day)

        btn = QPushButton("Create Medication Tracker")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_medication_tracker)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_clinical_cleaner(self):
        group = QGroupBox("I3: Clinical Data Cleaner")
        form = QFormLayout()
        form.setSpacing(8)

        self.clinic_dates = QCheckBox("Fix date formats (standardize to ISO 8601)")
        self.clinic_dates.setChecked(True)
        form.addRow("", self.clinic_dates)

        self.clinic_bp = QCheckBox("Normalize blood pressure readings (e.g., 120/80 → 120/80 mmHg)")
        self.clinic_bp.setChecked(True)
        form.addRow("", self.clinic_bp)

        self.clinic_phones = QCheckBox("Standardize phone numbers")
        self.clinic_phones.setChecked(True)
        form.addRow("", self.clinic_phones)

        self.clinic_missing = QCheckBox("Flag missing values")
        self.clinic_missing.setChecked(True)
        form.addRow("", self.clinic_missing)

        self.clinic_dupes = QCheckBox("Remove duplicate records")
        self.clinic_dupes.setChecked(False)
        form.addRow("", self.clinic_dupes)

        self.clinic_units = QCheckBox("Normalize measurement units (kg, cm, Celsius)")
        self.clinic_units.setChecked(False)
        form.addRow("", self.clinic_units)

        btn = QPushButton("Clean Clinical Data")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_clinical_cleaner)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _on_patient_schedule(self):
        buffer_min = 5 if self.sched_buffer.isChecked() else 0
        prompt = build_patient_schedule_prompt(
            start=self.sched_start.time().toString("HH:mm"),
            end=self.sched_end.time().toString("HH:mm"),
            duration=self.sched_duration.value(),
            lunch_start=self.sched_lunch_start.time().toString("HH:mm"),
            lunch_duration=self.sched_lunch_duration.value(),
            max_patients=self.sched_max.value(),
            buffer=buffer_min,
        )
        self.command_sent.emit(prompt)

    def _on_medication_tracker(self):
        time_of_day = "Morning/Afternoon/Evening/Night" if self.med_time_of_day.isChecked() else "Daily"
        prompt = build_medication_tracker_prompt(
            num_meds=self.med_count.value(),
            duration=self.med_duration.value(),
            side_effects=self.med_side_effects.isChecked(),
            refill=self.med_refill.isChecked(),
            time_of_day=time_of_day,
        )
        self.command_sent.emit(prompt)

    def _on_clinical_cleaner(self):
        options = {
            "fix_dates": self.clinic_dates.isChecked(),
            "normalize_units": self.clinic_units.isChecked(),
            "remove_phi": False,
            "validate_ranges": self.clinic_missing.isChecked(),
            "fill_missing": False,
            "standardize_codes": False,
        }
        if self.clinic_bp.isChecked():
            options["normalize_units"] = True
        if self.clinic_dupes.isChecked():
            options["fill_missing"] = True

        prompt = build_clinical_cleaner_prompt(options)
        self.command_sent.emit(prompt)
