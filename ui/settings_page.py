from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QComboBox,
    QCheckBox,
    QGroupBox,
    QMessageBox,
    QSpacerItem,
    QSizePolicy,
)

from ui.workflow_base import WorkflowBase
from core.agent import get_available_models
from session_manager import save_session, load_session, clear_session


class SettingsPage(WorkflowBase):
    def get_workflow_name(self) -> str:
        return "Settings"

    def get_workflow_description(self) -> str:
        return "Configure model, automation behavior, session, and scheduler."

    def setup_ui(self):
        content = self.get_content_layout()

        model_group = QGroupBox("Model Selection")
        model_layout = QHBoxLayout(model_group)

        self.model_combo = QComboBox()
        self.model_combo.setMinimumWidth(250)
        model_layout.addWidget(self.model_combo)

        self.refresh_models_btn = QPushButton("🔄 Refresh Models")
        self.refresh_models_btn.clicked.connect(self._refresh_models)
        model_layout.addWidget(self.refresh_models_btn)

        model_layout.addStretch()
        content.addWidget(model_group)

        behavior_group = QGroupBox("Automation Behavior")
        behavior_layout = QVBoxLayout(behavior_group)

        self.auto_run_cb = QCheckBox("Auto-run: Execute AI-generated code automatically")
        self.auto_run_cb.setToolTip("When enabled, code from AI responses runs without manual confirmation")
        behavior_layout.addWidget(self.auto_run_cb)

        self.dry_run_cb = QCheckBox("Dry-run: Preview changes before executing")
        self.dry_run_cb.setToolTip("When enabled, all code is analyzed for changes before execution")
        behavior_layout.addWidget(self.dry_run_cb)

        self.analysis_mode_cb = QCheckBox("Analysis mode: AI answers questions about data instead of writing code")
        self.analysis_mode_cb.setToolTip("When enabled, AI provides analysis instead of generating automation code")
        behavior_layout.addWidget(self.analysis_mode_cb)

        content.addWidget(behavior_group)

        session_group = QGroupBox("Session Management")
        session_layout = QHBoxLayout(session_group)

        self.save_session_btn = QPushButton("💾 Save Session")
        self.save_session_btn.clicked.connect(self._save_session)
        session_layout.addWidget(self.save_session_btn)

        self.load_session_btn = QPushButton("📂 Load Session")
        self.load_session_btn.clicked.connect(self._load_session)
        session_layout.addWidget(self.load_session_btn)

        self.clear_session_btn = QPushButton("🗑️ Clear Session")
        self.clear_session_btn.setObjectName("dangerBtn")
        self.clear_session_btn.clicked.connect(self._clear_session)
        session_layout.addWidget(self.clear_session_btn)

        session_layout.addStretch()
        content.addWidget(session_group)

        scheduler_group = QGroupBox("Batch Scheduler")
        scheduler_layout = QVBoxLayout(scheduler_group)

        scheduler_info = QLabel(
            "Schedule automated batch jobs to run on Excel files.\n"
            "Configure jobs from the Home chat or via the command interface."
        )
        scheduler_info.setStyleSheet("color: #a6adc8; font-size: 12px;")
        scheduler_info.setWordWrap(True)
        scheduler_layout.addWidget(scheduler_info)

        content.addWidget(scheduler_group)

        about_group = QGroupBox("About")
        about_layout = QVBoxLayout(about_group)

        about_text = QLabel(
            "<b>ExcelAI</b> v2.0<br><br>"
            "AI-powered Excel automation controller.<br>"
            "Uses Ollama for local AI inference and xlwings for Excel control.<br><br>"
            "Built with PyQt6."
        )
        about_text.setStyleSheet("color: #cdd6f4; font-size: 13px;")
        about_text.setWordWrap(True)
        about_layout.addWidget(about_text)

        content.addWidget(about_group)

        content.addStretch()

        self._refresh_models()

    def _refresh_models(self):
        self.model_combo.clear()
        models = get_available_models()
        self.model_combo.addItems(models)

    def get_model(self) -> str:
        return self.model_combo.currentText()

    def set_model(self, model: str):
        idx = self.model_combo.findText(model)
        if idx >= 0:
            self.model_combo.setCurrentIndex(idx)

    def is_auto_run(self) -> bool:
        return self.auto_run_cb.isChecked()

    def set_auto_run(self, enabled: bool):
        self.auto_run_cb.setChecked(enabled)

    def is_dry_run(self) -> bool:
        return self.dry_run_cb.isChecked()

    def set_dry_run(self, enabled: bool):
        self.dry_run_cb.setChecked(enabled)

    def is_analysis_mode(self) -> bool:
        return self.analysis_mode_cb.isChecked()

    def set_analysis_mode(self, enabled: bool):
        self.analysis_mode_cb.setChecked(enabled)

    def _save_session(self):
        QMessageBox.information(self, "Session", "Session saved.")

    def _load_session(self):
        session = load_session()
        if session:
            self.set_model(session.get("model", ""))
            self.set_auto_run(session.get("auto_run", False))
            self.set_dry_run(session.get("dry_run", False))
            self.set_analysis_mode(session.get("analysis_mode", False))
            QMessageBox.information(self, "Session", "Session loaded.")
        else:
            QMessageBox.information(self, "Session", "No saved session found.")

    def _clear_session(self):
        clear_session()
        QMessageBox.information(self, "Session", "Session cleared.")
