from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QStackedWidget,
    QComboBox,
    QToolBar,
    QStatusBar,
    QSplitter,
    QFileDialog,
    QCheckBox,
    QMessageBox,
    QApplication,
)

from core.excel_controller import ExcelController
from core.agent import ExcelAgent, get_available_models
from core.code_validator import validate_code
from core.dry_run import analyze_code
from history import CommandHistory
from templates import TEMPLATES, TEMPLATE_NAMES
from prompts import COMMON_COMMANDS
from session_manager import save_session, load_session, clear_session

from ui.chat_panel import ChatPanel
from ui.code_editor import CodeEditor
from ui.sheet_view import SheetView
from ui.home_page import HomePage
from ui.settings_page import SettingsPage
from ui.workflow_base import WorkflowBase
from ui.sheet_view_page import SheetViewPage
from ui.code_editor_page import CodeEditorPage
from ui.data_tools_page import DataToolsPage
from ui.qa_page import QAPage
from ui.finance_page import FinancePage
from ui.hr_page import HRPage
from ui.marketing_page import MarketingPage
from ui.project_mgmt_page import ProjectMgmtPage
from ui.education_page import EducationPage
from ui.operations_page import OperationsPage
from ui.real_estate_page import RealEstatePage
from ui.healthcare_page import HealthcarePage


PAGES = [
    ("🏠", "Home"),
    ("📊", "Sheet View"),
    ("💻", "Code Editor"),
    ("🔧", "Data Tools"),
    ("🧪", "QA Testing"),
    ("💰", "Finance"),
    ("👥", "HR"),
    ("📈", "Marketing"),
    ("📋", "Project Mgmt"),
    ("📚", "Education"),
    ("🏭", "Operations"),
    ("🏠", "Real Estate"),
    ("🏥", "Healthcare"),
    ("⚙️", "Settings"),
]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ExcelAI — AI-Powered Excel Controller")
        self.setMinimumSize(1100, 700)
        self.resize(1400, 850)

        self.excel_controller = ExcelController()
        self.agent = ExcelAgent()
        self.history = CommandHistory()
        self._current_file = ""

        self._sidebar_buttons = []
        self._workflow_pages = {}
        self._page_chat_panels = []

        self._build_ui()
        self._connect_signals()
        self._restore_session()
        self._update_connection_status()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(200)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(8, 12, 8, 12)
        sidebar_layout.setSpacing(2)

        logo_label = QLabel("  ExcelAI")
        logo_label.setStyleSheet(
            "color: #89b4fa; font-size: 18px; font-weight: bold; padding: 8px 4px 16px 4px;"
        )
        sidebar_layout.addWidget(logo_label)

        for i, (icon, name) in enumerate(PAGES):
            btn = QPushButton(f"  {icon}  {name}")
            btn.setObjectName("sidebarBtn")
            btn.setCheckable(True)
            btn.setMinimumHeight(36)
            btn.clicked.connect(lambda checked, idx=i: self._switch_page(idx))
            sidebar_layout.addWidget(btn)
            self._sidebar_buttons.append(btn)

        sidebar_layout.addStretch()

        version_label = QLabel("v2.0")
        version_label.setStyleSheet("color: #585b70; font-size: 11px; padding: 4px;")
        sidebar_layout.addWidget(version_label)

        main_layout.addWidget(sidebar)

        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        self._build_toolbar()
        right_layout.addWidget(self._toolbar)

        self._stacked = QStackedWidget()
        right_layout.addWidget(self._stacked, stretch=1)

        main_layout.addWidget(right_container, stretch=1)

        self._create_pages()

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self._status_label = QLabel("Disconnected")
        self.status_bar.addWidget(self._status_label)

    def _build_toolbar(self):
        self._toolbar = QToolBar()
        self._toolbar.setMovable(False)
        self._toolbar.setFixedHeight(44)

        open_action = QAction("📂 Open", self)
        open_action.setToolTip("Open Excel file")
        open_action.triggered.connect(self._open_file)
        self._toolbar.addAction(open_action)

        connect_action = QAction("🔗 Connect", self)
        connect_action.setToolTip("Connect to running Excel")
        connect_action.triggered.connect(self._connect_excel)
        self._toolbar.addAction(connect_action)

        new_action = QAction("📄 New", self)
        new_action.setToolTip("Create new workbook")
        new_action.triggered.connect(self._new_workbook)
        self._toolbar.addAction(new_action)

        self._toolbar.addSeparator()

        undo_action = QAction("↩️ Undo", self)
        undo_action.setToolTip("Undo last change")
        undo_action.triggered.connect(self._undo)
        self._toolbar.addAction(undo_action)

        redo_action = QAction("↪️ Redo", self)
        redo_action.setToolTip("Redo last change")
        redo_action.triggered.connect(self._redo)
        self._toolbar.addAction(redo_action)

        self._toolbar.addSeparator()

        spacer = QWidget()
        spacer.setMinimumWidth(20)
        self._toolbar.addWidget(spacer)

        model_label = QLabel("Model: ")
        model_label.setStyleSheet("color: #a6adc8; font-size: 12px;")
        self._toolbar.addWidget(model_label)

        self._model_combo = QComboBox()
        self._model_combo.setMinimumWidth(180)
        models = get_available_models()
        self._model_combo.addItems(models)
        self._model_combo.currentTextChanged.connect(self._on_model_changed)
        self._toolbar.addWidget(self._model_combo)

        self._toolbar.addSeparator()

        self._auto_run_cb = QCheckBox("Auto-Run")
        self._auto_run_cb.setToolTip("Execute AI-generated code automatically")
        self._toolbar.addWidget(self._auto_run_cb)

        self._dry_run_cb = QCheckBox("Dry-Run")
        self._dry_run_cb.setToolTip("Preview changes before executing")
        self._toolbar.addWidget(self._dry_run_cb)

        self._analysis_cb = QCheckBox("Analysis")
        self._analysis_cb.setToolTip("Analysis mode: AI answers questions instead of writing code")
        self._toolbar.addWidget(self._analysis_cb)

    def _create_pages(self):
        self._sheet_view = SheetView()

        page_builders = [
            ("Home", self._make_home),
            ("Sheet View", self._make_sheet_view),
            ("Code Editor", self._make_code_editor),
            ("Data Tools", self._make_data_tools),
            ("QA Testing", lambda: self._make_workflow("QA Testing", QAPage)),
            ("Finance", lambda: self._make_workflow("Finance", FinancePage)),
            ("HR", lambda: self._make_workflow("HR", HRPage)),
            ("Marketing", lambda: self._make_workflow("Marketing", MarketingPage)),
            ("Project Mgmt", lambda: self._make_workflow("Project Mgmt", ProjectMgmtPage)),
            ("Education", lambda: self._make_workflow("Education", EducationPage)),
            ("Operations", lambda: self._make_workflow("Operations", OperationsPage)),
            ("Real Estate", lambda: self._make_workflow("Real Estate", RealEstatePage)),
            ("Healthcare", lambda: self._make_workflow("Healthcare", HealthcarePage)),
            ("Settings", self._make_settings),
        ]

        for name, builder in page_builders:
            page = builder()
            self._stacked.addWidget(page)

        self._switch_page(0)

    def _make_home(self):
        self._home_page = HomePage()
        self._workflow_pages["Home"] = self._home_page
        self._page_chat_panels.append(self._home_page.get_chat_panel())
        return self._home_page

    def _make_sheet_view(self):
        self._sheet_view_page = SheetViewPage()
        self._sheet_view_page.set_excel_controller(self.excel_controller)
        self._sheet_view_page.set_agent(self.agent)
        self._workflow_pages["Sheet View"] = self._sheet_view_page
        self._page_chat_panels.append(self._sheet_view_page.chat_panel)
        return self._sheet_view_page

    def _make_code_editor(self):
        self._code_editor_page = CodeEditorPage()
        self._code_editor_page.set_excel_controller(self.excel_controller)
        self._code_editor_page.set_agent(self.agent)
        self._workflow_pages["Code Editor"] = self._code_editor_page
        self._page_chat_panels.append(self._code_editor_page.chat_panel)
        return self._code_editor_page

    def _make_data_tools(self):
        self._data_tools_page = DataToolsPage()
        self._data_tools_page.set_excel_controller(self.excel_controller)
        self._data_tools_page.set_agent(self.agent)
        self._workflow_pages["Data Tools"] = self._data_tools_page
        self._page_chat_panels.append(self._data_tools_page.chat_panel)
        return self._data_tools_page

    def _make_workflow(self, sidebar_name, cls):
        page = cls()
        page.agent = self.agent
        page.excel = self.excel_controller
        self._workflow_pages[sidebar_name] = page
        return page

    def _make_settings(self):
        self._settings_page = SettingsPage()
        self._workflow_pages["Settings"] = self._settings_page
        return self._settings_page

    def _connect_signals(self):
        for name, page in self._workflow_pages.items():
            page.command_sent.connect(self._handle_command)

        self._code_editor_page.code_editor.execute_requested = self._execute_code
        self._code_editor_page.code_editor.validate_requested = self._validate_code
        self._code_editor_page.code_editor.dry_run_requested = self._dry_run_code

        self._auto_run_cb.toggled.connect(self._on_auto_run_changed)
        self._dry_run_cb.toggled.connect(self._on_dry_run_changed)
        self._analysis_cb.toggled.connect(self._on_analysis_changed)

    def _switch_page(self, index):
        self._stacked.setCurrentIndex(index)
        for i, btn in enumerate(self._sidebar_buttons):
            btn.setProperty("active", i == index)
            btn.setStyle(btn.style())
            btn.setChecked(i == index)

        page_name = PAGES[index][1] if index < len(PAGES) else ""
        if page_name == "Sheet View":
            self._sheet_view_page.refresh_data()
        elif page_name == "Code Editor":
            self._code_editor_page.code_editor._update_line_numbers()

    def _handle_command(self, command: str):
        if command.startswith("__open_file__"):
            path = command[len("__open_file__"):]
            self._open_file_path(path)
            return

        if command == "__connect_excel__":
            self._connect_excel()
            return

        if command == "__new_workbook__":
            self._new_workbook()
            return

        if command.startswith("__template__"):
            name = command[len("__template__"):]
            self._apply_template(name)
            return

        self._send_to_agent(command)

    def _send_to_agent(self, command: str):
        chat = self._get_active_chat()
        if not chat:
            return

        if not chat._agent:
            chat.set_agent(self.agent)
        if not chat._excel_controller:
            chat.set_excel_controller(self.excel_controller)

        self.agent.set_analysis_mode(self._analysis_cb.isChecked())
        all_commands = list(set(COMMON_COMMANDS + self.history.get_commands()))
        chat.set_autocomplete_commands(all_commands)

        if self._auto_run_cb.isChecked():
            chat.add_message("Auto-run enabled — code will execute automatically.", "system")

    def _get_active_chat(self) -> ChatPanel:
        index = self._stacked.currentIndex()
        page_name = PAGES[index][1] if index < len(PAGES) else ""
        page = self._workflow_pages.get(page_name)
        if page and hasattr(page, 'chat_panel'):
            return page.chat_panel
        if page and hasattr(page, 'get_chat_panel'):
            return page.get_chat_panel()
        return None

    def _open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xlsm *.xlsb *.xls)"
        )
        if path:
            self._open_file_path(path)

    def _open_file_path(self, path: str):
        ok, msg = self.excel_controller.connect_or_open(filepath=path)
        if ok:
            self._current_file = path
            self._update_connection_status()
            self._notify_all_chats(f"Connected: {msg}", "success")
            self._sheet_view.refresh_data()
        else:
            self._notify_all_chats(f"Error: {msg}", "error")

    def _connect_excel(self):
        ok, msg = self.excel_controller.connect_or_open()
        if ok:
            self._update_connection_status()
            self._notify_all_chats(f"Connected: {msg}", "success")
            self._sheet_view.refresh_data()
        else:
            self._notify_all_chats(f"Error: {msg}", "error")

    def _new_workbook(self):
        ok, msg = self.excel_controller.connect_or_open()
        if ok:
            self._update_connection_status()
            self._notify_all_chats(f"New workbook created: {msg}", "success")
            self._sheet_view.refresh_data()
        else:
            self._notify_all_chats(f"Error: {msg}", "error")

    def _undo(self):
        if self.excel_controller.is_connected():
            ok, msg = self.excel_controller.undo()
            self._notify_all_chats(f"Undo: {msg}", "system" if ok else "error")
            self._sheet_view.refresh_data()

    def _redo(self):
        if self.excel_controller.is_connected():
            ok, msg = self.excel_controller.redo()
            self._notify_all_chats(f"Redo: {msg}", "system" if ok else "error")
            self._sheet_view.refresh_data()

    def _apply_template(self, name: str):
        if name not in TEMPLATES:
            self._notify_all_chats(f"Template not found: {name}", "error")
            return

        if not self.excel_controller.is_connected():
            self._notify_all_chats("Connect to Excel first before applying templates.", "error")
            return

        code = TEMPLATES[name]["code"]
        self._code_editor_page.set_code(code)
        self._notify_all_chats(f"Template loaded: {name}. Switch to Code Editor to review and execute.", "system")
        self._switch_page(2)

    def _execute_code(self):
        code = self._code_editor_page.get_code()
        if not code.strip():
            self._code_editor_page.code_editor.set_status("No code to execute", False)
            return

        if not self.excel_controller.is_connected():
            self._code_editor_page.code_editor.set_status("Not connected to Excel", False)
            return

        dry_run = self._dry_run_cb.isChecked()
        ok, msg = self.excel_controller.execute(code, dry_run_enabled=dry_run)
        self._code_editor_page.code_editor.set_status(msg, ok)

        tag = "success" if ok else "error"
        chat = self._get_active_chat()
        if chat:
            chat.add_message(f"Execute: {msg}", tag)

        self.history.add("code_execution", code, ok)
        self._sheet_view.refresh_data()

    def _validate_code(self):
        code = self._code_editor_page.get_code()
        if not code.strip():
            self._code_editor_page.code_editor.set_status("No code to validate", False)
            return

        result = validate_code(code)
        if result.is_safe:
            self._code_editor_page.code_editor.set_status("Validation passed ✓", True)
        else:
            issues = "\n".join(result.issues)
            self._code_editor_page.code_editor.set_status(f"Blocked: {issues}", False)

    def _dry_run_code(self):
        code = self._code_editor_page.get_code()
        if not code.strip():
            self._code_editor_page.code_editor.set_status("No code to analyze", False)
            return

        result = analyze_code(code)
        summary = result.summary()
        self._code_editor_page.code_editor.set_status("Dry run complete", True)

        chat = self._get_active_chat()
        if chat:
            chat.add_message(summary, "dryrun")

    def _on_model_changed(self, model: str):
        self.agent.set_model(model)

    def _on_auto_run_changed(self, enabled: bool):
        self.agent.set_analysis_mode(self._analysis_cb.isChecked())

    def _on_dry_run_changed(self, enabled: bool):
        pass

    def _on_analysis_changed(self, enabled: bool):
        self.agent.set_analysis_mode(enabled)

    def _update_connection_status(self):
        if self.excel_controller.is_connected():
            sheet = self.excel_controller.get_current_sheet_name()
            wb_name = ""
            try:
                wb_name = self.excel_controller.wb.name
            except Exception:
                pass
            self._status_label.setText(f"✅ Connected: {wb_name} — {sheet}")
            self._status_label.setStyleSheet("color: #a6e3a1;")
        else:
            self._status_label.setText("❌ Disconnected")
            self._status_label.setStyleSheet("color: #f38ba8;")

    def _notify_all_chats(self, message: str, tag: str = "system"):
        for chat in self._page_chat_panels:
            if chat:
                chat.add_message(message, tag)

    def _restore_session(self):
        session = load_session()
        if session:
            model = session.get("model", "")
            if model:
                idx = self._model_combo.findText(model)
                if idx >= 0:
                    self._model_combo.setCurrentIndex(idx)
                else:
                    self._model_combo.addItem(model)
                    self._model_combo.setCurrentText(model)
                self.agent.set_model(model)

            self._auto_run_cb.setChecked(session.get("auto_run", False))
            self._dry_run_cb.setChecked(session.get("dry_run", False))
            self._analysis_cb.setChecked(session.get("analysis_mode", False))

            conv_history = session.get("conversation_history", [])
            if conv_history:
                self.agent.set_history(conv_history)

            self._current_file = session.get("last_file", "")
            self._settings_page.set_model(model)
            self._settings_page.set_auto_run(session.get("auto_run", False))
            self._settings_page.set_dry_run(session.get("dry_run", False))
            self._settings_page.set_analysis_mode(session.get("analysis_mode", False))

    def closeEvent(self, event):
        save_session(
            conversation_history=self.agent.get_history(),
            model=self.agent.model,
            auto_run=self._auto_run_cb.isChecked(),
            dry_run=self._dry_run_cb.isChecked(),
            analysis_mode=self._analysis_cb.isChecked(),
            last_file=self._current_file,
        )
        self.excel_controller.cleanup()
        event.accept()
