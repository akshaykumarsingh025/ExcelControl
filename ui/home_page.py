from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QPushButton,
    QFileDialog,
    QFrame,
    QScrollArea,
    QSizePolicy,
)

from ui.workflow_base import WorkflowBase
from ui.chat_panel import ChatPanel
from templates import TEMPLATES, TEMPLATE_NAMES


class TemplateCard(QFrame):
    clicked = None

    def __init__(self, name: str, description: str, parent=None):
        super().__init__(parent)
        self.setObjectName("cardFrame")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedHeight(80)
        self._template_name = name
        self.clicked = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)

        name_label = QLabel(name)
        name_label.setStyleSheet("color: #89b4fa; font-weight: bold; font-size: 13px;")
        name_label.setWordWrap(False)
        layout.addWidget(name_label)

        desc_label = QLabel(description[:90] + ("..." if len(description) > 90 else ""))
        desc_label.setStyleSheet("color: #a6adc8; font-size: 11px;")
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.clicked:
            self.clicked(self._template_name)
        super().mousePressEvent(event)


class HomePage(WorkflowBase):
    def get_workflow_name(self) -> str:
        return "ExcelAI Home"

    def get_workflow_description(self) -> str:
        return "AI-powered Excel automation. Open a workbook or start chatting to begin."

    def setup_ui(self):
        content = self.get_content_layout()

        banner = QFrame()
        banner.setStyleSheet(
            "QFrame { background-color: #313244; border-radius: 12px; padding: 20px; }"
        )
        banner_layout = QVBoxLayout(banner)
        banner_layout.setSpacing(8)

        app_name = QLabel("ExcelAI")
        app_name.setStyleSheet(
            "color: #89b4fa; font-size: 32px; font-weight: bold; background: transparent;"
        )
        banner_layout.addWidget(app_name)

        app_ver = QLabel("v2.0 — AI-Powered Excel Controller")
        app_ver.setStyleSheet("color: #a6adc8; font-size: 14px; background: transparent;")
        banner_layout.addWidget(app_ver)

        content.addWidget(banner)

        actions_label = QLabel("Quick Actions")
        actions_label.setStyleSheet("color: #cdd6f4; font-size: 16px; font-weight: bold; margin-top: 12px;")
        content.addWidget(actions_label)

        actions_row = QHBoxLayout()
        actions_row.setSpacing(10)

        self.open_btn = QPushButton("📂 Open Excel File")
        self.open_btn.setObjectName("primaryBtn")
        self.open_btn.setMinimumHeight(40)
        self.open_btn.clicked.connect(self._open_file)
        actions_row.addWidget(self.open_btn)

        self.connect_btn = QPushButton("🔗 Connect to Running Excel")
        self.connect_btn.setMinimumHeight(40)
        self.connect_btn.clicked.connect(self._connect_excel)
        actions_row.addWidget(self.connect_btn)

        self.new_btn = QPushButton("📄 New Workbook")
        self.new_btn.setMinimumHeight(40)
        self.new_btn.clicked.connect(self._new_workbook)
        actions_row.addWidget(self.new_btn)

        actions_row.addStretch()
        content.addLayout(actions_row)

        templates_label = QLabel("Templates")
        templates_label.setStyleSheet("color: #cdd6f4; font-size: 16px; font-weight: bold; margin-top: 12px;")
        content.addWidget(templates_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setMaximumHeight(200)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        templates_widget = QWidget()
        templates_grid = QGridLayout(templates_widget)
        templates_grid.setSpacing(8)

        cols = 3
        for i, name in enumerate(TEMPLATE_NAMES):
            tmpl = TEMPLATES[name]
            card = TemplateCard(name, tmpl["description"])
            card.clicked = self._on_template_clicked
            templates_grid.addWidget(card, i // cols, i % cols)

        scroll.setWidget(templates_widget)
        content.addWidget(scroll)

        self._chat_panel = ChatPanel()
        self._chat_panel.setMinimumHeight(180)
        self._chat_panel.command_sent.connect(self.command_sent.emit)
        content.addWidget(self._chat_panel, stretch=1)

    def get_chat_panel(self):
        return self._chat_panel

    def _open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xlsm *.xlsb *.xls)"
        )
        if path:
            self.command_sent.emit(f"__open_file__{path}")

    def _connect_excel(self):
        self.command_sent.emit("__connect_excel__")

    def _new_workbook(self):
        self.command_sent.emit("__new_workbook__")

    def _on_template_clicked(self, name):
        self.command_sent.emit(f"__template__{name}")
