from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QSplitter,
)


class WorkflowBase(QWidget):
    command_sent = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._chat_panel = None
        self._setup_base_ui()
        self.setup_ui()

    def _setup_base_ui(self):
        self._main_layout = QVBoxLayout(self)
        self._main_layout.setContentsMargins(16, 16, 16, 8)
        self._main_layout.setSpacing(8)

        title_row = QHBoxLayout()
        self._title_label = QLabel(self.get_workflow_name())
        self._title_label.setObjectName("headingLabel")
        title_row.addWidget(self._title_label)
        title_row.addStretch()
        self._main_layout.addLayout(title_row)

        self._desc_label = QLabel(self.get_workflow_description())
        self._desc_label.setObjectName("subheadingLabel")
        self._desc_label.setWordWrap(True)
        self._main_layout.addWidget(self._desc_label)

        separator = QLabel()
        separator.setFixedHeight(1)
        separator.setStyleSheet("background-color: #313244;")
        self._main_layout.addWidget(separator)

        self._content_widget = QWidget()
        self._content_layout = QVBoxLayout(self._content_widget)
        self._content_layout.setContentsMargins(0, 0, 0, 0)
        self._main_layout.addWidget(self._content_widget, stretch=1)

    def setup_ui(self):
        pass

    def get_workflow_name(self) -> str:
        return "Workflow"

    def get_workflow_description(self) -> str:
        return ""

    def set_chat_panel(self, panel):
        self._chat_panel = panel
        if self._chat_panel:
            self._chat_panel.command_sent.connect(self.command_sent.emit)

    def get_content_layout(self):
        return self._content_layout

    def get_content_widget(self):
        return self._content_widget
