from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QTextCursor
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTextEdit,
    QLineEdit,
    QPushButton,
    QListWidget,
    QFileDialog,
    QApplication,
)


TAG_COLORS = {
    "user": "#89b4fa",
    "ai": "#a6e3a1",
    "code": "#cba6f7",
    "error": "#f38ba8",
    "success": "#a6e3a1",
    "system": "#f9e2af",
    "analysis": "#94e2d5",
    "dryrun": "#fab387",
}


class AgentWorker(QThread):
    response_ready = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, agent, command, context="", analysis_mode=False, images=None):
        super().__init__()
        self.agent = agent
        self.command = command
        self.context = context
        self.analysis_mode = analysis_mode
        self.images = images

    def run(self):
        try:
            if self.analysis_mode:
                result = self.agent.ask_with_context(
                    self.command, self.context, mode="analysis"
                )
            elif self.images:
                result = self.agent.ask(self.command, context=self.context, images=self.images)
            else:
                result = self.agent.ask(self.command, context=self.context)
            self.response_ready.emit(result)
        except Exception as e:
            self.error_occurred.emit(str(e))


class ChatPanel(QWidget):
    command_sent = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._agent = None
        self._excel_controller = None
        self._worker = None
        self._autocomplete_commands = []
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        self.chat_display = QTextEdit()
        self.chat_display.setReadOnly(True)
        self.chat_display.setObjectName("chatDisplay")
        self.chat_display.setStyleSheet(
            "QTextEdit#chatDisplay { background-color: #181825; border: 1px solid #313244; "
            "border-radius: 8px; padding: 8px; font-size: 13px; }"
        )
        layout.addWidget(self.chat_display, stretch=1)

        self.autocomplete_list = QListWidget()
        self.autocomplete_list.setMaximumHeight(160)
        self.autocomplete_list.hide()
        self.autocomplete_list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.autocomplete_list.itemClicked.connect(self._autocomplete_selected)
        layout.addWidget(self.autocomplete_list)

        input_row = QHBoxLayout()
        input_row.setSpacing(6)

        self.image_btn = QPushButton("\U0001f4f7")
        self.image_btn.setFixedWidth(36)
        self.image_btn.setToolTip("Import image for vision analysis")
        self.image_btn.clicked.connect(self._import_image)
        input_row.addWidget(self.image_btn)

        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("Type a command or question...")
        self.input_field.returnPressed.connect(self._send_command)
        self.input_field.textChanged.connect(self._on_text_changed)
        input_row.addWidget(self.input_field, stretch=1)

        self.send_btn = QPushButton("Send")
        self.send_btn.setObjectName("primaryBtn")
        self.send_btn.setFixedWidth(80)
        self.send_btn.clicked.connect(self._send_command)
        input_row.addWidget(self.send_btn)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setFixedWidth(60)
        self.clear_btn.clicked.connect(self._clear_chat)
        input_row.addWidget(self.clear_btn)

        layout.addLayout(input_row)

    def set_agent(self, agent):
        self._agent = agent

    def set_excel_controller(self, controller):
        self._excel_controller = controller

    def set_autocomplete_commands(self, commands):
        self._autocomplete_commands = list(commands)

    def add_message(self, text, tag="system"):
        color = TAG_COLORS.get(tag, "#cdd6f4")
        tag_name = f"tag_{tag}_{id(text)}"

        cursor = self.chat_display.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)

        fmt = cursor.charFormat()
        fmt.setForeground(self._color(color))
        if tag == "code":
            fmt.setFontFamily("Cascadia Code, Fira Code, Consolas, monospace")
            fmt.setFontPointSize(12)
        else:
            fmt.setFontFamily("Segoe UI, Helvetica Neue, Arial, sans-serif")
            fmt.setFontPointSize(13)

        self.chat_display.setTextCursor(cursor)
        cursor.insertText(text + "\n", fmt)
        self.chat_display.setTextCursor(cursor)
        self.chat_display.ensureCursorVisible()

    @staticmethod
    def _color(hex_str):
        from PyQt6.QtGui import QColor

        c = QColor(hex_str)
        if not c.isValid():
            c = QColor("#cdd6f4")
        return c

    def _send_command(self):
        text = self.input_field.text().strip()
        if not text:
            return
        self.input_field.clear()
        self.autocomplete_list.hide()
        self.add_message(f"You: {text}", "user")
        self.command_sent.emit(text)

        if self._agent:
            self._run_agent(text)

    def _run_agent(self, command):
        context = ""
        if self._excel_controller and self._excel_controller.is_connected():
            context = self._excel_controller.get_sheet_context()

        self.send_btn.setEnabled(False)
        self.add_message("Thinking...", "system")

        analysis = getattr(self._agent, "analysis_mode", False) if self._agent else False
        self._worker = AgentWorker(
            self._agent, command, context=context, analysis_mode=analysis
        )
        self._worker.response_ready.connect(self._on_agent_response)
        self._worker.error_occurred.connect(self._on_agent_error)
        self._worker.start()

    def _on_agent_response(self, response):
        self.send_btn.setEnabled(True)
        self._remove_last_thinking()
        self.add_message(f"AI:\n{response}", "ai")
        if self._agent and "```" not in response:
            pass
        tag = "code" if self._looks_like_code(response) else "ai"
        if tag == "code":
            self.add_message(response, "code")

    def _on_agent_error(self, error_msg):
        self.send_btn.setEnabled(True)
        self._remove_last_thinking()
        self.add_message(f"Error: {error_msg}", "error")

    def _remove_last_thinking(self):
        cursor = self.chat_display.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
        selected = cursor.selectedText()
        if "Thinking..." in selected:
            cursor.removeSelectedText()
            cursor.deletePreviousChar()

    @staticmethod
    def _looks_like_code(text):
        lines = text.strip().split("\n")
        code_indicators = ["ws.", "wb.", "import ", "for ", "if ", "= ", ".value", ".formula"]
        code_lines = sum(1 for l in lines if any(ind in l for ind in code_indicators))
        return code_lines > len(lines) * 0.4 if lines else False

    def _clear_chat(self):
        self.chat_display.clear()

    def _import_image(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Image", "", "Images (*.png *.jpg *.jpeg *.bmp *.tiff)"
        )
        if path and self._agent:
            try:
                with open(path, "rb") as f:
                    image_bytes = f.read()
                self.add_message(f"Image loaded: {path}", "system")
                self.add_message("Analyzing image...", "system")
                self.send_btn.setEnabled(False)
                self._worker = AgentWorker(
                    self._agent,
                    "Analyze this image and describe what you see, especially any tables or data.",
                    images=[image_bytes],
                )
                self._worker.response_ready.connect(self._on_agent_response)
                self._worker.error_occurred.connect(self._on_agent_error)
                self._worker.start()
            except Exception as e:
                self.add_message(f"Failed to load image: {e}", "error")

    def _on_text_changed(self, text):
        if not text or not self._autocomplete_commands:
            self.autocomplete_list.hide()
            return

        matches = [c for c in self._autocomplete_commands if text.lower() in c.lower()][:8]
        if not matches:
            self.autocomplete_list.hide()
            return

        self.autocomplete_list.clear()
        self.autocomplete_list.addItems(matches)
        self.autocomplete_list.show()
        self.autocomplete_list.setFixedWidth(self.input_field.width())

    def _autocomplete_selected(self, item):
        self.input_field.setText(item.text())
        self.autocomplete_list.hide()
        self.input_field.setFocus()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.autocomplete_list.hide()
        super().keyPressEvent(event)
