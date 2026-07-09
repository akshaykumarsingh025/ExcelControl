import keyword
import builtins

from PyQt6.QtCore import Qt, QRect
from PyQt6.QtGui import (
    QColor,
    QTextCharFormat,
    QFont,
    QSyntaxHighlighter,
    QTextDocument,
)
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPlainTextEdit,
    QPushButton,
    QLabel,
    QTextEdit,
)

CATPPUCCIN = {
    "keyword": "#cba6f7",
    "builtin": "#fab387",
    "string": "#a6e3a1",
    "comment": "#6c7086",
    "number": "#fab387",
    "decorator": "#f9e2af",
    "self": "#f38ba8",
    "operator": "#89b4fa",
    "bracket": "#94e2d5",
    "default": "#cdd6f4",
}


class PythonHighlighter(QSyntaxHighlighter):
    def __init__(self, document: QTextDocument):
        super().__init__(document)
        self._rules = []

        keyword_fmt = QTextCharFormat()
        keyword_fmt.setForeground(QColor(CATPPUCCIN["keyword"]))
        keyword_fmt.setFontWeight(QFont.Weight.Bold)
        for kw in keyword.kwlist:
            self._rules.append((rf"\b{kw}\b", keyword_fmt))

        builtin_fmt = QTextCharFormat()
        builtin_fmt.setForeground(QColor(CATPPUCCIN["builtin"]))
        builtin_names = [name for name in dir(builtins) if not name.startswith("_")]
        for bn in builtin_names:
            self._rules.append((rf"\b{bn}\b", builtin_fmt))

        string_fmt = QTextCharFormat()
        string_fmt.setForeground(QColor(CATPPUCCIN["string"]))
        self._rules.append((r'"[^"\\]*(\\.[^"\\]*)*"', string_fmt))
        self._rules.append((r"'[^'\\]*(\\.[^'\\]*)*'", string_fmt))
        self._rules.append((r'"""[^"]*"""', string_fmt))
        self._rules.append((r"'''[^']*'''", string_fmt))

        comment_fmt = QTextCharFormat()
        comment_fmt.setForeground(QColor(CATPPUCCIN["comment"]))
        comment_fmt.setFontItalic(True)
        self._rules.append((r"#[^\n]*", comment_fmt))

        number_fmt = QTextCharFormat()
        number_fmt.setForeground(QColor(CATPPUCCIN["number"]))
        self._rules.append((r"\b[0-9]+(\.[0-9]+)?\b", number_fmt))
        self._rules.append((r"\b0[xX][0-9a-fA-F]+\b", number_fmt))

        decorator_fmt = QTextCharFormat()
        decorator_fmt.setForeground(QColor(CATPPUCCIN["decorator"]))
        self._rules.append((r"@[a-zA-Z_]\w*", decorator_fmt))

        self_fmt = QTextCharFormat()
        self_fmt.setForeground(QColor(CATPPUCCIN["self"]))
        self_fmt.setFontWeight(QFont.Weight.Bold)
        self._rules.append((r"\bself\b", self_fmt))

    def highlightBlock(self, text: str):
        import re

        for pattern, fmt in self._rules:
            for match in re.finditer(pattern, text):
                start = match.start()
                end = match.end()
                self.setFormat(start, end - start, fmt)


class LineNumberArea(QPlainTextEdit):
    def __init__(self, editor: QPlainTextEdit):
        super().__init__(editor)
        self.editor = editor
        self.setReadOnly(True)
        self.setStyleSheet(
            "background-color: #11111b; color: #6c7086; border: none; "
            "font-family: 'Cascadia Code', 'Fira Code', 'Consolas', monospace; "
            "font-size: 13px;"
        )
        self.setMaximumWidth(48)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)

    def update_line_numbers(self):
        block_count = self.editor.document().blockCount()
        lines = [str(i) for i in range(1, block_count + 1)]
        self.setPlainText("\n".join(lines))
        self.verticalScrollBar().setValue(self.editor.verticalScrollBar().value())


class CodeEditor(QWidget):
    execute_requested = None
    validate_requested = None
    dry_run_requested = None

    def __init__(self, parent=None):
        super().__init__(parent)
        self.execute_requested = None
        self.validate_requested = None
        self.dry_run_requested = None
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        editor_layout = QHBoxLayout()
        editor_layout.setSpacing(0)

        self.line_numbers = LineNumberArea(None)

        self.code_edit = QPlainTextEdit()
        self.code_edit.setPlaceholderText("# Write Python code here...\n# Use ws for the active sheet, wb for the workbook")
        self.highlighter = PythonHighlighter(self.code_edit.document())
        self.code_edit.blockCountChanged.connect(self._update_line_numbers)
        self.code_edit.verticalScrollBar().valueChanged.connect(
            self.line_numbers.verticalScrollBar().setValue
        )

        self.line_numbers.editor = self.code_edit

        editor_layout.addWidget(self.line_numbers)
        editor_layout.addWidget(self.code_edit, stretch=1)

        layout.addLayout(editor_layout, stretch=1)

        button_row = QHBoxLayout()
        button_row.setSpacing(8)

        self.execute_btn = QPushButton("▶ Execute")
        self.execute_btn.setObjectName("successBtn")
        self.execute_btn.clicked.connect(self._on_execute)
        button_row.addWidget(self.execute_btn)

        self.validate_btn = QPushButton("✓ Validate")
        self.validate_btn.clicked.connect(self._on_validate)
        button_row.addWidget(self.validate_btn)

        self.dry_run_btn = QPushButton("🔍 Dry Run")
        self.dry_run_btn.setObjectName("primaryBtn")
        self.dry_run_btn.clicked.connect(self._on_dry_run)
        button_row.addWidget(self.dry_run_btn)

        button_row.addStretch()

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #6c7086; font-size: 12px; padding-right: 8px;")
        button_row.addWidget(self.status_label)

        layout.addLayout(button_row)

    def _update_line_numbers(self):
        self.line_numbers.update_line_numbers()

    def get_code(self) -> str:
        return self.code_edit.toPlainText()

    def set_code(self, code: str):
        self.code_edit.setPlainText(code)

    def set_status(self, text: str, success: bool = True):
        color = "#a6e3a1" if success else "#f38ba8"
        self.status_label.setStyleSheet(f"color: {color}; font-size: 12px; padding-right: 8px;")
        self.status_label.setText(text)

    def _on_execute(self):
        if self.execute_requested:
            self.execute_requested()

    def _on_validate(self):
        if self.validate_requested:
            self.validate_requested()

    def _on_dry_run(self):
        if self.dry_run_requested:
            self.dry_run_requested()
