CATPPUCCIN_MOCHA = {
    "base": "#1e1e2e",
    "mantle": "#181825",
    "crust": "#11111b",
    "surface0": "#313244",
    "surface1": "#45475a",
    "surface2": "#585b70",
    "overlay0": "#6c7086",
    "overlay1": "#7f849c",
    "overlay2": "#9399b2",
    "text": "#cdd6f4",
    "subtext0": "#a6adc8",
    "subtext1": "#bac2de",
    "blue": "#89b4fa",
    "green": "#a6e3a1",
    "red": "#f38ba8",
    "yellow": "#f9e2af",
    "purple": "#cba6f7",
    "peach": "#fab387",
    "teal": "#94e2d5",
    "pink": "#f5c2e7",
    "lavender": "#b4befe",
}

STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {CATPPUCCIN_MOCHA["base"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
    font-size: 13px;
}}

QMainWindow {{
    background-color: {CATPPUCCIN_MOCHA["base"]};
}}

QWidget#sidebar {{
    background-color: {CATPPUCCIN_MOCHA["crust"]};
    border-right: 1px solid {CATPPUCCIN_MOCHA["surface0"]};
}}

QPushButton {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    padding: 7px 16px;
    min-height: 20px;
}}

QPushButton:hover {{
    background-color: {CATPPUCCIN_MOCHA["surface1"]};
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QPushButton:pressed {{
    background-color: {CATPPUCCIN_MOCHA["surface2"]};
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QPushButton:disabled {{
    background-color: {CATPPUCCIN_MOCHA["mantle"]};
    color: {CATPPUCCIN_MOCHA["overlay0"]};
    border-color: {CATPPUCCIN_MOCHA["surface0"]};
}}

QPushButton#sidebarBtn {{
    background-color: transparent;
    border: none;
    border-radius: 8px;
    padding: 10px 12px;
    text-align: left;
    font-size: 13px;
    min-height: 18px;
    color: {CATPPUCCIN_MOCHA["subtext0"]};
}}

QPushButton#sidebarBtn:hover {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
}}

QPushButton#sidebarBtn[active="true"] {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["blue"]};
    border-left: 3px solid {CATPPUCCIN_MOCHA["blue"]};
    font-weight: bold;
}}

QPushButton#primaryBtn {{
    background-color: {CATPPUCCIN_MOCHA["blue"]};
    color: {CATPPUCCIN_MOCHA["crust"]};
    border: none;
    font-weight: bold;
    border-radius: 6px;
    padding: 8px 20px;
}}

QPushButton#primaryBtn:hover {{
    background-color: {CATPPUCCIN_MOCHA["lavender"]};
}}

QPushButton#primaryBtn:pressed {{
    background-color: {CATPPUCCIN_MOCHA["surface2"]};
}}

QPushButton#dangerBtn {{
    background-color: {CATPPUCCIN_MOCHA["red"]};
    color: {CATPPUCCIN_MOCHA["crust"]};
    border: none;
    font-weight: bold;
}}

QPushButton#dangerBtn:hover {{
    background-color: #e06c8c;
}}

QPushButton#successBtn {{
    background-color: {CATPPUCCIN_MOCHA["green"]};
    color: {CATPPUCCIN_MOCHA["crust"]};
    border: none;
    font-weight: bold;
}}

QPushButton#successBtn:hover {{
    background-color: #8bd489;
}}

QLineEdit {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    padding: 7px 12px;
    selection-background-color: {CATPPUCCIN_MOCHA["blue"]};
    selection-color: {CATPPUCCIN_MOCHA["crust"]};
}}

QLineEdit:focus {{
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QTextEdit {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    padding: 6px;
    selection-background-color: {CATPPUCCIN_MOCHA["blue"]};
    selection-color: {CATPPUCCIN_MOCHA["crust"]};
}}

QTextEdit:focus {{
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QPlainTextEdit {{
    background-color: {CATPPUCCIN_MOCHA["mantle"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    padding: 6px;
    selection-background-color: {CATPPUCCIN_MOCHA["blue"]};
    selection-color: {CATPPUCCIN_MOCHA["crust"]};
    font-family: "Cascadia Code", "Fira Code", "Consolas", monospace;
    font-size: 13px;
}}

QPlainTextEdit:focus {{
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QComboBox {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    padding: 6px 12px;
    min-height: 20px;
}}

QComboBox:hover {{
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QComboBox::drop-down {{
    border: none;
    width: 24px;
}}

QComboBox::down-arrow {{
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid {CATPPUCCIN_MOCHA["subtext0"]};
    margin-right: 6px;
}}

QComboBox QAbstractItemView {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    selection-background-color: {CATPPUCCIN_MOCHA["blue"]};
    selection-color: {CATPPUCCIN_MOCHA["crust"]};
    outline: none;
}}

QTableWidget {{
    background-color: {CATPPUCCIN_MOCHA["mantle"]};
    alternate-background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    gridline-color: {CATPPUCCIN_MOCHA["surface1"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    selection-background-color: {CATPPUCCIN_MOCHA["blue"]};
    selection-color: {CATPPUCCIN_MOCHA["crust"]};
}}

QTableWidget::item {{
    padding: 4px 8px;
    border-bottom: 1px solid {CATPPUCCIN_MOCHA["surface0"]};
}}

QTableWidget::item:selected {{
    background-color: {CATPPUCCIN_MOCHA["blue"]};
    color: {CATPPUCCIN_MOCHA["crust"]};
}}

QHeaderView::section {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    padding: 6px 8px;
    border: none;
    border-right: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-bottom: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    font-weight: bold;
}}

QHeaderView::section:hover {{
    background-color: {CATPPUCCIN_MOCHA["surface1"]};
}}

QTabWidget::pane {{
    background-color: {CATPPUCCIN_MOCHA["base"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    top: -1px;
}}

QTabBar::tab {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["subtext0"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-bottom: none;
    padding: 8px 18px;
    margin-right: 2px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
}}

QTabBar::tab:selected {{
    background-color: {CATPPUCCIN_MOCHA["base"]};
    color: {CATPPUCCIN_MOCHA["blue"]};
    border-bottom: 2px solid {CATPPUCCIN_MOCHA["blue"]};
}}

QTabBar::tab:hover:!selected {{
    background-color: {CATPPUCCIN_MOCHA["surface1"]};
    color: {CATPPUCCIN_MOCHA["text"]};
}}

QScrollBar:vertical {{
    background-color: {CATPPUCCIN_MOCHA["mantle"]};
    width: 10px;
    border-radius: 5px;
    margin: 0;
}}

QScrollBar::handle:vertical {{
    background-color: {CATPPUCCIN_MOCHA["surface2"]};
    min-height: 30px;
    border-radius: 5px;
}}

QScrollBar::handle:vertical:hover {{
    background-color: {CATPPUCCIN_MOCHA["overlay0"]};
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}

QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: none;
}}

QScrollBar:horizontal {{
    background-color: {CATPPUCCIN_MOCHA["mantle"]};
    height: 10px;
    border-radius: 5px;
    margin: 0;
}}

QScrollBar::handle:horizontal {{
    background-color: {CATPPUCCIN_MOCHA["surface2"]};
    min-width: 30px;
    border-radius: 5px;
}}

QScrollBar::handle:horizontal:hover {{
    background-color: {CATPPUCCIN_MOCHA["overlay0"]};
}}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}

QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
    background: none;
}}

QCheckBox {{
    color: {CATPPUCCIN_MOCHA["text"]};
    spacing: 8px;
}}

QCheckBox::indicator {{
    width: 18px;
    height: 18px;
    border-radius: 4px;
    border: 2px solid {CATPPUCCIN_MOCHA["surface2"]};
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
}}

QCheckBox::indicator:checked {{
    background-color: {CATPPUCCIN_MOCHA["blue"]};
    border-color: {CATPPUCCIN_MOCHA["blue"]};
    image: none;
}}

QCheckBox::indicator:checked {{
    background-color: {CATPPUCCIN_MOCHA["blue"]};
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QCheckBox::indicator:hover {{
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QRadioButton {{
    color: {CATPPUCCIN_MOCHA["text"]};
    spacing: 8px;
}}

QRadioButton::indicator {{
    width: 18px;
    height: 18px;
    border-radius: 9px;
    border: 2px solid {CATPPUCCIN_MOCHA["surface2"]};
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
}}

QRadioButton::indicator:checked {{
    background-color: {CATPPUCCIN_MOCHA["blue"]};
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QRadioButton::indicator:hover {{
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QGroupBox {{
    color: {CATPPUCCIN_MOCHA["blue"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 16px;
    font-weight: bold;
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 6px;
}}

QSplitter::handle {{
    background-color: {CATPPUCCIN_MOCHA["surface1"]};
}}

QSplitter::handle:horizontal {{
    width: 2px;
}}

QSplitter::handle:vertical {{
    height: 2px;
}}

QSplitter::handle:hover {{
    background-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QStatusBar {{
    background-color: {CATPPUCCIN_MOCHA["crust"]};
    color: {CATPPUCCIN_MOCHA["subtext0"]};
    border-top: 1px solid {CATPPUCCIN_MOCHA["surface0"]};
    padding: 2px 8px;
    font-size: 12px;
}}

QStatusBar::item {{
    border: none;
}}

QToolTip {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 4px;
    padding: 6px 10px;
    font-size: 12px;
}}

QLabel {{
    color: {CATPPUCCIN_MOCHA["text"]};
    background: transparent;
}}

QLabel#headingLabel {{
    font-size: 22px;
    font-weight: bold;
    color: {CATPPUCCIN_MOCHA["blue"]};
}}

QLabel#subheadingLabel {{
    font-size: 14px;
    color: {CATPPUCCIN_MOCHA["subtext0"]};
}}

QLabel#statusLabel {{
    font-size: 12px;
    padding: 2px 6px;
    border-radius: 4px;
}}

QLabel#successStatus {{
    color: {CATPPUCCIN_MOCHA["green"]};
}}

QLabel#errorStatus {{
    color: {CATPPUCCIN_MOCHA["red"]};
}}

QListWidget {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    outline: none;
}}

QListWidget::item {{
    padding: 6px 10px;
    border-bottom: 1px solid {CATPPUCCIN_MOCHA["surface0"]};
}}

QListWidget::item:selected {{
    background-color: {CATPPUCCIN_MOCHA["blue"]};
    color: {CATPPUCCIN_MOCHA["crust"]};
}}

QListWidget::item:hover:!selected {{
    background-color: {CATPPUCCIN_MOCHA["surface1"]};
}}

QScrollArea {{
    background-color: transparent;
    border: none;
}}

QScrollBar#chatScrollBar:vertical {{
    width: 8px;
}}

QFrame#separator {{
    background-color: {CATPPUCCIN_MOCHA["surface1"]};
    max-height: 1px;
}}

QFrame#cardFrame {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 8px;
    padding: 12px;
}}

QFrame#cardFrame:hover {{
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QToolButton {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    padding: 5px 10px;
}}

QToolButton:hover {{
    background-color: {CATPPUCCIN_MOCHA["surface1"]};
    border-color: {CATPPUCCIN_MOCHA["blue"]};
}}

QToolBar {{
    background-color: {CATPPUCCIN_MOCHA["crust"]};
    border-bottom: 1px solid {CATPPUCCIN_MOCHA["surface0"]};
    spacing: 6px;
    padding: 4px;
}}

QToolBar::separator {{
    width: 1px;
    background-color: {CATPPUCCIN_MOCHA["surface1"]};
    margin: 4px 6px;
}}

QSpinBox {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    color: {CATPPUCCIN_MOCHA["text"]};
    border: 1px solid {CATPPUCCIN_MOCHA["surface1"]};
    border-radius: 6px;
    padding: 4px 8px;
}}

QSpinBox::up-button, QSpinBox::down-button {{
    background-color: {CATPPUCCIN_MOCHA["surface1"]};
    border: none;
    width: 16px;
}}

QSpinBox::up-button:hover, QSpinBox::down-button:hover {{
    background-color: {CATPPUCCIN_MOCHA["surface2"]};
}}

QProgressBar {{
    background-color: {CATPPUCCIN_MOCHA["surface0"]};
    border: none;
    border-radius: 4px;
    text-align: center;
    color: {CATPPUCCIN_MOCHA["text"]};
    min-height: 8px;
}}

QProgressBar::chunk {{
    background-color: {CATPPUCCIN_MOCHA["blue"]};
    border-radius: 4px;
}}
"""
