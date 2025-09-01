# UI 스타일 정의

MAIN_STYLE = """
QMainWindow {
    background-color: #f5f5f5;
}

QWidget {
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 10pt;
}

QPushButton {
    background-color: #0078d4;
    color: white;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #106ebe;
}

QPushButton:pressed {
    background-color: #005a9e;
}

QPushButton:disabled {
    background-color: #cccccc;
    color: #666666;
}

QLineEdit {
    padding: 6px;
    border: 2px solid #e1e1e1;
    border-radius: 4px;
    background-color: white;
}

QLineEdit:focus {
    border-color: #0078d4;
}

QLineEdit:disabled {
    background-color: #f5f5f5;
    color: #666666;
}

QCheckBox {
    spacing: 8px;
}

QCheckBox::indicator {
    width: 18px;
    height: 18px;
    border: 2px solid #e1e1e1;
    border-radius: 3px;
    background-color: white;
}

QCheckBox::indicator:checked {
    background-color: #0078d4;
    border-color: #0078d4;
}

QTableWidget {
    gridline-color: #e1e1e1;
    background-color: white;
    alternate-background-color: #f9f9f9;
    selection-background-color: #e3f2fd;
}

QTableWidget::item {
    padding: 8px;
    border-bottom: 1px solid #f0f0f0;
}

QTableWidget::item:selected {
    background-color: #e3f2fd;
    color: black;
}

QHeaderView::section {
    background-color: #f0f0f0;
    padding: 8px;
    border: none;
    border-bottom: 2px solid #e1e1e1;
    font-weight: bold;
}

QGroupBox {
    font-weight: bold;
    border: 2px solid #e1e1e1;
    border-radius: 6px;
    margin-top: 12px;
    padding-top: 8px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 8px 0 8px;
    background-color: #f5f5f5;
}

QProgressBar {
    border: 2px solid #e1e1e1;
    border-radius: 4px;
    text-align: center;
    background-color: #f0f0f0;
}

QProgressBar::chunk {
    background-color: #0078d4;
    border-radius: 2px;
}

QLabel {
    color: #333333;
}

QLabel[class="title"] {
    font-size: 16pt;
    font-weight: bold;
    color: #0078d4;
}

QLabel[class="subtitle"] {
    font-size: 12pt;
    font-weight: bold;
    color: #666666;
}

QTextEdit {
    border: 2px solid #e1e1e1;
    border-radius: 4px;
    background-color: white;
    padding: 8px;
}

QScrollBar:vertical {
    background-color: #f0f0f0;
    width: 12px;
    border-radius: 6px;
}

QScrollBar::handle:vertical {
    background-color: #c1c1c1;
    border-radius: 6px;
    min-height: 20px;
}

QScrollBar::handle:vertical:hover {
    background-color: #a8a8a8;
}
"""

SUCCESS_STYLE = """
QPushButton {
    background-color: #107c10;
    color: white;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #0e6b0e;
}
"""

WARNING_STYLE = """
QPushButton {
    background-color: #d83b01;
    color: white;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #b32d01;
}
"""

EDIT_STYLE = """
QPushButton {
    background-color: #0078d4;
    color: white;
    border: none;
    padding: 6px 12px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 9pt;
}

QPushButton:hover {
    background-color: #106ebe;
}
"""

DELETE_STYLE = """
QPushButton {
    background-color: #d83b01;
    color: white;
    border: none;
    padding: 6px 12px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 9pt;
}

QPushButton:hover {
    background-color: #b32d01;
}
"""
