"""
PyQt6 UI Styling Constants
"""

MODERN_STYLESHEET = """
QMainWindow {
    background-color: #f5f5f5;
}

QGroupBox {
    font-weight: bold;
    border: 2px solid #ddd;
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 10px;
    background-color: white;
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px;
}

QLineEdit {
    padding: 8px;
    border: 1px solid #ddd;
    border-radius: 4px;
    background-color: white;
    font-size: 12px;
}

QLineEdit:focus {
    border: 2px solid #0078d4;
}

QPushButton {
    padding: 8px 16px;
    border: none;
    border-radius: 4px;
    background-color: #0078d4;
    color: white;
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

QProgressBar {
    border: 1px solid #ddd;
    border-radius: 4px;
    text-align: center;
    height: 25px;
}

QProgressBar::chunk {
    background-color: #0078d4;
    border-radius: 3px;
}

QTextEdit {
    background-color: #1e1e1e;
    color: #00ff00;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 10px;
    border: 1px solid #444;
    border-radius: 5px;
}
"""

# Color scheme
COLORS = {
    'primary': '#0078d4',
    'primary_hover': '#106ebe',
    'primary_pressed': '#005a9e',
    'success': '#00FF00',
    'fail': '#FF0000',
    'background': '#f5f5f5',
    'text_primary': '#333333',
    'text_secondary': '#666666',
    'border': '#dddddd'
}
