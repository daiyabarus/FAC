"""
FAC KPI Report Generator - Main Entry Point
Clean Architecture Implementation
File: main.py
"""

import sys
from pathlib import Path
from PyQt6.QtWidgets import QApplication
from gui.main_window import MainWindow


def main() -> None:
    """Main entry point for the application."""
    app = QApplication(sys.argv)
    app.setApplicationName("FAC KPI Report Generator")
    app.setOrganizationName("FAC Analytics")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
