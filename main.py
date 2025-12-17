"""
FAC Report Generator - Main Entry Point
Author: Network Optimization Team
Description: Generate FAC KPI Achievement reports from LTE, GSM, and Cluster data
"""

import sys
from PyQt6.QtWidgets import QApplication
from ui.main_window import MainWindow
from PyQt6.QtGui import QFont

def main():
    """Main application entry point"""
    app = QApplication(sys.argv)

    # Set application style
    app.setFont(QFont("Segoe UI", 9))
    app.setStyle("Fusion")

    # Create and show main window
    window = MainWindow()
    window.show()

    # Run application
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
