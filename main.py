"""
FAC Report Generator - Main Entry Point
Fixed template path: ./datatemplate.xlsx
"""
import sys
from PyQt6.QtWidgets import QApplication
from ui.main_window import MainWindow


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern style

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
