"""
GUI Layer - Main Application Window
File: gui/main_window.py
"""

from pathlib import Path
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog,
    QProgressBar, QTextEdit, QMessageBox, QGroupBox,
    QScrollArea
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

from application.report_service import ReportGenerationService
from visualization.chart_generator import ChartGenerator


class ReportWorker(QThread):
    """Worker thread for report generation."""
    progress = pyqtSignal(str, int)
    finished = pyqtSignal(bool, str)

    def __init__(self, fdd_files: list[Path], tdd_files: list[Path],
                 gsm_files: list[Path], tower_files: list[Path], output_path: Path):
        super().__init__()
        self.fdd_files = fdd_files
        self.tdd_files = tdd_files
        self.gsm_files = gsm_files
        self.tower_files = tower_files
        self.output_path = output_path
        self.service = ReportGenerationService()

    def run(self):
        """Execute report generation in background thread."""
        try:
            self.service.generate_report_from_files(
                self.fdd_files,
                self.tdd_files,
                self.gsm_files,
                self.tower_files,
                self.output_path,
                self.progress_callback
            )
            self.finished.emit(True, "Report generated successfully!")
        except Exception as e:
            self.finished.emit(False, f"Error: {str(e)}")

    def progress_callback(self, message: str, percentage: int):
        """Emit progress updates."""
        self.progress.emit(message, percentage)


class MainWindow(QMainWindow):
    """Main application window for FAC KPI Report Generator."""

    def __init__(self):
        super().__init__()
        self.fdd_files: list[Path] = []
        self.tdd_files: list[Path] = []
        self.gsm_files: list[Path] = []
        self.tower_files: list[Path] = []
        self.output_path: Path | None = None
        self.worker: ReportWorker | None = None

        # remember last directory for dialogs
        self.last_open_dir: Path = Path.home()

        # path edits and count labels per row
        self.fdd_path_edit = None
        self.tdd_path_edit = None
        self.gsm_path_edit = None
        self.tower_path_edit = None

        self.fdd_count_label = None
        self.tdd_count_label = None
        self.gsm_count_label = None
        self.tower_count_label = None

        self.output_path_edit = None

        self.init_ui()

    # ------------- UI INITIALIZATION -------------

    def init_ui(self):
        """Initialize the user interface (compact, 3-column rows)."""
        self.setWindowTitle("FAC KPI Report Generator")
        self.setMinimumSize(820, 680)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        central_widget = QWidget()
        scroll.setWidget(central_widget)
        self.setCentralWidget(scroll)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(8, 8, 8, 8)

        # FDD row
        fdd_group = self._create_file_section(
            "LTE FDD",
            "FDD CSV files",
            "fdd"
        )
        main_layout.addWidget(fdd_group)

        # TDD row
        tdd_group = self._create_file_section(
            "LTE TDD",
            "TDD CSV files",
            "tdd"
        )
        main_layout.addWidget(tdd_group)

        # GSM row
        gsm_group = self._create_file_section(
            "GSM (Optional)",
            "GSM CSV files",
            "gsm"
        )
        main_layout.addWidget(gsm_group)

        # TOWER row
        tower_group = self._create_file_section(
            "TOWERID (Required)",
            "TOWERID CSV file for cluster mapping",
            "tower"
        )
        main_layout.addWidget(tower_group)

        # Output row
        output_group = self._create_output_section()
        main_layout.addWidget(output_group)

        # Buttons row
        button_layout = self._create_action_buttons()
        main_layout.addLayout(button_layout)

        # Progress
        progress_group = self._create_progress_section()
        main_layout.addWidget(progress_group)

        # Log
        log_group = self._create_log_section()
        main_layout.addWidget(log_group)

        self.apply_styles()

    def _create_file_section(self, title: str, description: str, file_type: str) -> QGroupBox:
        """Create a 3-column row: name, path, browse, plus count label."""
        group = QGroupBox(title)
        layout = QVBoxLayout()
        layout.setSpacing(4)
        layout.setContentsMargins(6, 6, 6, 6)

        desc_label = QLabel(description)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #555; font-size: 8.5pt;")
        layout.addWidget(desc_label)

        row = QHBoxLayout()
        row.setSpacing(6)

        # 1) Name column
        name_label = QLabel(title)
        name_label.setStyleSheet("font-size: 8.5pt;")
        name_label.setFixedWidth(110)
        row.addWidget(name_label)

        # 2) Path column
        path_edit = QLineEdit()
        path_edit.setReadOnly(True)
        path_edit.setPlaceholderText("No file selected")
        path_edit.setStyleSheet("font-size: 8.5pt;")
        row.addWidget(path_edit, stretch=1)

        if file_type == "fdd":
            self.fdd_path_edit = path_edit
        elif file_type == "tdd":
            self.tdd_path_edit = path_edit
        elif file_type == "gsm":
            self.gsm_path_edit = path_edit
        elif file_type == "tower":
            self.tower_path_edit = path_edit

        # 3) Browse column
        browse_btn = QPushButton("Browse")
        browse_btn.setFixedSize(80, 24)
        browse_btn.clicked.connect(lambda: self.select_files(file_type))
        row.addWidget(browse_btn)

        layout.addLayout(row)

        # count label under row
        count_label = QLabel("0 file(s)")
        count_label.setStyleSheet("color: #666; font-size: 8pt;")
        layout.addWidget(count_label)

        if file_type == "fdd":
            self.fdd_count_label = count_label
        elif file_type == "tdd":
            self.tdd_count_label = count_label
        elif file_type == "gsm":
            self.gsm_count_label = count_label
        elif file_type == "tower":
            self.tower_count_label = count_label

        group.setLayout(layout)
        return group

    def _create_output_section(self) -> QGroupBox:
        """Create output selection row: name, path, browse."""
        group = QGroupBox("Output")
        layout = QVBoxLayout()
        layout.setSpacing(4)
        layout.setContentsMargins(6, 6, 6, 6)

        row = QHBoxLayout()
        row.setSpacing(6)

        name_label = QLabel("Excel file")
        name_label.setFixedWidth(110)
        name_label.setStyleSheet("font-size: 8.5pt;")
        row.addWidget(name_label)

        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("No file selected")
        self.output_path_edit.setReadOnly(True)
        self.output_path_edit.setStyleSheet("font-size: 8.5pt;")
        row.addWidget(self.output_path_edit, stretch=1)

        browse_btn = QPushButton("Browse")
        browse_btn.setFixedSize(80, 24)
        browse_btn.clicked.connect(self.select_output_file)
        row.addWidget(browse_btn)

        layout.addLayout(row)
        group.setLayout(layout)
        return group

    def _create_action_buttons(self) -> QHBoxLayout:
        """Create Clear, Generate, Charts buttons."""
        layout = QHBoxLayout()
        layout.setSpacing(8)
        layout.addStretch()

        self.clear_all_btn = QPushButton("Clear")
        self.clear_all_btn.setFixedSize(70, 26)
        self.clear_all_btn.clicked.connect(self.clear_all_files)
        layout.addWidget(self.clear_all_btn)

        self.generate_btn = QPushButton("Generate")
        self.generate_btn.setFixedSize(110, 28)
        self.generate_btn.clicked.connect(self.generate_report)
        self.generate_btn.setEnabled(False)
        layout.addWidget(self.generate_btn)

        self.view_charts_btn = QPushButton("Charts")
        self.view_charts_btn.setFixedSize(90, 28)
        self.view_charts_btn.clicked.connect(self.view_charts)
        self.view_charts_btn.setEnabled(False)
        layout.addWidget(self.view_charts_btn)

        layout.addStretch()
        return layout

    def _create_progress_section(self) -> QGroupBox:
        """Create progress bar section."""
        group = QGroupBox("Progress")
        layout = QVBoxLayout()
        layout.setSpacing(4)
        layout.setContentsMargins(6, 6, 6, 6)

        self.progress_label = QLabel("Ready")
        self.progress_label.setStyleSheet("font-size: 8.5pt;")
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(16)
        layout.addWidget(self.progress_bar)

        group.setLayout(layout)
        return group

    def _create_log_section(self) -> QGroupBox:
        """Create log output section."""
        group = QGroupBox("Log")
        layout = QVBoxLayout()
        layout.setSpacing(2)
        layout.setContentsMargins(6, 6, 6, 6)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet("font-size: 8pt;")
        layout.addWidget(self.log_text)

        group.setLayout(layout)
        return group

    # ------------- STYLING -------------

    def apply_styles(self):
        """Apply compact styling."""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 9pt;
                border: 1px solid #cccccc;
                border-radius: 5px;
                margin-top: 6px;
                padding-top: 10px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 6px;
                padding: 0 3px 0 3px;
            }
            QPushButton {
                background-color: #1976D2;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 2px 6px;
                font-size: 8.5pt;
            }
            QPushButton:hover {
                background-color: #1565C0;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
            QLineEdit {
                border: 1px solid #ddd;
                border-radius: 3px;
                padding: 2px 4px;
                background-color: white;
            }
            QProgressBar {
                border: 1px solid #ddd;
                border-radius: 3px;
                text-align: center;
                background-color: white;
                font-size: 8pt;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 2px;
            }
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 3px;
                background-color: #fafafa;
                font-family: 'Consolas', 'Courier New', monospace;
            }
        """)

    # ------------- HELPERS -------------

    def _short_path_text(self, path: Path, max_len: int = 45) -> str:
        """Return shortened string for a path."""
        s = str(path)
        if len(s) <= max_len:
            return s
        return "..." + s[-max_len:]

    # ------------- FILE SELECTION -------------

    def select_files(self, file_type: str):
        """Open dialog to select CSV files, remembering last folder."""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            f"Select {file_type.upper()} CSV File(s)",
            str(self.last_open_dir),
            "CSV Files (*.csv);;All Files (*.*)"
        )
        if not files:
            return

        file_paths = [Path(f) for f in files]
        self.last_open_dir = file_paths[0].parent

        if file_type == "fdd":
            self.fdd_files = file_paths
            if self.fdd_path_edit:
                self.fdd_path_edit.setText(self._short_path_text(self.fdd_files[0].parent))
            if self.fdd_count_label:
                self.fdd_count_label.setText(f"{len(self.fdd_files)} file(s)")
            self.log(f"Selected {len(self.fdd_files)} FDD file(s)")
        elif file_type == "tdd":
            self.tdd_files = file_paths
            if self.tdd_path_edit:
                self.tdd_path_edit.setText(self._short_path_text(self.tdd_files[0].parent))
            if self.tdd_count_label:
                self.tdd_count_label.setText(f"{len(self.tdd_files)} file(s)")
            self.log(f"Selected {len(self.tdd_files)} TDD file(s)")
        elif file_type == "gsm":
            self.gsm_files = file_paths
            if self.gsm_path_edit:
                self.gsm_path_edit.setText(self._short_path_text(self.gsm_files[0].parent))
            if self.gsm_count_label:
                self.gsm_count_label.setText(f"{len(self.gsm_files)} file(s)")
            self.log(f"Selected {len(self.gsm_files)} GSM file(s)")
        elif file_type == "tower":
            self.tower_files = file_paths
            if self.tower_path_edit:
                self.tower_path_edit.setText(self._short_path_text(self.tower_files[0].parent))
            if self.tower_count_label:
                self.tower_count_label.setText(f"{len(self.tower_files)} file(s)")
            self.log(f"Selected {len(self.tower_files)} TOWERID file(s)")

        self.check_ready_to_generate()

    def select_output_file(self):
        """Open dialog to select output file, remembering last folder."""
        default_name = self.last_open_dir / "FAC_KPI_Report.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Report As",
            str(default_name),
            "Excel Files (*.xlsx)"
        )
        if not file_path:
            return

        self.output_path = Path(file_path)
        self.last_open_dir = self.output_path.parent
        if self.output_path_edit:
            self.output_path_edit.setText(self._short_path_text(self.output_path))
        self.log(f"Output file: {self.output_path}")
        self.check_ready_to_generate()

    def clear_all_files(self):
        """Clear all selections."""
        self.fdd_files.clear()
        self.tdd_files.clear()
        self.gsm_files.clear()
        self.tower_files.clear()
        self.output_path = None

        if self.fdd_path_edit:
            self.fdd_path_edit.clear()
        if self.tdd_path_edit:
            self.tdd_path_edit.clear()
        if self.gsm_path_edit:
            self.gsm_path_edit.clear()
        if self.tower_path_edit:
            self.tower_path_edit.clear()
        if self.output_path_edit:
            self.output_path_edit.clear()

        if self.fdd_count_label:
            self.fdd_count_label.setText("0 file(s)")
        if self.tdd_count_label:
            self.tdd_count_label.setText("0 file(s)")
        if self.gsm_count_label:
            self.gsm_count_label.setText("0 file(s)")
        if self.tower_count_label:
            self.tower_count_label.setText("0 file(s)")

        self.progress_bar.setValue(0)
        self.progress_label.setText("Ready")
        self.log("Cleared all selections")
        self.check_ready_to_generate()

    # ------------- GENERATION LOGIC -------------

    def check_ready_to_generate(self):
        """Enable Generate only when required inputs are present."""
        has_data = bool(self.fdd_files or self.tdd_files or self.gsm_files)
        has_tower = bool(self.tower_files)
        has_output = self.output_path is not None

        self.generate_btn.setEnabled(has_data and has_tower and has_output)

    def generate_report(self):
        """Start report generation in background thread."""
        if not self.tower_files:
            QMessageBox.warning(self, "Missing TOWERID", "TOWERID file is required.")
            return
        if not (self.fdd_files or self.tdd_files or self.gsm_files):
            QMessageBox.warning(self, "Missing Data", "Select at least one data file (FDD, TDD, or GSM).")
            return
        if not self.output_path:
            QMessageBox.warning(self, "Missing Output", "Select an output Excel file.")
            return

        self.generate_btn.setEnabled(False)
        self.view_charts_btn.setEnabled(False)
        self.progress_bar.setValue(0)

        self.log("=" * 40)
        self.log("Start report generation")
        self.log(f"FDD: {len(self.fdd_files)}, TDD: {len(self.tdd_files)}, GSM: {len(self.gsm_files)}")
        self.log(f"TOWERID: {len(self.tower_files)}")
        self.log(f"Output: {self.output_path}")
        self.log("=" * 40)

        self.worker = ReportWorker(
            self.fdd_files,
            self.tdd_files,
            self.gsm_files,
            self.tower_files,
            self.output_path
        )
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.generation_finished)
        self.worker.start()

    def update_progress(self, message: str, percentage: int):
        """Update progress."""
        self.progress_label.setText(message)
        self.progress_bar.setValue(percentage)
        self.log(message)

    def generation_finished(self, success: bool, message: str):
        """Handle completion of report generation."""
        self.log(message)

        if success:
            QMessageBox.information(
                self,
                "Success",
                f"Report generated.\n\nSaved to:\n{self.output_path}"
            )
            self.view_charts_btn.setEnabled(True)
        else:
            QMessageBox.critical(self, "Error", f"Failed to generate report:\n\n{message}")

        self.generate_btn.setEnabled(True)
        self.worker = None

    def view_charts(self):
        """Placeholder for chart visualization."""
        try:
            self.log("Charts requested")
            QMessageBox.information(
                self,
                "Charts",
                "Chart visualization can be implemented using ChartGenerator."
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to display charts:\n\n{str(e)}")

    # ------------- LOGGING -------------

    def log(self, message: str):
        """Append message to log."""
        self.log_text.append(f"[{self._get_timestamp()}] {message}")
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    @staticmethod
    def _get_timestamp() -> str:
        """Current time string."""
        from datetime import datetime
        return datetime.now().strftime("%H:%M:%S")
