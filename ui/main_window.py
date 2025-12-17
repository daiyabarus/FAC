"""Main PyQt6 GUI window"""

from PyQt6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QFileDialog,
    QTextEdit,
    QProgressBar,
    QMessageBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
from pathlib import Path


class ProcessThread(QThread):
    """Background thread for processing"""

    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, lte_file, gsm_file, cluster_file, output_dir, template_file):
        super().__init__()
        self.lte_file = lte_file
        self.gsm_file = gsm_file
        self.cluster_file = cluster_file
        self.output_dir = output_dir
        self.template_file = template_file

    def run(self):
        """Run the processing"""
        try:
            from data.loader import DataLoader
            from data.transformer import DataTransformer
            from kpi.calculator import KPICalculator
            from kpi.validator import KPIValidator
            from report.excel_writer import ExcelReportWriter
            from assets.logos import get_xlsmart_logo, get_zte_logo

            # Load data
            self.progress.emit("Loading data files...")
            loader = DataLoader()
            loader.load_lte_file(self.lte_file)
            loader.load_gsm_file(self.gsm_file)
            loader.load_cluster_file(self.cluster_file)

            # Transform data
            self.progress.emit("Transforming and enriching data...")
            transformer = DataTransformer(loader)
            transformed_data = transformer.transform_all()

            # Calculate KPIs
            self.progress.emit("Calculating KPIs...")
            calculator = KPICalculator(transformed_data)
            kpi_data = calculator.calculate_all()

            # Validate KPIs
            self.progress.emit("Validating KPIs against baselines...")
            validator = KPIValidator(kpi_data)
            validation_results = validator.validate_all()

            # Generate reports
            self.progress.emit("Generating Excel reports...")

            clusters = transformed_data["lte"]["CLUSTER"].dropna().unique()

            for cluster in clusters:
                self.progress.emit(f"Writing report for {cluster}...")

                writer = ExcelReportWriter(
                    self.template_file,
                    self.output_dir,
                    validation_results,
                    kpi_data,
                    transformed_data,
                )

                # Set logos
                writer.set_logos(get_xlsmart_logo(), get_zte_logo())

                # Write report
                writer.write_report(cluster)

            self.progress.emit("✓ All reports generated successfully!")
            self.finished.emit(True, f"Generated {len(clusters)} reports")

        except Exception as e:
            import traceback

            error_detail = traceback.format_exc()
            self.progress.emit(f"✗ Error: {str(e)}")
            self.progress.emit(error_detail)
            self.finished.emit(False, str(e))


class MainWindow(QMainWindow):
    """Main application window"""

    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("FAC Report Generator")
        self.setGeometry(100, 100, 900, 700)

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main layout
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # Title
        title = QLabel("FAC KPI Report Generator")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Subtitle
        subtitle = QLabel("Generate FAC reports from LTE, GSM, and Cluster data")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(subtitle)

        layout.addSpacing(20)

        # File inputs
        self.lte_input = self._create_file_input(
            "LTE Data File (Excel):", "Browse LTE File"
        )
        self.gsm_input = self._create_file_input(
            "GSM Data File (Excel):", "Browse GSM File"
        )
        self.cluster_input = self._create_file_input(
            "Cluster Data File (Excel):", "Browse Cluster File"
        )

        layout.addLayout(self.lte_input["layout"])
        layout.addLayout(self.gsm_input["layout"])
        layout.addLayout(self.cluster_input["layout"])

        layout.addSpacing(10)

        # PERBAIKAN 1: Hardcode template - no browse button
        # template_info = QLabel("📄 Template: datatemplate.xlsx (hardcoded)")
        # template_info.setStyleSheet("color: #666; font-style: italic; padding: 5px;")
        # layout.addWidget(template_info)

        # Output directory
        self.output_input = self._create_folder_input(
            "Output Directory:", "Browse Output Folder"
        )
        layout.addLayout(self.output_input["layout"])

        layout.addSpacing(20)

        # Generate button
        self.generate_btn = QPushButton("Generate Reports")
        self.generate_btn.setMinimumHeight(45)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.generate_btn.clicked.connect(self.generate_reports)
        layout.addWidget(self.generate_btn)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)

        # Log output
        log_label = QLabel("Processing Log:")
        log_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(log_label)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMinimumHeight(200)
        self.log_output.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 10px;
                font-family: 'Courier New', monospace;
                font-size: 11px;
            }
        """)
        layout.addWidget(self.log_output)

        central_widget.setLayout(layout)

        self.log("FAC Report Generator initialized. Please select input files.")

    def _create_file_input(self, label_text, button_text):
        """Create a file input row"""
        layout = QHBoxLayout()

        label = QLabel(label_text)
        label.setMinimumWidth(200)

        line_edit = QLineEdit()
        line_edit.setPlaceholderText("No file selected")
        line_edit.setReadOnly(True)

        button = QPushButton(button_text)
        button.setMinimumWidth(150)
        button.clicked.connect(lambda: self._browse_file(line_edit))

        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button)

        return {"layout": layout, "input": line_edit, "button": button}

    def _create_folder_input(self, label_text, button_text):
        """Create a folder input row"""
        layout = QHBoxLayout()

        label = QLabel(label_text)
        label.setMinimumWidth(200)

        line_edit = QLineEdit()
        line_edit.setPlaceholderText("No folder selected")
        line_edit.setReadOnly(True)

        button = QPushButton(button_text)
        button.setMinimumWidth(150)
        button.clicked.connect(lambda: self._browse_folder(line_edit))

        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button)

        return {"layout": layout, "input": line_edit, "button": button}

    def _browse_file(self, line_edit):
        """Browse for a file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            line_edit.setText(file_path)
            self.log(f"Selected file: {Path(file_path).name}")

    def _browse_folder(self, line_edit):
        """Browse for a folder"""
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder")

        if folder_path:
            line_edit.setText(folder_path)
            self.log(f"Selected output folder: {folder_path}")

    def log(self, message):
        """Add message to log output"""
        self.log_output.append(message)
        self.log_output.verticalScrollBar().setValue(
            self.log_output.verticalScrollBar().maximum()
        )

    def generate_reports(self):
        """Start report generation"""
        # Validate inputs
        lte_file = self.lte_input["input"].text()
        gsm_file = self.gsm_input["input"].text()
        cluster_file = self.cluster_input["input"].text()
        output_dir = self.output_input["input"].text()

        if not all([lte_file, gsm_file, cluster_file, output_dir]):
            QMessageBox.warning(
                self,
                "Missing Input",
                "Please select all required input files and output directory.",
            )
            return

        # PERBAIKAN 1: Hardcode template file
        template_file = "./datatemplate.xlsx"

        # Check if template exists
        if not Path(template_file).exists():
            QMessageBox.warning(
                self,
                "Template Not Found",
                f"Template file '{template_file}' not found in current directory.\n"
                f"Please ensure datatemplate.xlsx exists in the same folder as main.py",
            )
            return

        self.log(f"Using template: {template_file}")

        # Disable button and show progress
        self.generate_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate

        self.log("\n" + "=" * 60)
        self.log("Starting report generation...")
        self.log("=" * 60)

        # Start processing thread
        self.thread = ProcessThread(
            lte_file, gsm_file, cluster_file, output_dir, template_file
        )
        self.thread.progress.connect(self.log)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def on_finished(self, success, message):
        """Handle processing completion"""
        self.generate_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        if success:
            self.log("\n" + "=" * 60)
            self.log("✓ REPORT GENERATION COMPLETE!")
            self.log(message)
            self.log("=" * 60)

            QMessageBox.information(
                self, "Success", f"Reports generated successfully!\n\n{message}"
            )
        else:
            self.log("\n" + "=" * 60)
            self.log("✗ REPORT GENERATION FAILED")
            self.log(f"Error: {message}")
            self.log("=" * 60)

            QMessageBox.critical(
                self, "Error", f"Report generation failed:\n\n{message}"
            )
