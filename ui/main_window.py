"""
Main Window for FAC Report Generator
Modern PyQt6 UI with file browser and output selection
"""
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog,
    QProgressBar, QMessageBox, QGroupBox, QTextEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
from pathlib import Path
from services.report_generator import ReportGenerator


class ProcessThread(QThread):
    """Background thread for processing reports"""
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, lte_file, gsm_file, cluster_file, output_dir):
        super().__init__()
        self.lte_file = lte_file
        self.gsm_file = gsm_file
        self.cluster_file = cluster_file
        self.output_dir = output_dir

    def run(self):
        """Run report generation in background"""
        try:
            self.status.emit("Loading template...")
            self.progress.emit(10)

            generator = ReportGenerator()

            self.status.emit("Loading data files...")
            self.progress.emit(20)

            generator.load_data(
                self.lte_file,
                self.gsm_file,
                self.cluster_file
            )

            self.status.emit("Processing data...")
            self.progress.emit(40)

            generator.process_data()

            self.status.emit("Calculating KPIs...")
            self.progress.emit(60)

            generator.calculate_kpis()

            self.status.emit("Generating charts...")
            self.progress.emit(80)

            generator.generate_charts()

            self.status.emit("Writing Excel reports...")
            self.progress.emit(90)

            output_files = generator.generate_reports(self.output_dir)

            self.progress.emit(100)
            self.finished.emit(
                True, f"Success! Generated {len(output_files)} reports")

        except Exception as e:
            self.finished.emit(False, f"Error: {str(e)}")


class MainWindow(QMainWindow):
    """Main application window"""

    TEMPLATE_PATH = "./datatemplate.xlsx"

    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        """Initialize user interface"""
        self.setWindowTitle("FAC Report Generator")
        self.setMinimumSize(800, 600)

        # Apply modern styling
        self.apply_modern_style()

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main layout
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        # Title
        title = QLabel("FAC Report Generator")
        title_font = QFont("Segoe UI", 24, QFont.Weight.Bold)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Template info
        template_label = QLabel(f"📋 Template: {self.TEMPLATE_PATH}")
        template_label.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(template_label)

        # Input files group
        input_group = self.create_input_group()
        layout.addWidget(input_group)

        # Output directory group
        output_group = self.create_output_group()
        layout.addWidget(output_group)

        # Progress section
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(self.status_label)

        # Log area
        log_group = QGroupBox("Processing Log")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #00ff00;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 10px;
                border: 1px solid #444;
                border-radius: 5px;
            }
        """)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # Generate button
        self.generate_btn = QPushButton("🚀 Generate Reports")
        self.generate_btn.setMinimumHeight(50)
        self.generate_btn.clicked.connect(self.generate_reports)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                border-radius: 5px;
                font-size: 16px;
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
        """)
        layout.addWidget(self.generate_btn)

        # Add stretch
        layout.addStretch()

    def create_input_group(self):
        """Create input files group"""
        group = QGroupBox("Input Files")
        layout = QVBoxLayout()

        # LTE File
        lte_layout = QHBoxLayout()
        lte_label = QLabel("LTE Data:")
        lte_label.setMinimumWidth(100)
        self.lte_input = QLineEdit()
        self.lte_input.setPlaceholderText("Select LTE Excel file (Sheet0)...")
        lte_btn = QPushButton("📁 Browse")
        lte_btn.clicked.connect(
            lambda: self.browse_file(self.lte_input, "LTE"))
        lte_layout.addWidget(lte_label)
        lte_layout.addWidget(self.lte_input)
        lte_layout.addWidget(lte_btn)
        layout.addLayout(lte_layout)

        # GSM File
        gsm_layout = QHBoxLayout()
        gsm_label = QLabel("GSM Data:")
        gsm_label.setMinimumWidth(100)
        self.gsm_input = QLineEdit()
        self.gsm_input.setPlaceholderText("Select GSM Excel file (Sheet0)...")
        gsm_btn = QPushButton("📁 Browse")
        gsm_btn.clicked.connect(
            lambda: self.browse_file(self.gsm_input, "GSM"))
        gsm_layout.addWidget(gsm_label)
        gsm_layout.addWidget(self.gsm_input)
        gsm_layout.addWidget(gsm_btn)
        layout.addLayout(gsm_layout)

        # Cluster File
        cluster_layout = QHBoxLayout()
        cluster_label = QLabel("Cluster Data:")
        cluster_label.setMinimumWidth(100)
        self.cluster_input = QLineEdit()
        self.cluster_input.setPlaceholderText(
            "Select Cluster Excel file (CLUSTER sheet)...")
        cluster_btn = QPushButton("📁 Browse")
        cluster_btn.clicked.connect(
            lambda: self.browse_file(self.cluster_input, "Cluster"))
        cluster_layout.addWidget(cluster_label)
        cluster_layout.addWidget(self.cluster_input)
        cluster_layout.addWidget(cluster_btn)
        layout.addLayout(cluster_layout)

        group.setLayout(layout)
        return group

    def create_output_group(self):
        """Create output directory group"""
        group = QGroupBox("Output Directory")
        layout = QHBoxLayout()

        self.output_input = QLineEdit()
        self.output_input.setPlaceholderText("Select output directory...")
        output_btn = QPushButton("📂 Browse")
        output_btn.clicked.connect(self.browse_output)

        layout.addWidget(self.output_input)
        layout.addWidget(output_btn)

        group.setLayout(layout)
        return group

    def browse_file(self, line_edit, file_type):
        """Browse for input file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"Select {file_type} File",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            line_edit.setText(file_path)
            self.log(f"✓ Selected {file_type} file: {Path(file_path).name}")

    def browse_output(self):
        """Browse for output directory"""
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory"
        )
        if dir_path:
            self.output_input.setText(dir_path)
            self.log(f"✓ Selected output directory: {dir_path}")

    def log(self, message):
        """Add message to log"""
        self.log_text.append(message)

    def generate_reports(self):
        """Start report generation"""
        # Validate inputs
        if not self.lte_input.text():
            QMessageBox.warning(self, "Missing Input",
                                "Please select LTE data file")
            return
        if not self.gsm_input.text():
            QMessageBox.warning(self, "Missing Input",
                                "Please select GSM data file")
            return
        if not self.cluster_input.text():
            QMessageBox.warning(self, "Missing Input",
                                "Please select Cluster data file")
            return
        if not self.output_input.text():
            QMessageBox.warning(self, "Missing Output",
                                "Please select output directory")
            return

        # Check template exists
        if not Path(self.TEMPLATE_PATH).exists():
            QMessageBox.critical(
                self,
                "Template Not Found",
                f"Template file not found: {self.TEMPLATE_PATH}\n\nPlease ensure datatemplate.xlsx is in the same directory as the application."
            )
            return

        # Disable button and show progress
        self.generate_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log("\n" + "="*50)
        self.log("Starting report generation...")

        # Start background processing
        self.process_thread = ProcessThread(
            self.lte_input.text(),
            self.gsm_input.text(),
            self.cluster_input.text(),
            self.output_input.text()
        )
        self.process_thread.progress.connect(self.update_progress)
        self.process_thread.status.connect(self.update_status)
        self.process_thread.finished.connect(self.process_finished)
        self.process_thread.start()

    def update_progress(self, value):
        """Update progress bar"""
        self.progress_bar.setValue(value)

    def update_status(self, message):
        """Update status label"""
        self.status_label.setText(message)
        self.log(f"⏳ {message}")

    def process_finished(self, success, message):
        """Handle process completion"""
        self.generate_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        if success:
            self.log(f"✅ {message}")
            QMessageBox.information(self, "Success", message)
        else:
            self.log(f"❌ {message}")
            QMessageBox.critical(self, "Error", message)

    def apply_modern_style(self):
        """Apply modern styling to the application"""
        self.setStyleSheet("""
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
        """)
