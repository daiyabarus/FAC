"""Ultra Minimalist Futuristic FAC Report Generator - PyQt6 (Updated with NGI Input)"""

import sys
from pathlib import Path

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
    QApplication,
    QToolTip,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QPixmap, QIcon


class ProcessThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, lte_file, gsm_file, ngi_file, cluster_file, output_dir, template_file):
        super().__init__()
        self.lte_file = lte_file
        self.gsm_file = gsm_file
        self.ngi_file = ngi_file
        self.cluster_file = cluster_file
        self.output_dir = output_dir
        self.template_file = template_file

    def run(self):
        try:
            from data.loader import DataLoader
            from data.transformer import DataTransformer
            from kpi.calculator import KPICalculator
            from kpi.validator import KPIValidator
            from report.excel_writer import ExcelReportWriter
            from assets.logos import get_xlsmart_logo, get_zte_logo

            self.progress.emit("Loading data...")
            loader = DataLoader()
            loader.load_lte_file(self.lte_file)
            loader.load_gsm_file(self.gsm_file)
            loader.load_ngi_file(self.ngi_file)
            loader.load_cluster_file(self.cluster_file)

            self.progress.emit("Transforming data...")
            transformer = DataTransformer(loader)
            transformed_data = transformer.transform_all()
            self.progress.emit(f"DEBUG: NGI data = {transformed_data.get('ngi') is not None}")
            if transformed_data.get('ngi') is not None:
                self.progress.emit(f"DEBUG: NGI rows = {len(transformed_data['ngi'])}")

            self.progress.emit("Calculating KPIs...")
            calculator = KPICalculator(transformed_data)
            kpi_data = calculator.calculate_all()

            self.progress.emit("Validating...")
            validator = KPIValidator(kpi_data)
            validation_results = validator.validate_all()

            self.progress.emit("Generating reports...")
            clusters = transformed_data["lte"]["CLUSTER"].dropna().unique()

            for cluster in clusters:
                self.progress.emit(f"{cluster} → writing...")
                writer = ExcelReportWriter(
                    self.template_file,
                    self.output_dir,
                    validation_results,
                    kpi_data,
                    transformed_data,
                )
                writer.set_logos(get_xlsmart_logo(), get_zte_logo())
                writer.write_report(cluster)

            self.progress.emit("✓ Done!")
            self.finished.emit(True, f"{len(clusters)} reports generated")

        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            self.progress.emit(f"✗ {str(e)}")
            self.progress.emit(error_detail)
            self.finished.emit(False, str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FAC-GR")
        self.setGeometry(100, 100, 800, 650)  # Sedikit lebih tinggi untuk 5 input

        QToolTip.setFont(QFont("Segoe UI", 10))
        self.setStyleSheet("""
            QMainWindow { background-color: #0a0e1a; }
            QToolTip {
                background-color: #1e293b;
                color: #06b6d4;
                border: 1px solid #06b6d4;
                padding: 6px;
                border-radius: 6px;
            }
            QLineEdit {
                background-color: #111827;
                border: 1px solid #1e293b;
                border-radius: 6px;
                padding: 8px 10px;
                color: #e2e8f0;
                font-size: 12px;
            }
            QLineEdit:focus { border: 1px solid #06b6d4; }
            QPushButton {
                background-color: #06b6d4;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover { background-color: #0891b2; }
            QPushButton:pressed { background-color: #0e7490; }
            QPushButton:disabled { background-color: #1e293b; }
            /* Transparent Browse Button */
            QPushButton#BrowseButton {
                background: transparent;
                border: none;
                padding: 8px;
            }
            QPushButton#BrowseButton:hover {
                background: rgba(6, 182, 212, 0.15);
                border-radius: 8px;
            }
            QPushButton#BrowseButton:pressed {
                background: rgba(6, 182, 212, 0.3);
            }
            QTextEdit {
                background-color: #111827;
                color: #94a3b8;
                border: 1px solid #1e293b;
                border-radius: 8px;
                padding: 10px;
                font-family: 'Consolas', monospace;
                font-size: 11px;
            }
            QProgressBar {
                border: none;
                border-radius: 4px;
                background-color: #1e293b;
            }
            QProgressBar::chunk {
                background-color: #06b6d4;
                border-radius: 4px;
            }
        """)

        self.init_ui()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout()
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(24, 24, 24, 24)
        central.setLayout(main_layout)

        self.lte_layout, self.lte_input = self.create_input_row(
            icon_path="assets/lte.png",
            tooltip="LTE Data File (.xlsx)\nUse the FAC LTE Template from Performance Management UME",
            is_output=False
        )
        self.gsm_layout, self.gsm_input = self.create_input_row(
            icon_path="assets/gsm.png",
            tooltip="GSM Data File (.xlsx)\nUse the FAC GSM Template from Performance Management UME",
            is_output=False
        )
        # === BARIS BARU: NGI ===
        self.ngi_layout, self.ngi_input = self.create_input_row(
            icon_path="assets/ngi.png",
            tooltip="NGI Data File (.xlsx)\nFeature coming soon — will be used in future updates",
            is_output=False
        )
        self.cluster_layout, self.cluster_input = self.create_input_row(
            icon_path="assets/cluster.png",
            tooltip="Cluster Mapping File (.xlsx)\nMust contain columns: CLUSTER | TOWERID | CELLNAME | TXRX | 2G SITENAME | CAT",
            is_output=False
        )
        self.output_layout, self.output_input = self.create_input_row(
            icon_path="assets/output.png",
            tooltip="Output Directory\nChoose where the generated FAC reports will be saved",
            is_output=True
        )

        main_layout.addLayout(self.lte_layout)
        main_layout.addLayout(self.gsm_layout)
        main_layout.addLayout(self.ngi_layout)      # ← Baris baru
        main_layout.addLayout(self.cluster_layout)
        main_layout.addLayout(self.output_layout)

        main_layout.addSpacing(24)

        self.generate_btn = QPushButton("START")
        self.generate_btn.setMinimumHeight(44)
        self.generate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.generate_btn.clicked.connect(self.generate_reports)
        self.generate_btn.setToolTip("Start generating FAC reports for all clusters")
        main_layout.addWidget(self.generate_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        main_layout.addSpacing(8)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMinimumHeight(160)
        main_layout.addWidget(self.log_output, 1)

        # === RICH HTML TUTORIAL WITH SMALL ICONS (updated with NGI) ===
        icon_size = 16
        self.log("<b>FAC Report Generator — Quick Guide</b><br><br>")

        self.log(f"<img src='assets/lte.png' width='{icon_size}' height='{icon_size}'> "
                 "<b>LTE File:</b> Use the official FAC LTE Template<br>"
                 "  from Performance Management UME<br><br>")

        self.log(f"<img src='assets/gsm.png' width='{icon_size}' height='{icon_size}'> "
                 "<b>GSM File:</b> Use the official FAC GSM Template<br>"
                 "  from Performance Management UME<br><br>")

        self.log(f"<img src='assets/ngi.png' width='{icon_size}' height='{icon_size}'> "
                 "<b>NGI File:</b> NGI data input<br>"
                 "  (Feature coming soon)<br><br>")

        self.log(f"<img src='assets/cluster.png' width='{icon_size}' height='{icon_size}'> "
                 "<b>Cluster Mapping File:</b> Must contain these columns:<br>"
                 "  CLUSTER | TOWERID | CELLNAME | TXRX | 2G SITENAME<br><br>")

        self.log(f"<img src='assets/output.png' width='{icon_size}' height='{icon_size}'> "
                 "<b>Output Directory:</b> Choose where reports will be saved<br><br>")

        self.log("Hover over icons for more details • Ready to generate!")

    def create_input_row(self, icon_path: str, tooltip: str, is_output: bool):
        layout = QHBoxLayout()
        layout.setSpacing(10)

        label = QLabel()
        label.setFixedSize(32, 32)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setToolTip(tooltip)

        if Path(icon_path).exists():
            pixmap = QPixmap(icon_path).scaled(32, 32, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            label.setPixmap(pixmap)
        else:
            fallback = "📊" if "LTE" in tooltip else "📡" if "GSM" in tooltip else "🔧" if "NGI" in tooltip else "🌐" if "Cluster" in tooltip else "📂"
            label.setText(fallback)
            label.setFont(QFont("Segoe UI Emoji", 20))

        line_edit = QLineEdit()
        line_edit.setPlaceholderText("—")
        line_edit.setReadOnly(True)
        line_edit.setToolTip(tooltip)

        # Transparent browse button (icon only)
        browse_btn = QPushButton()
        browse_btn.setObjectName("BrowseButton")
        browse_btn.setFixedSize(40, 40)
        browse_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        browse_btn.setToolTip("Browse " + ("folder" if is_output else "file"))

        browse_icon_path = "assets/search.png"
        if Path(browse_icon_path).exists():
            browse_pixmap = QPixmap(browse_icon_path).scaled(28, 28, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            browse_btn.setIcon(QIcon(browse_pixmap))
            browse_btn.setIconSize(browse_pixmap.size())
        else:
            browse_btn.setText("📁")

        browse_btn.clicked.connect(lambda: self.browse(line_edit, is_output))

        layout.addWidget(label)
        layout.addWidget(line_edit, 1)
        layout.addWidget(browse_btn)

        return layout, line_edit

    def log(self, message: str):
        self.log_output.insertHtml(message + "<br>")
        self.log_output.ensureCursorVisible()

    def browse(self, line_edit: QLineEdit, is_folder: bool = False):
        if is_folder:
            path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
            )

        if path:
            line_edit.setText(path)
            name = Path(path).name if not is_folder else Path(path).name or path
            self.log(f"✓ Selected: <b>{name}</b>")

    def generate_reports(self):
        lte = self.lte_input.text()
        gsm = self.gsm_input.text()
        ngi = self.ngi_input.text()          # ← Bisa dipakai nanti
        cluster = self.cluster_input.text()
        output = self.output_input.text()

        if not all([lte, gsm, cluster, output]):
            QMessageBox.warning(self, "", "Please select all required files and output folder\n(NGI is optional for now)")
            return

        template_file = "./datatemplate.xlsx"
        if not Path(template_file).exists():
            QMessageBox.critical(self, "", "datatemplate.xlsx not found in project root")
            return

        self.generate_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.log("<br><b>Starting report generation...</b><br>")

        self.thread = ProcessThread(lte, gsm, ngi, cluster, output, template_file)
        self.thread.progress.connect(self.log)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def on_finished(self, success: bool, message: str):
        self.generate_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        if success:
            self.log("<br><span style='color:#06b6d4; font-weight:bold;'>✓ Complete!</span><br>")
            self.log(message)
            QMessageBox.information(self, "", f"{message}")
        else:
            self.log("<br><span style='color:#ef4444; font-weight:bold;'>✗ Failed</span><br>")
            self.log(message)
            QMessageBox.critical(self, "", "Report generation failed")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 9))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())