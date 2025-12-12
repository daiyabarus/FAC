"""
Application Layer - Report Generation Use Case
File: application/report_service.py

Changes:
- After parsing each file we now log how many rows were parsed and skipped (CSVParser prints this).
- We only work with parsed cells (which are guaranteed to have a cluster mapping).
- Grouping and KPI calculation remain cluster-based.
"""

from pathlib import Path
from datetime import datetime
from typing import Callable
from collections import defaultdict

from domain.models import CellData, ClusterReport, Technology
from domain.kpi_calculator import KPICalculator
from domain.kpi_targets import KPITargets
from infrastructure.csv_parser import CSVParser
from infrastructure.excel_writer import ExcelReportWriter


class ReportGenerationService:
    """Orchestrates the report generation process."""

    def __init__(self):
        self.csv_parser = CSVParser()
        self.excel_writer = ExcelReportWriter()
        self.kpi_calculator = KPICalculator()
        self.targets = KPITargets.get_all_targets()

    def generate_report_from_files(
        self,
        fdd_files: list[Path],
        tdd_files: list[Path],
        gsm_files: list[Path],
        tower_files: list[Path],
        output_path: Path,
        progress_callback: Callable[[str, int], None] = None
    ) -> None:
        """
        Generate FAC KPI report from individual CSV files.
        Strict mapping: only rows that are mapped via TOWERID/SITE_NAME are processed.
        """
        try:
            self._update_progress(progress_callback, "Loading tower mapping...", 5)
            if not tower_files:
                raise ValueError("TOWERID file is required for cluster mapping")

            # Load mapping (first tower file)
            self.csv_parser.load_tower_mapping(tower_files[0])
            self._update_progress(progress_callback, f"Loaded tower mapping from {tower_files[0].name}", 10)

            all_cells = []
            total_files = len(fdd_files) + len(tdd_files) + len(gsm_files)

            if total_files == 0:
                raise ValueError("No data files provided. Please select at least one FDD, TDD, or GSM file.")

            processed = 0

            # Parse FDD
            for fdd_file in fdd_files:
                self._update_progress(
                    progress_callback,
                    f"Parsing FDD file: {fdd_file.name}...",
                    10 + int((processed / total_files) * 30)
                )
                cells = self.csv_parser.parse_lte_csv(fdd_file, Technology.LTE_FDD)
                all_cells.extend(cells)
                self.log_message(f"✓ Parsed {len(cells)} mapped cells from {fdd_file.name}", progress_callback)
                processed += 1

            # Parse TDD
            for tdd_file in tdd_files:
                self._update_progress(
                    progress_callback,
                    f"Parsing TDD file: {tdd_file.name}...",
                    10 + int((processed / total_files) * 30)
                )
                cells = self.csv_parser.parse_lte_csv(tdd_file, Technology.LTE_TDD)
                all_cells.extend(cells)
                self.log_message(f"✓ Parsed {len(cells)} mapped cells from {tdd_file.name}", progress_callback)
                processed += 1

            # Parse GSM
            for gsm_file in gsm_files:
                self._update_progress(
                    progress_callback,
                    f"Parsing GSM file: {gsm_file.name}...",
                    10 + int((processed / total_files) * 30)
                )
                cells = self.csv_parser.parse_gsm_csv(gsm_file)
                all_cells.extend(cells)
                self.log_message(f"✓ Parsed {len(cells)} mapped cells from {gsm_file.name}", progress_callback)
                processed += 1

            if not all_cells:
                raise ValueError("No mapped data found in CSV files. Please verify TOWERID mapping and input files.")

            self._update_progress(
                progress_callback,
                f"Successfully parsed {len(all_cells)} total mapped cells from {total_files} file(s)",
                40
            )

            # Group by cluster and month
            self._update_progress(progress_callback, "Grouping data by cluster and month...", 50)
            grouped_data = self._group_by_cluster_and_month(all_cells)

            num_clusters = len(grouped_data)
            self.log_message(f"Found {num_clusters} cluster(s)", progress_callback)

            cluster_num = 0
            for cluster, cluster_data in grouped_data.items():
                cluster_num += 1
                progress = 50 + int((cluster_num / num_clusters) * 30)

                self._update_progress(
                    progress_callback,
                    f"Calculating KPIs for cluster: {cluster} ({cluster_num}/{num_clusters})...",
                    progress
                )

                report = self._generate_cluster_report(cluster, cluster_data)

                if num_clusters == 1:
                    cluster_output_path = output_path
                else:
                    cluster_output_path = output_path.parent / f"{output_path.stem}_{cluster}{output_path.suffix}"

                self._update_progress(
                    progress_callback,
                    f"Writing Excel report for {cluster}...",
                    progress + 5
                )

                self.excel_writer.write_report(report, cluster_output_path)
                self.log_message(f"✓ Report saved: {cluster_output_path.name}", progress_callback)

            self._update_progress(progress_callback, "Report generation complete!", 100)

        except Exception as e:
            self._update_progress(progress_callback, f"Error: {str(e)}", 0)
            raise

    def generate_report(
        self,
        input_folder: Path,
        output_path: Path,
        progress_callback: Callable[[str, int], None] = None
    ) -> None:
        """Legacy folder-based generation (keeps same strict mapping behaviour)."""
        try:
            self._update_progress(progress_callback, "Loading tower mapping...", 10)
            tower_files = list(input_folder.glob("*TOWER*.csv")) + list(input_folder.glob("*tower*.csv"))
            if tower_files:
                self.csv_parser.load_tower_mapping(tower_files[0])

            self._update_progress(progress_callback, "Parsing CSV files...", 20)
            all_cells = []

            fdd_files = list(input_folder.glob("*FDD*.csv")) + list(input_folder.glob("*fdd*.csv"))
            for fdd_file in fdd_files:
                cells = self.csv_parser.parse_lte_csv(fdd_file, Technology.LTE_FDD)
                all_cells.extend(cells)
                print(f"Parsed {len(cells)} mapped cells from {fdd_file.name}")

            tdd_files = list(input_folder.glob("*TDD*.csv")) + list(input_folder.glob("*tdd*.csv"))
            for tdd_file in tdd_files:
                cells = self.csv_parser.parse_lte_csv(tdd_file, Technology.LTE_TDD)
                all_cells.extend(cells)
                print(f"Parsed {len(cells)} mapped cells from {tdd_file.name}")

            gsm_files = list(input_folder.glob("*GSM*.csv")) + list(input_folder.glob("*gsm*.csv"))
            for gsm_file in gsm_files:
                cells = self.csv_parser.parse_gsm_csv(gsm_file)
                all_cells.extend(cells)
                print(f"Parsed {len(cells)} mapped cells from {gsm_file.name}")

            if not all_cells:
                raise ValueError("No mapped data found in CSV files")

            self._update_progress(progress_callback, "Grouping data by cluster and month...", 40)
            grouped_data = self._group_by_cluster_and_month(all_cells)

            self._update_progress(progress_callback, "Calculating KPIs...", 60)

            for cluster, cluster_data in grouped_data.items():
                cluster_output_path = output_path.parent / f"{output_path.stem}_{cluster}{output_path.suffix}"
                report = self._generate_cluster_report(cluster, cluster_data)

                self._update_progress(progress_callback, f"Writing report for {cluster}...", 80)
                self.excel_writer.write_report(report, cluster_output_path)
                print(f"Report generated: {cluster_output_path}")

            self._update_progress(progress_callback, "Report generation complete!", 100)

        except Exception as e:
            self._update_progress(progress_callback, f"Error: {str(e)}", 0)
            raise

    def _group_by_cluster_and_month(
        self,
        cells: list[CellData]
    ) -> dict[str, dict[int, list[CellData]]]:
        """Group cells by cluster and month."""
        grouped: dict[str, dict[int, list[CellData]]] = defaultdict(lambda: defaultdict(list))

        for cell in cells:
            month = cell.begin_time.month
            cluster = cell.cluster
            grouped[cluster][month].append(cell)

        return grouped

    def _generate_cluster_report(
        self,
        cluster: str,
        monthly_data: dict[int, list[CellData]]
    ) -> ClusterReport:
        """Generate report for a single cluster."""
        months = sorted(monthly_data.keys())
        all_results = []

        all_cluster_cells = []
        for cells in monthly_data.values():
            all_cluster_cells.extend(cells)

        unique_sites = len(set(cell.site_name for cell in all_cluster_cells))
        unique_cells = len(set(cell.cell_name for cell in all_cluster_cells))

        # compute month_ranges: for each month, find min begin_time and max end_time
        month_ranges: dict[int, tuple[datetime, datetime]] = {}
        for month in months:
            cells = monthly_data[month]
            if not cells:
                continue
            min_begin = min(c.begin_time for c in cells)
            max_end = max(c.end_time for c in cells)
            month_ranges[month] = (min_begin, max_end)

        for month in months:
            cells = monthly_data[month]

            gsm_cells = [c for c in cells if c.technology == Technology.GSM]
            lte_cells = [c for c in cells if c.technology in [Technology.LTE_FDD, Technology.LTE_TDD]]

            for target in self.targets:
                if target.technology == "2G RAN" and gsm_cells:
                    result = self.kpi_calculator.aggregate_kpi_results(
                        gsm_cells, target, month, cluster
                    )
                    all_results.append(result)

                elif target.technology == "4G RAN" and lte_cells:
                    result = self.kpi_calculator.aggregate_kpi_results(
                        lte_cells, target, month, cluster
                    )
                    all_results.append(result)

        return ClusterReport(
            cluster_name=cluster,
            months=months,
            site_count=unique_sites,
            cell_count=unique_cells,
            kpi_results=all_results,
            last_update=datetime.now(),
            month_ranges=month_ranges
        )

    @staticmethod
    def _update_progress(
        callback: Callable[[str, int], None],
        message: str,
        percentage: int
    ) -> None:
        """Update progress if callback is provided."""
        if callback:
            callback(message, percentage)

    @staticmethod
    def log_message(
        message: str,
        callback: Callable[[str, int], None]
    ) -> None:
        """Log a message without changing progress percentage."""
        if callback:
            print(message)
