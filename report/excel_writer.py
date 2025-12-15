"""Excel report writer"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from io import BytesIO
import base64
from datetime import datetime
from report.formatter import ExcelFormatter
from utils.helpers import format_month_name, format_date_range, get_three_month_range
from config.settings import LTEColumns, GSMColumns
from openpyxl.styles import Alignment
from utils.helpers import format_date_mdy


class ExcelReportWriter:
    """Write reports to Excel files"""

    def __init__(
        self, template_path, output_path, validation_results, kpi_data, transformed_data
    ):
        self.template_path = template_path
        self.output_path = output_path
        self.validation_results = validation_results
        self.kpi_data = kpi_data
        self.transformed_data = transformed_data
        self.formatter = ExcelFormatter()

        self.xlsmart_logo = None
        self.zte_logo = None

    def set_logos(self, xlsmart_logo, zte_logo):
        """Set logo base64 strings"""
        self.xlsmart_logo = xlsmart_logo
        self.zte_logo = zte_logo

    # def write_report(self, cluster):
    #     """Write report for a specific cluster"""
    #     print(f"\n=== Writing report for cluster: {cluster} ===")

    #     # Load template
    #     try:
    #         wb = load_workbook(self.template_path)
    #     except:
    #         # If template doesn't exist, create new workbook
    #         from openpyxl import Workbook

    #         wb = Workbook()
    #         wb.remove(wb.active)

    #     # Get cluster validation results
    #     cluster_results = self.validation_results.get(cluster, {})
    #     months = sorted(cluster_results.keys())

    #     if len(months) == 0:
    #         print(f"⚠ No data for cluster {cluster}")
    #         return

    #     # PERBAIKAN 2: Sort months by date
    #     month_dates = []
    #     for month in months:
    #         try:
    #             dt = pd.to_datetime(month, format="%b-%y")
    #             month_dates.append((month, dt))
    #         except:
    #             month_dates.append((month, pd.Timestamp("1900-01-01")))

    #     # Sort ascending (September, October, November)
    #     month_dates.sort(key=lambda x: x[1])
    #     sorted_months = [m[0] for m in month_dates]

    #     print(f"Months for report (sorted): {sorted_months}")

    #     # Write FAC sheet
    #     self._write_fac_sheet(wb, cluster, sorted_months, cluster_results)

    #     # Write Contributors sheet
    #     self._write_contributors_sheet(wb, cluster, sorted_months)

    #     # Write RAW sheets
    #     self._write_raw_sheets(wb, cluster)

    #     # Save workbook
    #     output_file = f"{self.output_path}/FAC_Report_{cluster}.xlsx"
    #     wb.save(output_file)
    #     print(f"✓ Report saved: {output_file}")
    def write_report(self, cluster):
        """Write report for a specific cluster"""
        print(f"\n=== Writing report for cluster: {cluster} ===")

        # Load template
        try:
            wb = load_workbook(self.template_path)
            print(f"✓ Template loaded: {self.template_path}")
        except Exception as e:
            print(f"⚠ Could not load template: {e}")
            from openpyxl import Workbook
            wb = Workbook()
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb['Sheet'])

        # Get cluster validation results
        cluster_results = self.validation_results.get(cluster, {})
        months = sorted(cluster_results.keys())

        if len(months) == 0:
            print(f"⚠ No data for cluster {cluster}")
            return

        # Sort months by date
        month_dates = []
        for month in months:
            try:
                dt = pd.to_datetime(month, format='%b-%y')
                month_dates.append((month, dt))
            except:
                month_dates.append((month, pd.Timestamp('1900-01-01')))

        month_dates.sort(key=lambda x: x[1])
        sorted_months = [m[0] for m in month_dates]

        print(f"Months for report (sorted): {sorted_months}")

        try:
            # Write FAC sheet
            self._write_fac_sheet(wb, cluster, sorted_months, cluster_results)
            print("✓ FAC sheet written")
        except Exception as e:
            print(f"✗ Error writing FAC sheet: {e}")
            import traceback
            traceback.print_exc()

        try:
            # Write Contributors sheet
            self._write_contributors_sheet(wb, cluster, sorted_months)
            print("✓ Contributors sheet written")
        except Exception as e:
            print(f"✗ Error writing Contributors sheet: {e}")
            import traceback
            traceback.print_exc()

        try:
            # Write RAW sheets
            self._write_raw_sheets(wb, cluster)
            print("✓ RAW sheets written")
        except Exception as e:
            print(f"✗ Error writing RAW sheets: {e}")
            import traceback
            traceback.print_exc()

        try:
            # TAMBAHAN: Write Charts sheet
            self._write_charts_sheet(wb, cluster)
            print("✓ Charts sheet written")
        except Exception as e:
            print(f"✗ Error writing Charts sheet: {e}")
            import traceback
            traceback.print_exc()

        # Save with better error handling
        try:
            output_file = f"{self.output_path}/FAC_Report_{cluster}.xlsx"

            import os
            if os.path.exists(output_file):
                os.remove(output_file)

            wb.save(output_file)
            wb.close()

            print(f"✓ Report saved: {output_file}")

            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"✓ File size: {file_size:,} bytes")

        except Exception as e:
            print(f"✗ Error saving report: {e}")
            import traceback
            traceback.print_exc()
            raise
    # TODO: Write Charts Sheet
    # def _write_charts_sheet(self, wb, cluster):
    #     """Write Charts sheet with KPI trend charts"""
    #     from report.chart_generator import ChartGenerator
    #     from openpyxl.drawing.image import Image as XLImage

    #     print(f"\n=== Generating charts for {cluster} ===")

    #     # Generate charts
    #     chart_gen = ChartGenerator(
    #         self.kpi_data, self.transformed_data, cluster)
    #     charts = chart_gen.generate_all_charts()

    #     if len(charts) == 0:
    #         print("⚠ No charts generated")
    #         return

    #     # Create Charts sheet
    #     if 'Charts' in wb.sheetnames:
    #         wb.remove(wb['Charts'])
    #     ws = wb.create_sheet('Charts')

    #     # Write title
    #     ws['A1'] = f'KPI Trend Charts - {cluster}'
    #     ws['A1'].font = self.formatter.header_font

    #     # Position charts in grid layout (2 columns)
    #     current_row = 3
    #     current_col = 1
    #     charts_per_row = 2
    #     chart_height = 30  # rows per chart
    #     chart_width = 10   # columns per chart

    #     chart_num = 0

    #     for chart_key, chart_base64 in charts.items():
    #         try:
    #             # Decode base64 image
    #             img_data = base64.b64decode(chart_base64)
    #             img = XLImage(BytesIO(img_data))

    #             # Resize to fit Excel
    #             img.width = 700
    #             img.height = 350

    #             # Calculate position
    #             # A, K, U, etc.
    #             col_letter = chr(65 + (current_col - 1) * chart_width)
    #             cell_position = f'{col_letter}{current_row}'

    #             # Add image to sheet
    #             ws.add_image(img, cell_position)

    #             chart_num += 1

    #             # Move to next position
    #             current_col += 1
    #             if current_col > charts_per_row:
    #                 current_col = 1
    #                 current_row += chart_height

    #             print(
    #                 f"✓ Added chart {chart_num}: {chart_key} at {cell_position}")

    #         except Exception as e:
    #             print(f"⚠ Could not add chart {chart_key}: {e}")

    #     print(f"✓ Written {chart_num} charts to Charts sheet")
    def _write_charts_sheet(self, wb, cluster):
        """Write Charts sheet with KPI trend charts"""
        from report.chart_generator import ChartGenerator
        from openpyxl.drawing.image import Image as XLImage

        print(f"\n=== Generating charts for {cluster} ===")

        # Generate charts
        chart_gen = ChartGenerator(
            self.kpi_data, self.transformed_data, cluster)
        charts = chart_gen.generate_all_charts()

        if len(charts) == 0:
            print("⚠ No charts generated")
            return

        # Create Charts sheet
        if 'Charts' in wb.sheetnames:
            wb.remove(wb['Charts'])
        ws = wb.create_sheet('Charts')

        # Write title
        ws['A1'] = f'KPI Trend Charts - {cluster}'
        ws['A1'].font = self.formatter.header_font_small

        # ============================================================
        # BAGIAN INI YANG PERLU DISESUAIKAN UNTUK POSISI CHART
        # ============================================================

        current_row = 3
        current_col = 1
        charts_per_row = 2

        # UBAH INI: Jarak vertikal antar chart (dalam rows)
        # Nilai lebih kecil = chart lebih rapat
        # Coba: 25, 22, 20, dll
        chart_height = 20  # <--- UBAH INI untuk rapatkan vertikal

        # UBAH INI: Jarak horizontal antar chart (dalam columns)
        # Nilai lebih besar = chart lebih bergeser ke kanan
        # Coba: 12, 13, 14, 15, dll
        chart_width = 13   # <--- UBAH INI untuk geser horizontal

        chart_num = 0

        for chart_key, chart_base64 in charts.items():
            try:
                # Decode base64 image
                img_data = base64.b64decode(chart_base64)
                img = XLImage(BytesIO(img_data))

                # UBAH INI: Ukuran chart di Excel
                # Width lebih besar = chart lebih lebar
                # Height lebih besar = chart lebih tinggi
                img.width = 700   # <--- UBAH INI untuk lebar chart
                img.height = 350  # <--- UBAH INI untuk tinggi chart

                # TAMBAHAN: Add border to image
                # Sayangnya openpyxl tidak support border langsung pada image
                # Tapi kita bisa add border pada cells di sekitar image

                # Calculate position
                col_offset = (current_col - 1) * chart_width
                col_letter = self._get_column_letter(col_offset + 1)
                cell_position = f'{col_letter}{current_row}'

                # Add image to sheet
                ws.add_image(img, cell_position)

                # TAMBAHAN: Add border cells around chart (optional visual guide)
                # Uncomment jika ingin border di sekitar chart
                # self._add_chart_border(ws, current_row, col_offset + 1, chart_height, chart_width)

                chart_num += 1

                # Move to next position
                current_col += 1
                if current_col > charts_per_row:
                    current_col = 1
                    current_row += chart_height

                print(
                    f"✓ Added chart {chart_num}: {chart_key} at {cell_position}")

            except Exception as e:
                print(f"⚠ Could not add chart {chart_key}: {e}")

        print(f"✓ Written {chart_num} charts to Charts sheet")

    def _get_column_letter(self, col_num):
        """Convert column number to Excel column letter"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + 65) + result
            col_num //= 26
        return result

    def _add_chart_border(self, ws, start_row, start_col, height, width):
        """Add border around chart area (optional)"""
        from openpyxl.styles import Border, Side

        thin_border = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )

        # Add border to corner cells only (untuk tidak menutupi chart)
        # Top-left corner
        ws.cell(row=start_row, column=start_col).border = thin_border

        # Top-right corner
        ws.cell(row=start_row, column=start_col +
                width - 1).border = thin_border

        # Bottom-left corner
        ws.cell(row=start_row + height - 1,
                column=start_col).border = thin_border

        # Bottom-right corner
        ws.cell(row=start_row + height - 1, column=start_col +
                width - 1).border = thin_border

    # TODO: Update 2: Tambahkan Info di Cell A14

    def _write_fac_sheet(self, wb, cluster, months, cluster_results):
        """Write FAC/Template sheet"""
        # Get or create sheet
        if 'Template' in wb.sheetnames:
            ws = wb['Template']
        elif 'FAC' in wb.sheetnames:
            ws = wb['FAC']
        else:
            ws = wb.create_sheet('FAC')

        # Add logos (with error handling)
        self._add_logos(ws)

        # Get date range for title
        lte_data = self.transformed_data['lte']
        gsm_data = self.transformed_data['gsm']
        cluster_lte = lte_data[lte_data['CLUSTER'] == cluster]
        cluster_gsm = gsm_data[gsm_data['CLUSTER'] == cluster]
        dates = pd.to_datetime(
            cluster_lte.iloc[:, LTEColumns.BEGIN_TIME]).unique()
        date_range = get_three_month_range(dates)

        # Write title
        ws['A6'] = f"FAC KPI Achievement Summary {cluster} - {date_range}"
        ws['A6'].font = self.formatter.header_font

        # PERBAIKAN: Tambahkan info di cell A14
        # Count unique tower IDs
        unique_towers = cluster_lte['TOWER_ID'].dropna().nunique()

        # Count unique 4G cells
        unique_4g_cells = cluster_lte.iloc[:,
                                           LTEColumns.CELL_NAME].dropna().nunique()

        # Count unique 2G cells
        unique_2g_cells = cluster_gsm.iloc[:,
                                           GSMColumns.BTS_NAME].dropna().nunique()

        # Get last date (LAST CLUSTER PAC DATE)
        if len(dates) > 0:
            last_date = pd.to_datetime(max(dates))
            last_pac_date = last_date.strftime('%d %B %Y')
        else:
            last_pac_date = "N/A"

        # PERBAIKAN MASALAH 2: Jangan merge cell (sudah merged di template A14:A63)
        # Langsung write ke A14 saja
        info_text = (
            f"City name: {cluster}\n"
            f"Site Number: {unique_towers}\n"
            f"Cell Number: {unique_4g_cells} 4G & {unique_2g_cells} 2G\n"
            f"LAST CLUSTER PAC DATE: {last_pac_date}"
        )

        ws['A14'] = info_text
        ws['A14'].alignment = Alignment(
            wrap_text=True, vertical='top', horizontal='left')

        # JANGAN merge - template sudah merge A14:A63
        # Hapus baris ini: ws.merge_cells('A14:L14')

        # Write month headers (row 12) - already sorted
        for idx, month in enumerate(months[:3]):  # Max 3 months
            col = 13 + (idx * 2)  # M=13, O=15, Q=17 columns

            # Convert month format to full name (Sep-25 -> September)
            try:
                dt = pd.to_datetime(month, format='%b-%y')
                month_name = dt.strftime('%B')  # Full month name
            except:
                month_name = month

            ws.cell(row=12, column=col).value = month_name
            ws.cell(row=12, column=col).font = self.formatter.header_font

            print(f"Month {idx+1} header: {month_name} (column {col})")

        # Write date ranges (row 13)
        for idx, month in enumerate(months[:3]):
            col = 13 + (idx * 2)
            month_data = cluster_lte[cluster_lte['MONTH'] == month]
            if len(month_data) > 0:
                dates_month = pd.to_datetime(
                    month_data.iloc[:, LTEColumns.BEGIN_TIME])
                date_range_str = format_date_range(
                    dates_month.min(), dates_month.max())
                ws.cell(row=13, column=col).value = date_range_str

        # Write KPI results
        self._write_kpi_results(ws, months[:3], cluster_results)

        # Write overall results (row 63)
        for idx, month in enumerate(months[:3]):
            col = 13 + (idx * 2)
            month_results = cluster_results[month]
            overall_pass = month_results.get('overall_pass', False)

            result_cell = ws.cell(row=63, column=col)
            self.formatter.format_pass_fail(result_cell, overall_pass)

    def _write_kpi_results(self, ws, months, cluster_results):
        """Write KPI results to FAC sheet"""
        # Define KPI mapping to rows
        kpi_mapping = {
            # GSM KPIs
            ("gsm", "cssr"): 14,
            ("gsm", "sdcch_sr"): 15,
            ("gsm", "drop_rate"): 16,
            # LTE KPIs
            ("lte", "session_ssr"): 17,
            ("lte", "rach_sr_high"): 18,
            ("lte", "rach_sr_low"): 19,
            ("lte", "ho_sr"): 20,
            ("lte", "erab_drop"): 21,
            ("lte", "dl_thp_high"): 22,
            ("lte", "dl_thp_low"): 23,
            ("lte", "ul_thp_high"): 24,
            ("lte", "ul_thp_low"): 25,
            ("lte", "ul_ploss"): 26,
            ("lte", "dl_ploss"): 27,
            ("lte", "cqi"): 28,
            ("lte", "mimo_rank2_high"): 29,
            ("lte", "mimo_rank2_low"): 30,
            ("lte", "ul_rssi"): 31,
            ("lte", "latency_low"): 32,
            ("lte", "latency_medium"): 33,
            ("lte", "ltc_non_cap"): 34,
            ("lte", "overlap_rate"): 35,
            # Spectral Efficiency
            ("lte", "se_850_2t2r"): 36,
            ("lte", "se_900_2t2r"): 37,
            ("lte", "se_2100_2t2r"): 38,
            ("lte", "se_1800_2t2r"): 39,
            ("lte", "se_1800_4t4r"): 40,
            ("lte", "se_2100_4t4r"): 41,
            ("lte", "se_2300_32t32r"): 43,
            # VoLTE
            ("lte", "volte_cssr"): 45,
            ("lte", "volte_drop"): 46,
            ("lte", "srvcc_sr"): 47,
        }

        # Write results for each month (already sorted)
        for month_idx, month in enumerate(months):
            value_col = 13 + (month_idx * 2)  # M=13, O=15, Q=17
            result_col = value_col + 1  # N=14, P=16, R=18

            month_results = cluster_results[month]

            # Process GSM KPIs
            for kpi_key, kpi_result in month_results.get("gsm", {}).items():
                row_key = ("gsm", kpi_key)
                if row_key in kpi_mapping:
                    row = kpi_mapping[row_key]

                    # Write value
                    value_cell = ws.cell(row=row, column=value_col)
                    self.formatter.format_value(
                        value_cell, kpi_result["value"], "0.00")

                    # Write pass/fail
                    result_cell = ws.cell(row=row, column=result_col)
                    self.formatter.format_pass_fail(
                        result_cell, kpi_result["pass"])

            # Process LTE KPIs
            for kpi_key, kpi_result in month_results.get("lte", {}).items():
                row_key = ("lte", kpi_key)
                if row_key in kpi_mapping:
                    row = kpi_mapping[row_key]

                    # Write value
                    value_cell = ws.cell(row=row, column=value_col)
                    self.formatter.format_value(
                        value_cell, kpi_result["value"], "0.00")

                    # Write pass/fail
                    result_cell = ws.cell(row=row, column=result_col)
                    self.formatter.format_pass_fail(
                        result_cell, kpi_result["pass"])

    # TODO: Remove Duplicate Method
    def _write_contributors_sheet(self, wb, cluster, months):
        """Write Contributors sheet with cells that failed"""
        if 'Contributors' in wb.sheetnames:
            ws = wb['Contributors']
            wb.remove(ws)

        ws = wb.create_sheet('Contributors')

        # Write headers
        headers = ['Month', 'Clause Type', 'Name', 'Reference', 'TOWER_ID',
                   'CELL_NAME', 'Value', 'Baseline', 'Status']
        for col_idx, header in enumerate(headers, 1):
            self.formatter.format_header_small(
                ws.cell(row=1, column=col_idx), header)

        # Collect failed cells
        contributors = []

        lte_data = self.kpi_data['lte']
        gsm_data = self.kpi_data['gsm']

        # Filter by cluster
        lte_cluster = lte_data[lte_data['CLUSTER'] == cluster]
        gsm_cluster = gsm_data[gsm_data['CLUSTER'] == cluster]

        # Check GSM failures for each month
        for month in months:
            gsm_month = gsm_cluster[gsm_cluster['MONTH'] == month]

            # CSSR failures
            cssr_fails = gsm_month[gsm_month['CSSR'] < 98.5]
            for _, row in cssr_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Accessibility',
                    'Name': 'Call Setup Success Rate',
                    'Reference': '2G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[GSMColumns.BTS_NAME],
                    'Value': row['CSSR'],
                    'Baseline': 98.5,
                    'Status': 'FAIL'
                })

            # SDCCH failures
            sdcch_fails = gsm_month[gsm_month['SDCCH_SR'] < 98.5]
            for _, row in sdcch_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Accessibility',
                    'Name': 'SDCCH Success rate',
                    'Reference': '2G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[GSMColumns.BTS_NAME],
                    'Value': row['SDCCH_SR'],
                    'Baseline': 98.5,
                    'Status': 'FAIL'
                })

            # Drop Rate failures
            drop_fails = gsm_month[gsm_month['DROP_RATE'] >= 2]
            for _, row in drop_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Retainability',
                    'Name': 'Perceive Drop Rate',
                    'Reference': '2G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[GSMColumns.BTS_NAME],
                    'Value': row['DROP_RATE'],
                    'Baseline': 2,
                    'Status': 'FAIL'
                })

        # Check LTE failures for each month
        for month in months:
            lte_month = lte_cluster[lte_cluster['MONTH'] == month]

            # Session SSR failures
            session_fails = lte_month[lte_month['SESSION_SSR'] < 99]
            for _, row in session_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Accessibility',
                    'Name': 'Session Setup Success Rate',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['SESSION_SSR'],
                    'Baseline': 99,
                    'Status': 'FAIL'
                })

            # RACH SR failures (< 85%)
            rach_fails = lte_month[lte_month['RACH_SR'] < 85]
            for _, row in rach_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Accessibility',
                    'Name': 'RACH Success Rate (< 85%)',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['RACH_SR'],
                    'Baseline': 85,
                    'Status': 'FAIL'
                })

            # RACH SR critical failures (< 55%)
            rach_critical = lte_month[lte_month['RACH_SR'] < 55]
            for _, row in rach_critical.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Accessibility',
                    'Name': 'RACH Success Rate (< 55%)',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['RACH_SR'],
                    'Baseline': 55,
                    'Status': 'FAIL'
                })

            # Handover SR failures
            ho_fails = lte_month[lte_month['HO_SR'] < 97]
            for _, row in ho_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Mobility',
                    'Name': 'Handover Success Rate Inter and Intra-Frequency',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['HO_SR'],
                    'Baseline': 97,
                    'Status': 'FAIL'
                })

            # E-RAB Drop failures
            erab_fails = lte_month[lte_month['ERAB_DROP'] >= 2]
            for _, row in erab_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Retainability',
                    'Name': 'E-RAB Drop Rate',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['ERAB_DROP'],
                    'Baseline': 2,
                    'Status': 'FAIL'
                })

            # DL Throughput failures (< 3 Mbps)
            dl_fails = lte_month[lte_month['DL_THP'] < 3]
            for _, row in dl_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Integrity',
                    'Name': 'Downlink User Throughput (< 3 Mbps)',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['DL_THP'],
                    'Baseline': 3,
                    'Status': 'FAIL'
                })

            # DL Throughput critical failures (< 1 Mbps)
            dl_critical = lte_month[lte_month['DL_THP'] < 1]
            for _, row in dl_critical.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Integrity',
                    'Name': 'Downlink User Throughput (< 1 Mbps)',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['DL_THP'],
                    'Baseline': 1,
                    'Status': 'FAIL'
                })

            # UL Throughput failures (< 1 Mbps)
            ul_fails = lte_month[lte_month['UL_THP'] < 1]
            for _, row in ul_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Integrity',
                    'Name': 'Uplink User Throughput (< 1 Mbps)',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['UL_THP'],
                    'Baseline': 1,
                    'Status': 'FAIL'
                })

            # UL Throughput critical failures (< 0.256 Mbps)
            ul_critical = lte_month[lte_month['UL_THP'] < 0.256]
            for _, row in ul_critical.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Integrity',
                    'Name': 'Uplink User Throughput (< 0.256 Mbps)',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['UL_THP'],
                    'Baseline': 0.256,
                    'Status': 'FAIL'
                })

            # UL Packet Loss failures (>= 0.85%)
            ul_ploss_fails = lte_month[lte_month['UL_PLOSS'] >= 0.85]
            for _, row in ul_ploss_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Integrity',
                    'Name': 'UL Packet Loss (PDCP)',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['UL_PLOSS'],
                    'Baseline': 0.85,
                    'Status': 'FAIL'
                })

            # DL Packet Loss failures (>= 0.10%)
            dl_ploss_fails = lte_month[lte_month['DL_PLOSS'] >= 0.10]
            for _, row in dl_ploss_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Integrity',
                    'Name': 'DL Packet Loss (PDCP)',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['DL_PLOSS'],
                    'Baseline': 0.10,
                    'Status': 'FAIL'
                })

            # CQI failures
            cqi_fails = lte_month[lte_month['CQI'] < 7]
            for _, row in cqi_fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': 'Integrity',
                    'Name': 'CQI',
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row['CQI'],
                    'Baseline': 7,
                    'Status': 'FAIL'
                })

        # PERBAIKAN: Convert to DataFrame and remove duplicates
        if len(contributors) > 0:
            df_contributors = pd.DataFrame(contributors)
            # Remove duplicates based on all columns except Value (might vary slightly)
            df_contributors = df_contributors.drop_duplicates(
                subset=['Month', 'Clause Type', 'Name', 'Reference',
                        'TOWER_ID', 'CELL_NAME', 'Baseline'],
                keep='first'
            )
            contributors = df_contributors.to_dict('records')

        # Write contributors to sheet
        for row_idx, contrib in enumerate(contributors, 2):
            ws.cell(row=row_idx, column=1).value = contrib['Month']
            ws.cell(row=row_idx, column=2).value = contrib['Clause Type']
            ws.cell(row=row_idx, column=3).value = contrib['Name']
            ws.cell(row=row_idx, column=4).value = contrib['Reference']
            ws.cell(row=row_idx, column=5).value = contrib['TOWER_ID']
            ws.cell(row=row_idx, column=6).value = contrib['CELL_NAME']

            value_cell = ws.cell(row=row_idx, column=7)
            self.formatter.format_value(value_cell, contrib['Value'], '0.00')

            ws.cell(row=row_idx, column=8).value = contrib['Baseline']

            status_cell = ws.cell(row=row_idx, column=9)
            self.formatter.format_pass_fail(status_cell, False)

        print(
            f"✓ Written {len(contributors)} contributors (duplicates removed)")

    def _write_raw_sheets(self, wb, cluster):
        """Write RAW 2G and RAW 4G sheets"""

        # RAW 2G
        gsm_data = self.kpi_data['gsm']
        gsm_cluster = gsm_data[gsm_data['CLUSTER'] == cluster].copy()

        if len(gsm_cluster) > 0:
            # Select columns
            gsm_raw = gsm_cluster[[
                gsm_cluster.columns[GSMColumns.BEGIN_TIME],
                'TOWER_ID',
                gsm_cluster.columns[GSMColumns.BTS_NAME],
                'CSSR',
                'SDCCH_SR',
                'DROP_RATE'
            ]].copy()

            gsm_raw.columns = ['BEGIN_TIME', 'TOWER_ID', 'BTS_NAME',
                               'CSSR', 'SDCCH_SR', 'DROP_RATE']

            # PERBAIKAN: Remove duplicates
            gsm_raw = gsm_raw.drop_duplicates()

            # PERBAIKAN: Convert BEGIN_TIME to m/d/yyyy format
            gsm_raw['BEGIN_TIME'] = pd.to_datetime(
                gsm_raw['BEGIN_TIME']).dt.strftime('%-m/%-d/%Y')

            # Write to sheet
            if 'RAW 2G' in wb.sheetnames:
                wb.remove(wb['RAW 2G'])
            ws_gsm = wb.create_sheet('RAW 2G')

            # Write headers
            for c_idx, col in enumerate(gsm_raw.columns, 1):
                self.formatter.format_header_small(
                    ws_gsm.cell(row=1, column=c_idx), col)

            # Write data
            for r_idx, row in enumerate(gsm_raw.itertuples(index=False), 2):
                for c_idx, value in enumerate(row, 1):
                    cell = ws_gsm.cell(row=r_idx, column=c_idx)
                    cell.value = value
                    if c_idx > 3:  # Numeric columns
                        cell.number_format = '0.00'

            print(
                f"✓ Written RAW 2G: {len(gsm_raw)} records (duplicates removed)")

        # RAW 4G
        lte_data = self.kpi_data['lte']
        lte_cluster = lte_data[lte_data['CLUSTER'] == cluster].copy()

        if len(lte_cluster) > 0:
            # Select columns
            lte_raw = lte_cluster[[
                lte_cluster.columns[LTEColumns.BEGIN_TIME],
                'TOWER_ID',
                lte_cluster.columns[LTEColumns.CELL_NAME],
                'LTE_BAND',
                'SESSION_SSR',
                'RACH_SR',
                'HO_SR',
                'ERAB_DROP',
                'DL_THP',
                'UL_THP',
                'UL_PLOSS',
                'DL_PLOSS',
                'CQI',
                'MIMO_RANK2',
                'UL_RSSI',
                'LATENCY',
                'LTC_NON_CAP',
                'OVERLAP_RATE',
                'SPECTRAL_EFF',
                'VOLTE_CSSR',
                'VOLTE_DROP',
                'SRVCC_SR'
            ]].copy()

            lte_raw.columns = ['BEGIN_TIME', 'TOWER_ID', 'CELL_NAME', 'BAND',
                               'SESSION_SSR', 'RACH_SR', 'HO_SR', 'ERAB_DROP',
                               'DL_THP', 'UL_THP', 'UL_PLOSS', 'DL_PLOSS',
                               'CQI', 'MIMO_RANK2', 'UL_RSSI', 'LATENCY',
                               'LTC_NON_CAP', 'OVERLAP_RATE', 'SPECTRAL_EFF',
                               'VOLTE_CSSR', 'VOLTE_DROP', 'SRVCC_SR']

            # PERBAIKAN: Remove duplicates
            lte_raw = lte_raw.drop_duplicates()

            # PERBAIKAN: Convert BEGIN_TIME to m/d/yyyy format
            lte_raw['BEGIN_TIME'] = pd.to_datetime(
                lte_raw['BEGIN_TIME']).dt.strftime('%-m/%-d/%Y')

            # Write to sheet
            if 'RAW 4G' in wb.sheetnames:
                wb.remove(wb['RAW 4G'])
            ws_lte = wb.create_sheet('RAW 4G')

            # Write headers
            for c_idx, col in enumerate(lte_raw.columns, 1):
                self.formatter.format_header_small(
                    ws_lte.cell(row=1, column=c_idx), col)

            # Write data
            for r_idx, row in enumerate(lte_raw.itertuples(index=False), 2):
                for c_idx, value in enumerate(row, 1):
                    cell = ws_lte.cell(row=r_idx, column=c_idx)
                    cell.value = value
                    if c_idx > 4:  # Numeric columns
                        cell.number_format = '0.00'

            print(
                f"✓ Written RAW 4G: {len(lte_raw)} records (duplicates removed)")

    def _add_logos(self, ws):
        """Add logos to worksheet"""
        try:
            if self.xlsmart_logo:
                img_data = base64.b64decode(self.xlsmart_logo)
                img = XLImage(BytesIO(img_data))
                img.width, img.height = 120, 54
                ws.add_image(img, "A1")

            if self.zte_logo:
                img_data = base64.b64decode(self.zte_logo)
                img = XLImage(BytesIO(img_data))
                img.width, img.height = 100, 50
                ws.add_image(img, "S1")
        except Exception as e:
            print(f"⚠ Warning: Could not add logos: {e}")
