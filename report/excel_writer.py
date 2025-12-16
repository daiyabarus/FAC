"""Excel report writer"""

import os
import base64
import traceback
from io import BytesIO
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Side

from report.formatter import ExcelFormatter
from report.chart_generator import ChartGenerator
from utils.helpers import format_date_range, get_three_month_range
from config.settings import LTEColumns, GSMColumns


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

    def write_report(self, cluster):
        """Write report for a specific cluster"""
        print(f"\n=== Writing report for cluster: {cluster} ===")

        # Load template
        wb = self._load_template()

        # Get and sort months
        sorted_months = self._get_sorted_months(cluster)
        if not sorted_months:
            print(f"⚠ No data for cluster {cluster}")
            return

        # Write all sheets
        self._write_all_sheets(wb, cluster, sorted_months)

        # Save workbook
        self._save_workbook(wb, cluster)

    def _load_template(self):
        """Load Excel template or create new workbook"""
        try:
            wb = load_workbook(self.template_path)
            print(f"✓ Template loaded: {self.template_path}")
        except Exception as e:
            print(f"⚠ Could not load template: {e}")
            wb = Workbook()
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb['Sheet'])
        return wb

    def _get_sorted_months(self, cluster):
        """Get and sort months for the cluster"""
        cluster_results = self.validation_results.get(cluster, {})
        months = sorted(cluster_results.keys())

        if not months:
            return []

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
        return sorted_months

    def _write_all_sheets(self, wb, cluster, sorted_months):
        """Write all sheets with error handling"""
        cluster_results = self.validation_results.get(cluster, {})

        sheets = [
            ('FAC', lambda: self._write_fac_sheet(
                wb, cluster, sorted_months, cluster_results)),
            ('Contributors', lambda: self._write_contributors_sheet(
                wb, cluster, sorted_months)),
            ('RAW', lambda: self._write_raw_sheets(wb, cluster)),
            ('Charts', lambda: self._write_charts_sheet(wb, cluster))
        ]

        for sheet_name, write_func in sheets:
            try:
                write_func()
                print(f"✓ {sheet_name} sheet written")
            except Exception as e:
                print(f"✗ Error writing {sheet_name} sheet: {e}")
                traceback.print_exc()

    def _save_workbook(self, wb, cluster):
        """Save workbook with proper naming and error handling"""
        try:
            # Generate timestamp for filename
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f"{self.output_path}/FAC_Report_{cluster}_{timestamp}.xlsx"

            # Remove existing file if present
            if os.path.exists(output_file):
                os.remove(output_file)

            # Save workbook
            wb.save(output_file)
            wb.close()

            print(f"✓ Report saved: {output_file}")

            # Verify file creation
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"✓ File size: {file_size:,} bytes")

        except Exception as e:
            print(f"✗ Error saving report: {e}")
            traceback.print_exc()
            raise

    def _write_charts_sheet(self, wb, cluster):
        """Write Charts sheet with KPI trend charts"""
        print(f"\n=== Generating charts for {cluster} ===")

        # Generate charts
        chart_gen = ChartGenerator(
            self.kpi_data, self.transformed_data, cluster)
        charts = chart_gen.generate_all_charts()

        if not charts:
            print("⚠ No charts generated")
            return

        # Create Charts sheet
        if 'Charts' in wb.sheetnames:
            wb.remove(wb['Charts'])
        ws = wb.create_sheet('Charts')

        # Write title
        ws['A1'] = f'KPI Trend Charts - {cluster}'
        ws['A1'].font = self.formatter.header_font_small

        # Chart layout configuration
        current_row = 3
        current_col = 1
        charts_per_row = 2
        chart_height = 20
        chart_width = 11.5
        chart_num = 0

        # Add charts to sheet
        for chart_key, chart_base64 in charts.items():
            try:
                # Decode and create image
                img_data = base64.b64decode(chart_base64)
                img = XLImage(BytesIO(img_data))
                img.width = 700
                img.height = 350

                # Calculate position
                col_offset = (current_col - 1) * chart_width
                col_letter = self._get_column_letter(col_offset + 1)
                cell_position = f'{col_letter}{current_row}'

                # Add image to sheet
                ws.add_image(img, cell_position)

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
        thin_border = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )

        # Add border to corner cells only
        corners = [
            (start_row, start_col),
            (start_row, start_col + width - 1),
            (start_row + height - 1, start_col),
            (start_row + height - 1, start_col + width - 1)
        ]

        for row, col in corners:
            ws.cell(row=row, column=col).border = thin_border

    def _write_fac_sheet(self, wb, cluster, months, cluster_results):
        """Write FAC/Template sheet"""
        # Get or create sheet
        if 'Template' in wb.sheetnames:
            ws = wb['Template']
        elif 'FAC' in wb.sheetnames:
            ws = wb['FAC']
        else:
            ws = wb.create_sheet('FAC')

        # Add logos
        self._add_logos(ws)

        # Get cluster data
        lte_data = self.transformed_data['lte']
        gsm_data = self.transformed_data['gsm']
        cluster_lte = lte_data[lte_data['CLUSTER'] == cluster]
        cluster_gsm = gsm_data[gsm_data['CLUSTER'] == cluster]

        # Get date range for title
        dates = pd.to_datetime(
            cluster_lte.iloc[:, LTEColumns.BEGIN_TIME]).unique()
        date_range = get_three_month_range(dates)

        # Write title
        ws['A6'] = f"FAC KPI Achievement {cluster} - {date_range}"
        ws['A6'].font = self.formatter.header_font

        # Write cluster information
        self._write_cluster_info(ws, cluster, cluster_lte, cluster_gsm, dates)

        # Write month headers
        self._write_month_headers(ws, months[:3], cluster_lte)

        # Write KPI results
        self._write_kpi_results(ws, months[:3], cluster_results)

        # Write overall results
        self._write_overall_results(ws, months[:3], cluster_results)

    def _write_cluster_info(self, ws, cluster, cluster_lte, cluster_gsm, dates):
        """Write cluster information to A14"""
        unique_towers = cluster_lte['TOWER_ID'].dropna().nunique()
        unique_4g_cells = cluster_lte.iloc[:,
                                           LTEColumns.CELL_NAME].dropna().nunique()
        unique_2g_cells = cluster_gsm.iloc[:,
                                           GSMColumns.BTS_NAME].dropna().nunique()

        if len(dates) > 0:
            last_date = pd.to_datetime(max(dates))
            last_pac_date = last_date.strftime('%d %B %Y')
        else:
            last_pac_date = "N/A"

        info_text = (
            f"City name: {cluster}\n"
            f"Site Number: {unique_towers}\n"
            f"Cell Number: {unique_4g_cells} 4G & {unique_2g_cells} 2G\n"
            f"LAST CLUSTER PAC DATE: {last_pac_date}"
        )

        ws['A14'] = info_text
        ws['A14'].alignment = Alignment(
            wrap_text=True, vertical='top', horizontal='left')

    def _write_month_headers(self, ws, months, cluster_lte):
        """Write month headers to row 12 and date ranges to row 13"""
        for idx, month in enumerate(months):
            col = 13 + (idx * 2)

            # Write month name
            try:
                dt = pd.to_datetime(month, format='%b-%y')
                month_name = dt.strftime('%B')
            except:
                month_name = month

            ws.cell(row=12, column=col).value = month_name
            ws.cell(row=12, column=col).font = self.formatter.header_font

            # Write date range
            month_data = cluster_lte[cluster_lte['MONTH'] == month]
            if len(month_data) > 0:
                dates_month = pd.to_datetime(
                    month_data.iloc[:, LTEColumns.BEGIN_TIME])
                date_range_str = format_date_range(
                    dates_month.min(), dates_month.max())
                ws.cell(row=13, column=col).value = date_range_str

            print(f"Month {idx+1} header: {month_name} (column {col})")

    def _write_overall_results(self, ws, months, cluster_results):
        """Write overall pass/fail results to row 63"""
        for idx, month in enumerate(months):
            col = 13 + (idx * 2)
            month_results = cluster_results[month]
            overall_pass = month_results.get('overall_pass', False)

            result_cell = ws.cell(row=63, column=col)
            self.formatter.format_pass_fail(result_cell, overall_pass)

    def _write_kpi_results(self, ws, months, cluster_results):
        """Write KPI results to FAC sheet"""
        kpi_mapping = self._get_kpi_mapping()

        for month_idx, month in enumerate(months):
            value_col = 13 + (month_idx * 2)
            result_col = value_col + 1

            month_results = cluster_results[month]

            # Process GSM KPIs
            self._write_tech_kpi_results(
                ws, month_results.get(
                    "gsm", {}), kpi_mapping, "gsm", value_col, result_col
            )

            # Process LTE KPIs
            self._write_tech_kpi_results(
                ws, month_results.get(
                    "lte", {}), kpi_mapping, "lte", value_col, result_col
            )

    def _get_kpi_mapping(self):
        """Get KPI to row mapping"""
        return {
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

    def _write_tech_kpi_results(self, ws, tech_results, kpi_mapping, tech, value_col, result_col):
        """Write KPI results for a specific technology"""
        for kpi_key, kpi_result in tech_results.items():
            row_key = (tech, kpi_key)
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

    def _write_contributors_sheet(self, wb, cluster, months):
        """Write Contributors sheet with cells that failed"""
        if 'Contributors' in wb.sheetnames:
            wb.remove(wb['Contributors'])

        ws = wb.create_sheet('Contributors')

        # Write headers
        headers = ['Month', 'Clause Type', 'Name', 'Reference', 'TOWER_ID',
                   'CELL_NAME', 'Value', 'Baseline', 'Status']
        for col_idx, header in enumerate(headers, 1):
            self.formatter.format_header_small(
                ws.cell(row=1, column=col_idx), header)

        # Collect contributors
        contributors = self._collect_contributors(cluster, months)

        # Remove duplicates
        contributors = self._remove_duplicate_contributors(contributors)

        # Write contributors
        self._write_contributor_rows(ws, contributors)

        print(
            f"✓ Written {len(contributors)} contributors (duplicates removed)")

    def _collect_contributors(self, cluster, months):
        """Collect all failed cells as contributors"""
        contributors = []
        lte_data = self.kpi_data['lte']
        gsm_data = self.kpi_data['gsm']

        lte_cluster = lte_data[lte_data['CLUSTER'] == cluster]
        gsm_cluster = gsm_data[gsm_data['CLUSTER'] == cluster]

        for month in months:
            # Collect GSM failures
            contributors.extend(
                self._collect_gsm_contributors(gsm_cluster, month))

            # Collect LTE failures
            contributors.extend(
                self._collect_lte_contributors(lte_cluster, month))

        return contributors

    def _collect_gsm_contributors(self, gsm_cluster, month):
        """Collect GSM failed cells"""
        contributors = []
        gsm_month = gsm_cluster[gsm_cluster['MONTH'] == month]

        gsm_checks = [
            ('CSSR', 98.5, '<', 'Accessibility', 'Call Setup Success Rate'),
            ('SDCCH_SR', 98.5, '<', 'Accessibility', 'SDCCH Success rate'),
            ('DROP_RATE', 2, '>=', 'Retainability', 'Perceive Drop Rate')
        ]

        for kpi_col, baseline, operator, clause, name in gsm_checks:
            if operator == '<':
                fails = gsm_month[gsm_month[kpi_col] < baseline]
            else:
                fails = gsm_month[gsm_month[kpi_col] >= baseline]

            for _, row in fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': clause,
                    'Name': name,
                    'Reference': '2G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[GSMColumns.BTS_NAME],
                    'Value': row[kpi_col],
                    'Baseline': baseline,
                    'Status': 'FAIL'
                })

        return contributors

    def _collect_lte_contributors(self, lte_cluster, month):
        """Collect LTE failed cells"""
        contributors = []
        lte_month = lte_cluster[lte_cluster['MONTH'] == month]

        lte_checks = [
            ('SESSION_SSR', 99, '<', 'Accessibility', 'Session Setup Success Rate'),
            ('RACH_SR', 85, '<', 'Accessibility', 'RACH Success Rate (< 85%)'),
            ('RACH_SR', 55, '<', 'Accessibility', 'RACH Success Rate (< 55%)'),
            ('HO_SR', 97, '<', 'Mobility',
             'Handover Success Rate Inter and Intra-Frequency'),
            ('ERAB_DROP', 2, '>=', 'Retainability', 'E-RAB Drop Rate'),
            ('DL_THP', 3, '<', 'Integrity', 'Downlink User Throughput (< 3 Mbps)'),
            ('DL_THP', 1, '<', 'Integrity', 'Downlink User Throughput (< 1 Mbps)'),
            ('UL_THP', 1, '<', 'Integrity', 'Uplink User Throughput (< 1 Mbps)'),
            ('UL_THP', 0.256, '<', 'Integrity',
             'Uplink User Throughput (< 0.256 Mbps)'),
            ('UL_PLOSS', 0.85, '>=', 'Integrity', 'UL Packet Loss (PDCP)'),
            ('DL_PLOSS', 0.10, '>=', 'Integrity', 'DL Packet Loss (PDCP)'),
            ('CQI', 7, '<', 'Integrity', 'CQI')
        ]

        for kpi_col, baseline, operator, clause, name in lte_checks:
            if operator == '<':
                fails = lte_month[lte_month[kpi_col] < baseline]
            else:
                fails = lte_month[lte_month[kpi_col] >= baseline]

            for _, row in fails.iterrows():
                contributors.append({
                    'Month': month,
                    'Clause Type': clause,
                    'Name': name,
                    'Reference': '4G RAN',
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTEColumns.CELL_NAME],
                    'Value': row[kpi_col],
                    'Baseline': baseline,
                    'Status': 'FAIL'
                })

        return contributors

    def _remove_duplicate_contributors(self, contributors):
        """Remove duplicate contributors"""
        if not contributors:
            return []

        df = pd.DataFrame(contributors)
        df = df.drop_duplicates(
            subset=['Month', 'Clause Type', 'Name', 'Reference',
                    'TOWER_ID', 'CELL_NAME', 'Baseline'],
            keep='first'
        )
        return df.to_dict('records')

    def _write_contributor_rows(self, ws, contributors):
        """Write contributor rows to worksheet"""
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

    def _write_raw_sheets(self, wb, cluster):
        """Write RAW 2G and RAW 4G sheets"""
        self._write_raw_2g_sheet(wb, cluster)
        self._write_raw_4g_sheet(wb, cluster)

    def _write_raw_2g_sheet(self, wb, cluster):
        """Write RAW 2G sheet"""
        gsm_data = self.kpi_data['gsm']
        gsm_cluster = gsm_data[gsm_data['CLUSTER'] == cluster].copy()

        if len(gsm_cluster) == 0:
            return

        # Prepare data
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

        gsm_raw = gsm_raw.drop_duplicates()
        gsm_raw['BEGIN_TIME'] = pd.to_datetime(
            gsm_raw['BEGIN_TIME']).dt.strftime('%-m/%-d/%Y')

        # Write to sheet
        if 'RAW 2G' in wb.sheetnames:
            wb.remove(wb['RAW 2G'])
        ws = wb.create_sheet('RAW 2G')

        self._write_raw_data(ws, gsm_raw)
        print(f"✓ Written RAW 2G: {len(gsm_raw)} records")

    def _write_raw_4g_sheet(self, wb, cluster):
        """Write RAW 4G sheet"""
        lte_data = self.kpi_data['lte']
        lte_cluster = lte_data[lte_data['CLUSTER'] == cluster].copy()

        if len(lte_cluster) == 0:
            return

        # Prepare data
        lte_raw = lte_cluster[[
            lte_cluster.columns[LTEColumns.BEGIN_TIME],
            'TOWER_ID',
            lte_cluster.columns[LTEColumns.CELL_NAME],
            'LTE_BAND',
            'SESSION_SSR', 'RACH_SR', 'HO_SR', 'ERAB_DROP',
            'DL_THP', 'UL_THP', 'UL_PLOSS', 'DL_PLOSS',
            'CQI', 'MIMO_RANK2', 'UL_RSSI', 'LATENCY',
            'LTC_NON_CAP', 'OVERLAP_RATE', 'SPECTRAL_EFF',
            'VOLTE_CSSR', 'VOLTE_DROP', 'SRVCC_SR'
        ]].copy()

        lte_raw.columns = ['BEGIN_TIME', 'TOWER_ID', 'CELL_NAME', 'BAND',
                           'SESSION_SSR', 'RACH_SR', 'HO_SR', 'ERAB_DROP',
                           'DL_THP', 'UL_THP', 'UL_PLOSS', 'DL_PLOSS',
                           'CQI', 'MIMO_RANK2', 'UL_RSSI', 'LATENCY',
                           'LTC_NON_CAP', 'OVERLAP_RATE', 'SPECTRAL_EFF',
                           'VOLTE_CSSR', 'VOLTE_DROP', 'SRVCC_SR']

        lte_raw = lte_raw.drop_duplicates()
        lte_raw['BEGIN_TIME'] = pd.to_datetime(
            lte_raw['BEGIN_TIME']).dt.strftime('%-m/%-d/%Y')

        # Write to sheet
        if 'RAW 4G' in wb.sheetnames:
            wb.remove(wb['RAW 4G'])
        ws = wb.create_sheet('RAW 4G')

        self._write_raw_data(ws, lte_raw)
        print(f"✓ Written RAW 4G: {len(lte_raw)} records")

    def _write_raw_data(self, ws, df):
        """Write raw data to worksheet"""
        # Write headers
        for c_idx, col in enumerate(df.columns, 1):
            self.formatter.format_header_small(
                ws.cell(row=1, column=c_idx), col)

        # Write data
        for r_idx, row in enumerate(df.itertuples(index=False), 2):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = value
                # Format numeric columns
                if c_idx > 3 and isinstance(value, (int, float)):
                    cell.number_format = '0.00'

    def _add_logos(self, ws):
        """Add logos to worksheet"""
        try:
            if self.xlsmart_logo:
                img_data = base64.b64decode(self.xlsmart_logo)
                img = XLImage(BytesIO(img_data))
                img.width, img.height = 200, 70
                ws.add_image(img, "A1")

            if self.zte_logo:
                img_data = base64.b64decode(self.zte_logo)
                img = XLImage(BytesIO(img_data))
                img.width, img.height = 100, 50
                ws.add_image(img, "R1")
        except Exception as e:
            print(f"⚠ Warning: Could not add logos: {e}")
