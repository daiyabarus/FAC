"""
KPI Calculator Service - COMPLETE WITH ZERO-DIVISION PROTECTION
Calculate all KPIs based on formulas and thresholds
"""
import pandas as pd
import numpy as np
import json
from pathlib import Path
from models.column_enums import LTECol, GSMCol


class KPICalculator:
    """Calculate KPIs and determine PASS/FAIL status"""

    def __init__(self):
        self.config = self._load_kpi_config()
        self.band_config = self._load_band_mapping()
        self.results = {
            '2g_ran': {},
            '4g_ran': {},
            'contributors': []
        }

    def _load_kpi_config(self) -> dict:
        """Load KPI configuration from JSON"""
        config_path = Path('config/kpi_config.json')
        if config_path.exists():
            with open(config_path, 'r') as f:
                return json.load(f)
        return {'2g_ran_kpis': [], '4g_ran_kpis': []}

    def _load_band_mapping(self) -> dict:
        """Load band mapping configuration"""
        config_path = Path('config/band_mapping.json')
        if config_path.exists():
            with open(config_path, 'r') as f:
                return json.load(f)
        return {}

    def _safe_divide(self, numerator, denominator):
        """Safe division that handles zero denominator - pandas 2.x compatible"""
        # Convert to numeric and use numpy for safe division
        num = pd.to_numeric(numerator, errors='coerce')
        den = pd.to_numeric(denominator, errors='coerce')

        # Use numpy where to avoid division by zero
        with np.errstate(divide='ignore', invalid='ignore'):
            result = np.where(den != 0, num / den, np.nan)

        return pd.Series(result, index=numerator.index if hasattr(numerator, 'index') else range(len(result)))

    def calculate_all_kpis(self, lte_df: pd.DataFrame, gsm_df: pd.DataFrame) -> dict:
        """Calculate all KPIs for all months"""
        # Get unique months (sorted descending - newest first)
        months = self._get_sorted_months(lte_df, gsm_df)

        if len(months) == 0:
            raise Exception("No months found in data")

        # Calculate 2G RAN KPIs
        for month in months[:3]:  # Process up to 3 months
            self._calculate_2g_kpis_for_month(gsm_df, month)

        # Calculate 4G RAN KPIs
        for month in months[:3]:
            self._calculate_4g_kpis_for_month(lte_df, month)

        return self.results

    def _get_sorted_months(self, lte_df: pd.DataFrame, gsm_df: pd.DataFrame) -> list:
        """Get sorted list of unique months from data - NEWEST FIRST"""
        months = set()

        if 'YEAR_MONTH' in lte_df.columns:
            month_data = lte_df['YEAR_MONTH'].dropna().unique()
            if len(month_data) > 0:
                # Convert string back to Period for proper sorting
                if isinstance(month_data[0], str):
                    months.update(pd.PeriodIndex(month_data, freq='M'))
                else:
                    months.update(month_data)

        if 'YEAR_MONTH' in gsm_df.columns:
            month_data = gsm_df['YEAR_MONTH'].dropna().unique()
            if len(month_data) > 0:
                if isinstance(month_data[0], str):
                    months.update(pd.PeriodIndex(month_data, freq='M'))
                else:
                    months.update(month_data)

        # Sort DESCENDING (newest first) and take first 3 months
        return sorted(list(months), reverse=True)[:3]

    def _calculate_2g_kpis_for_month(self, gsm_df: pd.DataFrame, month):
        """Calculate 2G RAN KPIs for specific month"""
        # Convert string month to Period if needed
        if isinstance(month, str):
            month = pd.Period(month, freq='M')

        # Filter by month - handle both Period and string
        if isinstance(gsm_df['YEAR_MONTH'].iloc[0], str):
            month_data = gsm_df[gsm_df['YEAR_MONTH'] == str(month)]
        else:
            month_data = gsm_df[gsm_df['YEAR_MONTH'] == month]

        if len(month_data) == 0:
            return

        month_key = f"month_{len([k for k in self.results['2g_ran'].keys() if k.startswith('month_')]) + 1}"
        self.results['2g_ran'][month_key] = {}

        for kpi_config in self.config['2g_ran_kpis']:
            kpi_name = kpi_config['name']
            result = self._calculate_2g_kpi(month_data, kpi_config)
            self.results['2g_ran'][month_key][kpi_name] = result

    def _calculate_2g_kpi(self, data: pd.DataFrame, kpi_config: dict) -> dict:
        """Calculate individual 2G KPI with zero-division protection"""
        kpi_name = kpi_config['name']
        data = data.copy()

        # Calculate KPI value for each cell
        if kpi_name == "Call Setup Success Rate":
            data['KPI_VALUE'] = 100 * self._safe_divide(
                data.iloc[:, GSMCol.GSM_CSSR_NUM],
                data.iloc[:, GSMCol.GSM_CSSR_DEN]
            )
            baseline = 98.5
            target_pct = 95
            comparison = '>='

        elif kpi_name == "SDCCH Success rate":
            data['KPI_VALUE'] = 100 * self._safe_divide(
                data.iloc[:, GSMCol.GSM_SDCCH_SR_NUM],
                data.iloc[:, GSMCol.GSM_SDCCH_SR_DEN]
            )
            baseline = 98.5
            target_pct = 95
            comparison = '>='

        elif kpi_name == "Perceive Drop Rate":
            data['KPI_VALUE'] = 100 * self._safe_divide(
                data.iloc[:, GSMCol.GSM_DROP_NUM],
                data.iloc[:, GSMCol.GSM_DROP_DEN]
            )
            baseline = 2
            target_pct = 95
            comparison = '<'

        # Remove NaN and inf values
        data = data[data['KPI_VALUE'].notna() & ~data['KPI_VALUE'].isin([
            float('inf'), float('-inf')])]

        # Count cells meeting baseline
        if comparison == '>=':
            cells_pass = (data['KPI_VALUE'] >= baseline).sum()
        else:  # '<'
            cells_pass = (data['KPI_VALUE'] < baseline).sum()

        total_cells = len(data)

        if total_cells > 0:
            pass_percentage = (cells_pass / total_cells) * 100
        else:
            pass_percentage = 0

        # Determine PASS/FAIL
        status = 'PASS' if pass_percentage >= target_pct else 'FAIL'

        # Find contributors (cells that don't meet baseline)
        if status == 'FAIL' and len(data) > 0:
            if comparison == '>=':
                failed_cells = data[data['KPI_VALUE'] < baseline]
            else:
                failed_cells = data[data['KPI_VALUE'] >= baseline]

            for _, row in failed_cells.iterrows():
                self.results['contributors'].append({
                    'Month': row.get('MONTH_NAME', 'Unknown'),
                    'Clause Type': kpi_config.get('clause_type', ''),
                    'KPI Domain': kpi_config.get('kpi_domain', ''),
                    'Name': kpi_name,
                    'Reference': kpi_config.get('reference', ''),
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[GSMCol.GSM_BTS_NAME],
                    'Value': row['KPI_VALUE'],
                    'Baseline': baseline,
                    'Status': 'FAIL'
                })

        return {
            'value': pass_percentage,
            'status': status,
            'cells_pass': cells_pass,
            'total_cells': total_cells,
            'baseline': baseline,
            'target': target_pct
        }

    def _calculate_4g_kpis_for_month(self, lte_df: pd.DataFrame, month):
        """Calculate 4G RAN KPIs for specific month"""
        # Convert string month to Period if needed
        if isinstance(month, str):
            month = pd.Period(month, freq='M')

        # Filter by month - handle both Period and string
        if isinstance(lte_df['YEAR_MONTH'].iloc[0], str):
            month_data = lte_df[lte_df['YEAR_MONTH'] == str(month)]
        else:
            month_data = lte_df[lte_df['YEAR_MONTH'] == month]

        if len(month_data) == 0:
            return

        month_key = f"month_{len([k for k in self.results['4g_ran'].keys() if k.startswith('month_')]) + 1}"
        self.results['4g_ran'][month_key] = {}

        for kpi_config in self.config['4g_ran_kpis']:
            kpi_name = kpi_config['name']

            # Handle KPIs with multiple conditions
            if 'conditions' in kpi_config:
                results = self._calculate_4g_kpi_multi_condition(
                    month_data, kpi_config)
                self.results['4g_ran'][month_key][kpi_name] = results
            elif kpi_name == "Spectral Efficiency":
                results = self._calculate_spectral_efficiency(
                    month_data, kpi_config)
                self.results['4g_ran'][month_key][kpi_name] = results
            else:
                result = self._calculate_4g_kpi_single(month_data, kpi_config)
                self.results['4g_ran'][month_key][kpi_name] = result

    def _calculate_4g_kpi_single(self, data: pd.DataFrame, kpi_config: dict) -> dict:
        """Calculate 4G KPI with single condition - zero-division safe"""
        kpi_name = kpi_config['name']
        data = data.copy()

        # Calculate KPI based on formula with safe division
        if kpi_name == "Session Setup Success Rate":
            rrc = self._safe_divide(
                data.iloc[:, LTECol.LTE_RRC_SSR_NUM],
                data.iloc[:, LTECol.LTE_RRC_SSR_DEN]
            )
            erab = self._safe_divide(
                data.iloc[:, LTECol.LTE_ERAB_SSR_NUM],
                data.iloc[:, LTECol.LTE_ERAB_SSR_DEN]
            )
            s1 = self._safe_divide(
                data.iloc[:, LTECol.LTE_S1_SSR_NUM],
                data.iloc[:, LTECol.LTE_S1_SSR_DEN]
            )
            data['KPI_VALUE'] = 100 * rrc * erab * s1
            baseline = 99
            target = 97
            comparison = '>='

        elif kpi_name == "Handover Success Rate Inter and Intra-Frequency":
            data['KPI_VALUE'] = 100 * self._safe_divide(
                data.iloc[:, LTECol.LTE_HO_SR_NUM],
                data.iloc[:, LTECol.LTE_HO_SR_DEN]
            )
            baseline = 97
            target = 95
            comparison = '>='

        elif kpi_name == "E-RAB Drop Rate":
            data['KPI_VALUE'] = 100 * self._safe_divide(
                data.iloc[:, LTECol.LTE_ERAB_DROP_NUM],
                data.iloc[:, LTECol.LTE_ERAB_DROP_DEN]
            )
            baseline = 2
            target = 95
            comparison = '<'

        elif kpi_name == "UL Packet Loss (PDCP )":
            data['KPI_VALUE'] = data.iloc[:, LTECol.LTE_UL_PLOSS]
            # Exclude zero values
            data = data[data['KPI_VALUE'] != 0]
            baseline = 0.85
            target = 97
            comparison = '<'

        elif kpi_name == "DL Packet Loss (PDCP )":
            data['KPI_VALUE'] = data.iloc[:, LTECol.LTE_DL_PLOSS]
            # Exclude zero values
            data = data[data['KPI_VALUE'] != 0]
            baseline = 0.10
            target = 97
            comparison = '<'

        elif kpi_name == "CQI":
            data['KPI_VALUE'] = self._safe_divide(
                data.iloc[:, LTECol.LTE_CQI_NUM],
                data.iloc[:, LTECol.LTE_CQI_DEN]
            )
            baseline = 7
            target = 95
            comparison = '>='

        elif kpi_name == "UL RSSI":
            data['KPI_VALUE'] = self._safe_divide(
                data.iloc[:, LTECol.LTE_RSSI_PUSCH_NUM],
                data.iloc[:, LTECol.LTE_RSSI_PUSCH_DEN]
            )
            baseline = -105
            target = 97
            comparison = '<'

        elif kpi_name == "LTC Non Capacity":
            data['KPI_VALUE'] = data.iloc[:, LTECol.LTE_LTC_NON_CAP]
            baseline = 3
            target = 5
            comparison = '<'

        elif kpi_name == "Coverage Overlaping ratio ":
            data['KPI_VALUE'] = data.iloc[:, LTECol.LTE_OVERLAP_RATE]
            baseline = 35
            target = 80
            comparison = '<'

        elif kpi_name == "Voice Call Success Rate (VoLTE)":
            data['KPI_VALUE'] = 100 * self._safe_divide(
                data.iloc[:, LTECol.LTE_VOLTE_CSSR_NUM],
                data.iloc[:, LTECol.LTE_VOLTE_CSSR_DEN]
            )
            baseline = 97
            target = 95
            comparison = '>'

        elif kpi_name == "Voice Call Drop Rate (VoLTE)":
            data['KPI_VALUE'] = 100 * self._safe_divide(
                data.iloc[:, LTECol.LTE_VOLTE_DROP_NUM],
                data.iloc[:, LTECol.LTE_VOLTE_DROP_DEN]
            )
            baseline = 2
            target = 95
            comparison = '<'

        elif kpi_name == "SRVCC Success Rate":
            data['KPI_VALUE'] = 100 * self._safe_divide(
                data.iloc[:, LTECol.LTE_SRVCC_SR_NUM],
                data.iloc[:, LTECol.LTE_SRVCC_SR_DEN]
            )
            baseline = 97
            target = 95
            comparison = '>='

        else:
            return {'value': 0, 'status': 'N/A', 'cells_pass': 0, 'total_cells': 0}

        # Clean data - remove NaN and inf values
        data = data[data['KPI_VALUE'].notna() & ~data['KPI_VALUE'].isin([
            float('inf'), float('-inf')])]

        # Count cells meeting baseline
        if comparison == '>=':
            cells_pass = (data['KPI_VALUE'] >= baseline).sum()
        elif comparison == '>':
            cells_pass = (data['KPI_VALUE'] > baseline).sum()
        else:  # '<'
            cells_pass = (data['KPI_VALUE'] < baseline).sum()

        total_cells = len(data)

        if total_cells > 0:
            pass_percentage = (cells_pass / total_cells) * 100
        else:
            # Return early if no valid data
            return {
                'value': 0,
                'status': 'N/A',
                'cells_pass': 0,
                'total_cells': 0,
                'baseline': baseline,
                'target': target
            }

        # Determine PASS/FAIL
        if comparison in ['>=', '>']:
            status = 'PASS' if pass_percentage >= target else 'FAIL'
        else:
            status = 'PASS' if pass_percentage >= target else 'FAIL'

        # Track contributors if FAIL
        if status == 'FAIL' and len(data) > 0:
            if comparison in ['>=', '>']:
                failed_cells = data[data['KPI_VALUE'] < baseline]
            else:
                failed_cells = data[data['KPI_VALUE'] >= baseline]

            for _, row in failed_cells.iterrows():
                self.results['contributors'].append({
                    'Month': row.get('MONTH_NAME', 'Unknown'),
                    'Clause Type': kpi_config.get('clause_type', ''),
                    'KPI Domain': kpi_config.get('kpi_domain', ''),
                    'Name': kpi_name,
                    'Reference': kpi_config.get('reference', ''),
                    'TOWER_ID': row.get('TOWER_ID', ''),
                    'CELL_NAME': row.iloc[LTECol.LTE_CELL_NAME],
                    'Value': row['KPI_VALUE'],
                    'Baseline': baseline,
                    'Status': 'FAIL'
                })

        return {
            'value': pass_percentage,
            'status': status,
            'cells_pass': cells_pass,
            'total_cells': total_cells,
            'baseline': baseline,
            'target': target
        }

    def _calculate_4g_kpi_multi_condition(self, data: pd.DataFrame, kpi_config: dict) -> list:
        """Calculate 4G KPI with multiple conditions - zero-division safe"""
        kpi_name = kpi_config['name']
        conditions = kpi_config.get('conditions', [])
        results = []

        for condition in conditions:
            data_copy = data.copy()

            if kpi_name == "RACH Success Rate":
                data_copy['KPI_VALUE'] = 100 * self._safe_divide(
                    data_copy.iloc[:, LTECol.LTE_RACH_SETUP_NUM],
                    data_copy.iloc[:, LTECol.LTE_RACH_SETUP_DEN]
                )

                # Parse baseline from condition
                baseline_str = condition['baseline']
                if baseline_str == '>=85':
                    baseline = 85
                    comparison = '>='
                    target = float(condition['target'].replace(
                        '>=', '').replace('%', ''))
                elif baseline_str == '<55':
                    baseline = 55
                    comparison = '<'
                    target = float(condition['target'].replace(
                        '<', '').replace('%', ''))

            elif kpi_name == "Downlink User Throughput":
                data_copy['KPI_VALUE'] = self._safe_divide(
                    data_copy.iloc[:, LTECol.LTE_DL_THP_NUM],
                    data_copy.iloc[:, LTECol.LTE_DL_THP_DEN]
                )

                baseline_str = condition['baseline']
                if baseline_str == '>=3':
                    baseline = 3
                    comparison = '>='
                    target = float(condition['target'].replace(
                        '>=', '').replace('%', ''))
                elif baseline_str == '<1':
                    baseline = 1
                    comparison = '<'
                    target = float(condition['target'].replace(
                        '<', '').replace('%', ''))

            elif kpi_name == "Uplink User Throughput":
                data_copy['KPI_VALUE'] = self._safe_divide(
                    data_copy.iloc[:, LTECol.LTE_UL_THP_NUM],
                    data_copy.iloc[:, LTECol.LTE_UL_THP_DEN]
                )

                baseline_str = condition['baseline']
                if baseline_str == '>=1':
                    baseline = 1
                    comparison = '>='
                    target = float(condition['target'].replace(
                        '>=', '').replace('%', ''))
                elif baseline_str == '<0.256':
                    baseline = 0.256
                    comparison = '<'
                    target = float(condition['target'].replace(
                        '<', '').replace('%', ''))

            elif kpi_name == "MIMO Transmission Rank2 Rate":
                data_copy['KPI_VALUE'] = 100 * self._safe_divide(
                    data_copy.iloc[:, LTECol.LTE_RANK_GT2_NUM],
                    data_copy.iloc[:, LTECol.LTE_RANK_GT2_DEN]
                )

                baseline_str = condition['baseline']
                if baseline_str == '>=35':
                    baseline = 35
                    comparison = '>='
                    target = float(condition['target'].replace(
                        '>=', '').replace('%', ''))
                elif baseline_str == '<20':
                    baseline = 20
                    comparison = '<'
                    target = float(condition['target'].replace(
                        '<', '').replace('%', ''))

            elif kpi_name == "Packet Latency":
                data_copy['KPI_VALUE'] = self._safe_divide(
                    data_copy.iloc[:, LTECol.LTE_RAN_LAT_NUM],
                    data_copy.iloc[:, LTECol.LTE_RAN_LAT_DEN]
                )

                baseline_str = condition['baseline']
                if baseline_str == '<30':
                    baseline = 30
                    comparison = '<'
                    target = float(condition['target'].replace(
                        '>=', '').replace('%', ''))
                elif '>30 and <40' in baseline_str:
                    # Special case: between 30 and 40
                    data_copy = data_copy[
                        (data_copy['KPI_VALUE'] > 30) & (
                            data_copy['KPI_VALUE'] < 40)
                    ]
                    baseline = 40
                    comparison = 'between'
                    target = float(condition['target'].replace(
                        '<', '').replace('%', ''))

            # Clean data - remove NaN and inf
            data_copy = data_copy[data_copy['KPI_VALUE'].notna() & ~data_copy['KPI_VALUE'].isin([
                float('inf'), float('-inf')])]

            # Count cells meeting baseline
            if comparison == '>=':
                cells_pass = (data_copy['KPI_VALUE'] >= baseline).sum()
            elif comparison == '<':
                cells_pass = (data_copy['KPI_VALUE'] < baseline).sum()
            elif comparison == 'between':
                cells_pass = len(data_copy)  # Already filtered

            total_cells = len(data_copy)

            if total_cells > 0:
                pass_percentage = (cells_pass / total_cells) * 100
            else:
                pass_percentage = 0

            # Determine PASS/FAIL based on comparison type
            if comparison in ['>=', 'between']:
                status = 'PASS' if pass_percentage >= target else 'FAIL'
            else:  # '<'
                # For '<' baseline with '<' target, means we want low percentage
                if '<' in condition['target']:
                    status = 'PASS' if pass_percentage < target else 'FAIL'
                else:
                    status = 'PASS' if pass_percentage >= target else 'FAIL'

            results.append({
                'value': pass_percentage,
                'status': status,
                'cells_pass': cells_pass,
                'total_cells': total_cells,
                'baseline': baseline_str,
                'target': target,
                'condition': baseline_str
            })

        return results

    def _calculate_spectral_efficiency(self, data: pd.DataFrame, kpi_config: dict) -> list:
        """Calculate Spectral Efficiency with dynamic baselines - zero-division safe"""
        results = []
        se_baselines = self.band_config.get(
            'spectral_efficiency_baselines', {})

        # DEBUG: Print unique TX and BAND values
        print(f"DEBUG SE: Available TX values: {data['TX'].unique()}")
        print(f"DEBUG SE: Available BAND values: {data['BAND'].unique()}")

        # Define configurations to check
        configs = [
            ('2T2R', '850', 1.1, 90),
            ('2T2R', '900', 1.1, 90),
            ('2T2R', '2100', 1.3, 90),
            ('2T2R', '1800', 1.25, 90),
            ('4T4R', '1800', 1.5, 90),
            ('4T4R', '2100', 1.7, 90),
            ('4T4R', '2300', 1.7, 90),
            ('8T8R', '1800', 1.5, 90),
            ('8T8R', '2100', 1.7, 90),
            ('8T8R', '2300', 1.7, 90),
            ('32T32R', '2300', 2.1, 90),
        ]

        for tx, band, baseline, target in configs:
            # Filter data by TX and BAND - handle string matching
            filtered_data = data[
                (data['TX'].astype(str).str.strip() == tx) &
                (data['BAND'].astype(str).str.strip() == band)
            ].copy()

            print(f"DEBUG SE: {tx} {band} - Found {len(filtered_data)} cells")

            if len(filtered_data) == 0:
                continue

            # Calculate SE with safe division
            filtered_data['KPI_VALUE'] = self._safe_divide(
                filtered_data.iloc[:, LTECol.LTE_DL_SE_NUM],
                filtered_data.iloc[:, LTECol.LTE_DL_SE_DEN]
            )

            # Clean data - remove NaN and inf
            filtered_data = filtered_data[
                filtered_data['KPI_VALUE'].notna() &
                ~filtered_data['KPI_VALUE'].isin([float('inf'), float('-inf')])
            ]

            # Count cells meeting baseline
            cells_pass = (filtered_data['KPI_VALUE'] >= baseline).sum()
            total_cells = len(filtered_data)

            if total_cells > 0:
                pass_percentage = (cells_pass / total_cells) * 100
            else:
                pass_percentage = 0

            status = 'PASS' if pass_percentage > target else 'FAIL'

            # Track contributors if FAIL
            if status == 'FAIL' and len(filtered_data) > 0:
                failed_cells = filtered_data[filtered_data['KPI_VALUE'] < baseline]
                for _, row in failed_cells.iterrows():
                    self.results['contributors'].append({
                        'Month': row.get('MONTH_NAME', 'Unknown'),
                        'Clause Type': kpi_config.get('clause_type', ''),
                        'KPI Domain': kpi_config.get('kpi_domain', ''),
                        'Name': f"Spectral Efficiency ({tx} {band})",
                        'Reference': kpi_config.get('reference', ''),
                        'TOWER_ID': row.get('TOWER_ID', ''),
                        'CELL_NAME': row.iloc[LTECol.LTE_CELL_NAME],
                        'Value': row['KPI_VALUE'],
                        'Baseline': baseline,
                        'Status': 'FAIL',
                        'TX': tx,
                        'BAND': band
                    })

            results.append({
                'value': pass_percentage,
                'status': status,
                'cells_pass': cells_pass,
                'total_cells': total_cells,
                'baseline': baseline,
                'target': target,
                'tx': tx,
                'band': band,
                'config': f"{tx}_{band}"
            })

        return results

    def get_overall_status_by_month(self) -> dict:
        """Determine overall PASS/FAIL for each month"""
        overall = {}

        for month_key in ['month_1', 'month_2', 'month_3']:
            has_fail = False

            # Check 2G RAN
            if month_key in self.results['2g_ran']:
                for kpi_name, kpi_result in self.results['2g_ran'][month_key].items():
                    if isinstance(kpi_result, dict) and kpi_result.get('status') == 'FAIL':
                        has_fail = True
                        break

            # Check 4G RAN
            if month_key in self.results['4g_ran']:
                for kpi_name, kpi_result in self.results['4g_ran'][month_key].items():
                    if isinstance(kpi_result, list):
                        for result in kpi_result:
                            if result.get('status') == 'FAIL':
                                has_fail = True
                                break
                    elif isinstance(kpi_result, dict) and kpi_result.get('status') == 'FAIL':
                        has_fail = True
                        break

            overall[month_key] = 'FAIL' if has_fail else 'PASS'

        return overall
