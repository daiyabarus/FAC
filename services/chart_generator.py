"""
Chart Generator Service
Generate matplotlib charts for all KPIs and save to BytesIO for Excel embedding
"""
import warnings
from models.column_enums import LTECol, GSMCol
from datetime import datetime
from io import BytesIO
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Use non-GUI backend
warnings.filterwarnings('ignore', category=RuntimeWarning)


class ChartGenerator:
    """Generate charts for KPIs"""

    def __init__(self):
        self.charts = []
        self.setup_matplotlib_style()

    def setup_matplotlib_style(self):
        """Setup modern matplotlib style"""
        plt.style.use('seaborn-v0_8-darkgrid')
        plt.rcParams['figure.figsize'] = (12, 6)
        plt.rcParams['font.size'] = 10
        plt.rcParams['axes.labelsize'] = 11
        plt.rcParams['axes.titlesize'] = 13
        plt.rcParams['xtick.labelsize'] = 9
        plt.rcParams['ytick.labelsize'] = 9
        plt.rcParams['legend.fontsize'] = 9

    def _safe_divide(self, numerator, denominator):
        """Safe division for aggregated data with zero protection"""
        # Handle division by zero properly
        num = np.array(numerator, dtype=float)
        den = np.array(denominator, dtype=float)

        # Create result array filled with NaN
        result = np.full_like(num, np.nan, dtype=float)

        # Only divide where denominator is not zero
        mask = den != 0
        result[mask] = num[mask] / den[mask]

        return result

    def generate_all_charts(self, lte_df: pd.DataFrame, gsm_df: pd.DataFrame,
                            cluster: str) -> list:
        """Generate all KPI charts for a specific cluster"""
        self.charts = []

        # Filter data by cluster
        lte_cluster = lte_df[lte_df['CLUSTER'] == cluster].copy()
        gsm_cluster = gsm_df[gsm_df['CLUSTER'] == cluster].copy()

        if len(lte_cluster) == 0 and len(gsm_cluster) == 0:
            return []

        # Generate 2G charts
        if len(gsm_cluster) > 0:
            try:
                self._generate_2g_charts(gsm_cluster, cluster)
            except Exception as e:
                print(f"Error generating 2G charts: {e}")

        # Generate 4G charts
        if len(lte_cluster) > 0:
            try:
                self._generate_4g_charts(lte_cluster, cluster)
            except Exception as e:
                print(f"Error generating 4G charts: {e}")

        return self.charts

    def _generate_2g_charts(self, data: pd.DataFrame, cluster: str):
        """Generate charts for 2G KPIs with zero-division protection"""
        # Ensure BEGIN_TIME is datetime
        data = data.copy()
        date_col = pd.to_datetime(
            data.iloc[:, GSMCol.GSM_BEGIN_TIME], errors='coerce')

        # Group by day
        daily_data = data.groupby(date_col.dt.date).agg({
            data.columns[GSMCol.GSM_CSSR_NUM]: 'sum',
            data.columns[GSMCol.GSM_CSSR_DEN]: 'sum',
            data.columns[GSMCol.GSM_SDCCH_SR_NUM]: 'sum',
            data.columns[GSMCol.GSM_SDCCH_SR_DEN]: 'sum',
            data.columns[GSMCol.GSM_DROP_NUM]: 'sum',
            data.columns[GSMCol.GSM_DROP_DEN]: 'sum',
        }).reset_index()

        daily_data.columns = ['Date', 'CSSR_NUM', 'CSSR_DEN', 'SDCCH_NUM',
                              'SDCCH_DEN', 'DROP_NUM', 'DROP_DEN']

        # Calculate KPIs with safe division
        daily_data['CSSR'] = self._safe_divide(
            daily_data['CSSR_NUM'].values, daily_data['CSSR_DEN'].values) * 100
        daily_data['SDCCH_SR'] = self._safe_divide(
            daily_data['SDCCH_NUM'].values, daily_data['SDCCH_DEN'].values) * 100
        daily_data['DROP_RATE'] = self._safe_divide(
            daily_data['DROP_NUM'].values, daily_data['DROP_DEN'].values) * 100

        # Remove NaN values
        daily_data = daily_data.dropna(
            subset=['CSSR', 'SDCCH_SR', 'DROP_RATE'])

        if len(daily_data) == 0:
            return  # Skip if no valid data

        # Chart 1: Call Setup Success Rate
        try:
            chart_data = self._create_line_chart(
                daily_data['Date'],
                daily_data['CSSR'],
                'Call Setup Success Rate - 2G RAN',
                f'{cluster} - Daily Trend',
                'Date',
                'CSSR (%)',
                baseline=98.5,
                target_color='green'
            )
            self.charts.append({
                'name': '2G_CSSR',
                'title': 'Call Setup Success Rate (2G)',
                'data': chart_data
            })
        except Exception as e:
            print(f"Warning: Could not generate 2G CSSR chart: {e}")

        # Chart 2: SDCCH Success Rate
        try:
            chart_data = self._create_line_chart(
                daily_data['Date'],
                daily_data['SDCCH_SR'],
                'SDCCH Success Rate - 2G RAN',
                f'{cluster} - Daily Trend',
                'Date',
                'SDCCH SR (%)',
                baseline=98.5,
                target_color='green'
            )
            self.charts.append({
                'name': '2G_SDCCH',
                'title': 'SDCCH Success Rate (2G)',
                'data': chart_data
            })
        except Exception as e:
            print(f"Warning: Could not generate 2G SDCCH chart: {e}")

        # Chart 3: Drop Rate
        try:
            chart_data = self._create_line_chart(
                daily_data['Date'],
                daily_data['DROP_RATE'],
                'Perceive Drop Rate - 2G RAN',
                f'{cluster} - Daily Trend',
                'Date',
                'Drop Rate (%)',
                baseline=2,
                target_color='red',
                invert_baseline=True
            )
            self.charts.append({
                'name': '2G_DROP',
                'title': 'Perceive Drop Rate (2G)',
                'data': chart_data
            })
        except Exception as e:
            print(f"Warning: Could not generate 2G DROP chart: {e}")

    def _generate_4g_charts(self, data: pd.DataFrame, cluster: str):
        """Generate charts for 4G KPIs with zero-division protection"""
        # Ensure BEGIN_TIME is datetime
        data = data.copy()
        date_col = pd.to_datetime(
            data.iloc[:, LTECol.LTE_BEGIN_TIME], errors='coerce')

        # Group by day
        daily_data = data.groupby(date_col.dt.date).agg({
            data.columns[LTECol.LTE_RRC_SSR_NUM]: 'sum',
            data.columns[LTECol.LTE_RRC_SSR_DEN]: 'sum',
            data.columns[LTECol.LTE_ERAB_SSR_NUM]: 'sum',
            data.columns[LTECol.LTE_ERAB_SSR_DEN]: 'sum',
            data.columns[LTECol.LTE_S1_SSR_NUM]: 'sum',
            data.columns[LTECol.LTE_S1_SSR_DEN]: 'sum',
            data.columns[LTECol.LTE_RACH_SETUP_NUM]: 'sum',
            data.columns[LTECol.LTE_RACH_SETUP_DEN]: 'sum',
            data.columns[LTECol.LTE_HO_SR_NUM]: 'sum',
            data.columns[LTECol.LTE_HO_SR_DEN]: 'sum',
            data.columns[LTECol.LTE_ERAB_DROP_NUM]: 'sum',
            data.columns[LTECol.LTE_ERAB_DROP_DEN]: 'sum',
            data.columns[LTECol.LTE_DL_THP_NUM]: 'sum',
            data.columns[LTECol.LTE_DL_THP_DEN]: 'sum',
            data.columns[LTECol.LTE_UL_THP_NUM]: 'sum',
            data.columns[LTECol.LTE_UL_THP_DEN]: 'sum',
            data.columns[LTECol.LTE_CQI_NUM]: 'sum',
            data.columns[LTECol.LTE_CQI_DEN]: 'sum',
            data.columns[LTECol.LTE_VOLTE_CSSR_NUM]: 'sum',
            data.columns[LTECol.LTE_VOLTE_CSSR_DEN]: 'sum',
        }).reset_index()

        daily_data.columns = ['Date', 'RRC_NUM', 'RRC_DEN', 'ERAB_NUM', 'ERAB_DEN',
                              'S1_NUM', 'S1_DEN', 'RACH_NUM', 'RACH_DEN',
                              'HO_NUM', 'HO_DEN', 'DROP_NUM', 'DROP_DEN',
                              'DL_NUM', 'DL_DEN', 'UL_NUM', 'UL_DEN',
                              'CQI_NUM', 'CQI_DEN', 'VOLTE_NUM', 'VOLTE_DEN']

        # Calculate KPIs with safe division - USE .values to get numpy arrays
        rrc_sr = self._safe_divide(
            daily_data['RRC_NUM'].values, daily_data['RRC_DEN'].values)
        erab_sr = self._safe_divide(
            daily_data['ERAB_NUM'].values, daily_data['ERAB_DEN'].values)
        s1_sr = self._safe_divide(
            daily_data['S1_NUM'].values, daily_data['S1_DEN'].values)

        daily_data['SESSION_SR'] = 100 * rrc_sr * erab_sr * s1_sr
        daily_data['RACH_SR'] = 100 * self._safe_divide(
            daily_data['RACH_NUM'].values, daily_data['RACH_DEN'].values)
        daily_data['HO_SR'] = 100 * \
            self._safe_divide(
                daily_data['HO_NUM'].values, daily_data['HO_DEN'].values)
        daily_data['ERAB_DROP'] = 100 * self._safe_divide(
            daily_data['DROP_NUM'].values, daily_data['DROP_DEN'].values)
        daily_data['DL_THP'] = self._safe_divide(
            daily_data['DL_NUM'].values, daily_data['DL_DEN'].values)
        daily_data['UL_THP'] = self._safe_divide(
            daily_data['UL_NUM'].values, daily_data['UL_DEN'].values)
        daily_data['CQI'] = self._safe_divide(
            daily_data['CQI_NUM'].values, daily_data['CQI_DEN'].values)
        daily_data['VOLTE_CSSR'] = 100 * self._safe_divide(
            daily_data['VOLTE_NUM'].values, daily_data['VOLTE_DEN'].values)

        # Generate charts for each KPI
        kpi_charts = [
            ('SESSION_SR', 'Session Setup Success Rate', 99, 'green', False),
            ('RACH_SR', 'RACH Success Rate', 85, 'green', False),
            ('HO_SR', 'Handover Success Rate', 97, 'green', False),
            ('ERAB_DROP', 'E-RAB Drop Rate', 2, 'red', True),
            ('DL_THP', 'Downlink User Throughput (Mbps)', 3, 'green', False),
            ('UL_THP', 'Uplink User Throughput (Mbps)', 1, 'green', False),
            ('CQI', 'CQI', 7, 'green', False),
            ('VOLTE_CSSR', 'VoLTE Call Setup Success Rate', 97, 'green', False),
        ]

        for kpi_col, kpi_title, baseline, color, invert in kpi_charts:
            if kpi_col in daily_data.columns:
                # Remove NaN and inf values for this specific KPI
                chart_data_subset = daily_data[['Date', kpi_col]].copy()
                chart_data_subset = chart_data_subset[
                    np.isfinite(chart_data_subset[kpi_col])
                ]

                if len(chart_data_subset) > 0:
                    try:
                        chart_data = self._create_line_chart(
                            chart_data_subset['Date'],
                            chart_data_subset[kpi_col],
                            f'{kpi_title} - 4G RAN',
                            f'{cluster} - Daily Trend',
                            'Date',
                            kpi_title,
                            baseline=baseline,
                            target_color=color,
                            invert_baseline=invert
                        )
                        self.charts.append({
                            'name': f'4G_{kpi_col}',
                            'title': kpi_title,
                            'data': chart_data
                        })
                    except Exception as e:
                        print(
                            f"Warning: Could not generate 4G {kpi_title} chart: {e}")

    def _create_line_chart(self, x_data, y_data, title, subtitle, xlabel, ylabel,
                           baseline=None, target_color='green', invert_baseline=False):
        """Create a line chart with baseline and return as BytesIO"""
        fig, ax = plt.subplots(figsize=(12, 6))

        # Plot main data
        ax.plot(x_data, y_data, marker='o', linewidth=2, markersize=5,
                label='Actual', color='#0078d4')

        # Add baseline if provided
        if baseline is not None:
            ax.axhline(y=baseline, color=target_color, linestyle='--',
                       linewidth=2, label=f'Baseline: {baseline}', alpha=0.7)

        # Formatting
        ax.set_title(f'{title}\n{subtitle}', fontsize=14,
                     fontweight='bold', pad=20)
        ax.set_xlabel(xlabel, fontsize=11)
        ax.set_ylabel(ylabel, fontsize=11)
        ax.grid(True, alpha=0.3)
        ax.legend(loc='best')

        # Rotate x-axis labels
        plt.xticks(rotation=45, ha='right')

        # Tight layout
        plt.tight_layout()

        # Save to BytesIO
        buf = BytesIO()
        plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        buf.seek(0)
        plt.close(fig)

        return buf
