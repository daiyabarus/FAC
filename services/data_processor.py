"""
Data Processor Service
Process and transform raw data, mapping clusters, converting bands, grouping
"""
import pandas as pd
import json
import re
from pathlib import Path
from models.column_enums import LTECol, GSMCol, ClusterCol
from datetime import datetime
import numpy as np


class DataProcessor:
    """Process and transform data for KPI calculations"""

    def __init__(self):
        self.lte_data = None
        self.gsm_data = None
        self.cluster_data = None
        self.band_mapping = self._load_band_mapping()

    def _load_band_mapping(self) -> dict:
        """Load band mapping from JSON config"""
        config_path = Path('config/band_mapping.json')
        if config_path.exists():
            with open(config_path, 'r') as f:
                return json.load(f)
        return {
            "freq_band_mapping": {
                "5": "850", "3": "1800", "1": "2100", "40": "2300"
            }
        }

    def process(self, lte_df: pd.DataFrame, gsm_df: pd.DataFrame,
                cluster_df: pd.DataFrame) -> tuple:
        """Main processing pipeline"""
        self.lte_data = lte_df.copy()
        self.gsm_data = gsm_df.copy()
        self.cluster_data = cluster_df.copy()

        # Process LTE data
        self._add_tower_id_to_lte()
        self._add_band_to_lte()
        self._map_cluster_to_lte()
        self._add_month_columns_lte()

        # Process GSM data
        self._map_cluster_to_gsm()
        self._add_month_columns_gsm()

        return self.lte_data, self.gsm_data

    def _add_tower_id_to_lte(self):
        """Extract TOWER_ID from ME_NAME using regex"""
        pattern = r"#([^#]+)#"

        def extract_tower_id(me_name):
            if pd.isna(me_name):
                return None
            match = re.search(pattern, str(me_name))
            return match.group(1) if match else None

        self.lte_data['TOWER_ID'] = self.lte_data.iloc[:, LTECol.LTE_ME_NAME].apply(
            extract_tower_id
        )

    def _add_band_to_lte(self):
        """Convert FREQ_BAND to BAND name"""
        freq_mapping = self.band_mapping.get('freq_band_mapping', {})

        def convert_band(freq_band):
            if pd.isna(freq_band):
                return None
            freq_str = str(int(freq_band)) if isinstance(
                freq_band, (int, float)) else str(freq_band)
            return freq_mapping.get(freq_str, freq_str)

        self.lte_data['BAND'] = self.lte_data.iloc[:, LTECol.LTE_FREQ_BAND].apply(
            convert_band
        )

    def _map_cluster_to_lte(self):
        """Map CLUSTER and TX info from cluster_data to LTE"""
        # Create mapping dictionary from cluster data
        tower_to_cluster = {}
        for _, row in self.cluster_data.iterrows():
            tower_id = row.iloc[ClusterCol.CLUSTER_TOWERID]
            tower_to_cluster[tower_id] = {
                'CLUSTER': row.iloc[ClusterCol.CLUSTER_CLUSTER],
                'TX': row.iloc[ClusterCol.CLUSTER_TX]
            }

        # Map to LTE data
        self.lte_data['CLUSTER'] = self.lte_data['TOWER_ID'].map(
            lambda x: tower_to_cluster.get(x, {}).get('CLUSTER', None)
        )
        self.lte_data['TX'] = self.lte_data['TOWER_ID'].map(
            lambda x: tower_to_cluster.get(x, {}).get('TX', None)
        )

    def _map_cluster_to_gsm(self):
        """Map CLUSTER and TOWER_ID from cluster_data to GSM using SITENAME"""
        # Create mapping from SITENAME
        sitename_to_cluster = {}
        for _, row in self.cluster_data.iterrows():
            sitename = row.iloc[ClusterCol.CLUSTER_SITENAME]
            sitename_to_cluster[sitename] = {
                'CLUSTER': row.iloc[ClusterCol.CLUSTER_CLUSTER],
                'TOWER_ID': row.iloc[ClusterCol.CLUSTER_TOWERID]
            }

        # Map to GSM data
        self.gsm_data['CLUSTER'] = self.gsm_data.iloc[:, GSMCol.GSM_SITE_NAME].map(
            lambda x: sitename_to_cluster.get(x, {}).get('CLUSTER', None)
        )
        self.gsm_data['TOWER_ID'] = self.gsm_data.iloc[:, GSMCol.GSM_SITE_NAME].map(
            lambda x: sitename_to_cluster.get(x, {}).get('TOWER_ID', None)
        )

    def _add_month_columns_lte(self):
        """Add formatted month and date columns to LTE data"""
        # Ensure BEGIN_TIME is datetime
        begin_time_col = pd.to_datetime(
            self.lte_data.iloc[:, LTECol.LTE_BEGIN_TIME], errors='coerce')

        # Extract month name (e.g., "September")
        self.lte_data['MONTH_NAME'] = begin_time_col.dt.strftime('%B')

        # Extract year-month for grouping
        self.lte_data['YEAR_MONTH'] = begin_time_col.dt.to_period('M')

        # Extract day for daily grouping
        self.lte_data['DAY'] = begin_time_col.dt.day

    def _add_month_columns_gsm(self):
        """Add formatted month and date columns to GSM data"""
        # Ensure BEGIN_TIME is datetime
        begin_time_col = pd.to_datetime(
            self.gsm_data.iloc[:, GSMCol.GSM_BEGIN_TIME], errors='coerce')

        # Extract month name
        self.gsm_data['MONTH_NAME'] = begin_time_col.dt.strftime('%B')

        # Extract year-month for grouping
        self.gsm_data['YEAR_MONTH'] = begin_time_col.dt.to_period('M')

        # Extract day for daily grouping
        self.gsm_data['DAY'] = begin_time_col.dt.day

    def group_lte_by_cluster_month(self) -> pd.DataFrame:
        """
        Group LTE data by CLUSTER, YEAR_MONTH, ME_NAME, CELL_NAME, BAND, TX
        Aggregate numerators and denominators with SUM
        """
        group_cols = ['CLUSTER', 'YEAR_MONTH', 'TOWER_ID',
                      self.lte_data.columns[LTECol.LTE_ME_NAME],
                      self.lte_data.columns[LTECol.LTE_CELL_NAME],
                      'BAND', 'TX']

        # Define aggregation functions for numeric columns
        agg_dict = {}

        # Sum all numeric KPI columns (convert to numeric first)
        for col_idx in range(LTECol.LTE_RRC_SSR_NUM, min(LTECol.LTE_LTC_NON_CAP + 1, len(self.lte_data.columns))):
            col_name = self.lte_data.columns[col_idx]
            # Ensure numeric type before aggregation
            self.lte_data[col_name] = pd.to_numeric(
                self.lte_data[col_name], errors='coerce')
            agg_dict[col_name] = 'sum'

        # Also keep first value of some metadata
        agg_dict[self.lte_data.columns[LTECol.LTE_BEGIN_TIME]] = 'first'

        try:
            grouped = self.lte_data.groupby(
                group_cols, dropna=False).agg(agg_dict).reset_index()
            # Fill NaN with 0 for numeric columns to prevent division errors
            numeric_cols = grouped.select_dtypes(include=[np.number]).columns
            grouped[numeric_cols] = grouped[numeric_cols].fillna(0)
            return grouped
        except Exception as e:
            print(f"Warning: Error during LTE grouping: {e}")
            return self.lte_data

    def group_gsm_by_cluster_month(self) -> pd.DataFrame:
        """
        Group GSM data by TOWER_ID, YEAR_MONTH, BTS_NAME
        Aggregate numerators and denominators with SUM
        """
        group_cols = ['CLUSTER', 'TOWER_ID', 'YEAR_MONTH',
                      self.gsm_data.columns[GSMCol.GSM_BTS_NAME]]

        # Define aggregation
        agg_dict = {}

        # Sum numeric KPI columns (convert to numeric first)
        for col_idx in range(GSMCol.GSM_CSSR_NUM, min(GSMCol.GSM_DROP_DEN + 1, len(self.gsm_data.columns))):
            col_name = self.gsm_data.columns[col_idx]
            # Ensure numeric type before aggregation
            self.gsm_data[col_name] = pd.to_numeric(
                self.gsm_data[col_name], errors='coerce')
            agg_dict[col_name] = 'sum'

        # Keep first value of metadata
        agg_dict[self.gsm_data.columns[GSMCol.GSM_BEGIN_TIME]] = 'first'

        try:
            grouped = self.gsm_data.groupby(
                group_cols, dropna=False).agg(agg_dict).reset_index()
            # Fill NaN with 0 for numeric columns to prevent division errors
            numeric_cols = grouped.select_dtypes(include=[np.number]).columns
            grouped[numeric_cols] = grouped[numeric_cols].fillna(0)
            return grouped
        except Exception as e:
            print(f"Warning: Error during GSM grouping: {e}")
            return self.gsm_data

    def get_unique_clusters(self) -> list:
        """Get list of unique clusters"""
        clusters = set()

        if self.lte_data is not None and 'CLUSTER' in self.lte_data.columns:
            clusters.update(self.lte_data['CLUSTER'].dropna().unique())

        if self.gsm_data is not None and 'CLUSTER' in self.gsm_data.columns:
            clusters.update(self.gsm_data['CLUSTER'].dropna().unique())

        return sorted(list(clusters))

    def get_months_range(self) -> tuple:
        """Get date range of data (first month, last month)"""
        dates = []

        if self.lte_data is not None:
            lte_dates = pd.to_datetime(
                self.lte_data.iloc[:, LTECol.LTE_BEGIN_TIME], errors='coerce')
            dates.extend(lte_dates.dropna().tolist())

        if self.gsm_data is not None:
            gsm_dates = pd.to_datetime(
                self.gsm_data.iloc[:, GSMCol.GSM_BEGIN_TIME], errors='coerce')
            dates.extend(gsm_dates.dropna().tolist())

        if dates:
            dates_series = pd.Series(dates)
            return dates_series.min(), dates_series.max()

        return None, None

    def _safe_datetime_conversion(self, series):
        """Safely convert series to datetime"""
        try:
            return pd.to_datetime(series, errors='coerce')
        except Exception as e:
            print(f"Warning: Error converting to datetime: {e}")
            return series
