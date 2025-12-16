"""Data transformation module"""
import pandas as pd
import numpy as np
from utils.helpers import (
    extract_tower_id,
    map_frequency_band,
    format_date_mmm_yy
)
from config.settings import LTEColumns, GSMColumns, ClusterColumns


class DataTransformer:
    """Transform and merge data"""

    def __init__(self, data_loader):
        self.loader = data_loader
        self.lte_enriched = None
        self.gsm_enriched = None

    def transform_all(self):
        """Execute all transformations"""
        print("\n=== Data Transformation ===")

        self._enrich_lte_data()
        self._enrich_gsm_data()

        print("✓ Transformation complete")

        return {
            'lte': self.lte_enriched,
            'gsm': self.gsm_enriched
        }

    def _enrich_lte_data(self):
        """Enrich LTE data with cluster info and band mapping"""
        df = self.loader.lte_data.copy()
        cluster_df = self.loader.cluster_data.copy()

        # Extract tower ID
        df['TOWER_ID'] = df.iloc[:, LTEColumns.ME_NAME].apply(extract_tower_id)

        # Map frequency band
        df['LTE_BAND'] = df.iloc[:, LTEColumns.FREQ_BAND].apply(
            map_frequency_band)

        # Add month column
        df['MONTH'] = pd.to_datetime(
            df.iloc[:, LTEColumns.BEGIN_TIME]).apply(format_date_mmm_yy)

        # Merge with cluster data
        cluster_df_merge = cluster_df[[
            cluster_df.columns[ClusterColumns.TOWERID],
            cluster_df.columns[ClusterColumns.CLUSTER],
            cluster_df.columns[ClusterColumns.TX]
        ]].copy()

        cluster_df_merge.columns = ['TOWER_ID', 'CLUSTER', 'TX']

        df = df.merge(cluster_df_merge, on='TOWER_ID', how='left')

        self.lte_enriched = df
        print(f"✓ LTE data enriched: {len(df)} records")

    def _enrich_gsm_data(self):
        """Enrich GSM data with cluster info"""
        df = self.loader.gsm_data.copy()
        cluster_df = self.loader.cluster_data.copy()

        # Add month column
        df['MONTH'] = pd.to_datetime(
            df.iloc[:, GSMColumns.BEGIN_TIME]).apply(format_date_mmm_yy)

        # Merge with cluster data on SITENAME
        cluster_df_merge = cluster_df[[
            cluster_df.columns[ClusterColumns.SITENAME],
            cluster_df.columns[ClusterColumns.CLUSTER],
            cluster_df.columns[ClusterColumns.TOWERID]
        ]].copy()

        cluster_df_merge.columns = ['SITE_NAME', 'CLUSTER', 'TOWER_ID']

        # Merge on site name
        df = df.merge(
            cluster_df_merge,
            left_on=df.columns[GSMColumns.SITE_NAME],
            right_on='SITE_NAME',
            how='left'
        )

        self.gsm_enriched = df
        print(f"✓ GSM data enriched: {len(df)} records")
