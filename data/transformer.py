"""Data transformation module"""

import pandas as pd
import numpy as np
from utils.helpers import extract_tower_id, map_frequency_band, format_date_mmm_yy
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

        return {"lte": self.lte_enriched, "gsm": self.gsm_enriched}

    def _enrich_lte_data(self):
        """Enrich LTE data with cluster info and band mapping"""
        df = self.loader.lte_data.copy()
        cluster_df = self.loader.cluster_data.copy()

        # Extract tower ID
        df["TOWER_ID"] = df.iloc[:, LTEColumns.ME_NAME].apply(extract_tower_id)

        # Map frequency band
        df["LTE_BAND"] = df.iloc[:, LTEColumns.FREQ_BAND].apply(map_frequency_band)

        # Add month column
        df["MONTH"] = pd.to_datetime(df.iloc[:, LTEColumns.BEGIN_TIME]).apply(
            format_date_mmm_yy
        )

        # ========================================
        # SIMPLE MERGE by CELL NAME
        # ========================================

        # Prepare cluster mapping
        cluster_df_merge = cluster_df[
            [
                cluster_df.columns[ClusterColumns.LTE_CELL],
                cluster_df.columns[ClusterColumns.CLUSTER],
                cluster_df.columns[ClusterColumns.TX],
            ]
        ].copy()

        cluster_df_merge.columns = ["CELL_NAME", "CLUSTER", "TX"]

        # Remove duplicates - keep first
        cluster_df_merge = cluster_df_merge.drop_duplicates(
            subset=["CELL_NAME"], keep="first"
        )

        # Get LTE cell name column
        lte_cell_col = df.columns[LTEColumns.CELL_NAME]

        # Simple merge - if not match, TX and CLUSTER will be None
        df = df.merge(
            cluster_df_merge, left_on=lte_cell_col, right_on="CELL_NAME", how="left"
        )

        # Drop duplicate CELL_NAME column from merge
        if "CELL_NAME" in df.columns and lte_cell_col != "CELL_NAME":
            df.drop(["CELL_NAME"], axis=1, inplace=True)

        # DEBUG: Show sample
        print("\n=== DEBUG: Sample Data After Merge ===")
        sample = (
            df[[lte_cell_col, "TOWER_ID", "TX", "LTE_BAND", "CLUSTER"]]
            .drop_duplicates()
            .head(20)
        )
        print(sample.to_string(index=False))
        print("=" * 80)

        # Count cells with/without TX mapping
        total_cells = df[lte_cell_col].nunique()
        cells_with_tx = df[df["TX"].notna()][lte_cell_col].nunique()
        cells_without_tx = total_cells - cells_with_tx

        print(f"\nMapping Summary:")
        print(f"  Total unique cells: {total_cells}")
        print(f"  Cells with TX mapping: {cells_with_tx}")
        print(f"  Cells without TX mapping: {cells_without_tx}")

        if cells_without_tx > 0:
            print(
                f"\n  ⚠ {cells_without_tx} cells not found in cluster file (will be excluded from SE validation)"
            )

        self.lte_enriched = df
        print(f"✓ LTE data enriched: {len(df)} records")

    def _enrich_gsm_data(self):
        """Enrich GSM data with cluster info"""
        df = self.loader.gsm_data.copy()
        cluster_df = self.loader.cluster_data.copy()

        # Add month column
        df["MONTH"] = pd.to_datetime(df.iloc[:, GSMColumns.BEGIN_TIME]).apply(
            format_date_mmm_yy
        )

        # Prepare cluster mapping
        cluster_df_merge = cluster_df[
            [
                cluster_df.columns[ClusterColumns.SITENAME],
                cluster_df.columns[ClusterColumns.CLUSTER],
                cluster_df.columns[ClusterColumns.TOWERID],
            ]
        ].copy()

        cluster_df_merge.columns = ["SITE_NAME", "CLUSTER", "TOWER_ID"]

        # Remove duplicates
        cluster_df_merge = cluster_df_merge.drop_duplicates(
            subset=["SITE_NAME"], keep="first"
        )

        # Get GSM site name column
        gsm_site_col = df.columns[GSMColumns.SITE_NAME]

        # Simple merge
        df = df.merge(
            cluster_df_merge, left_on=gsm_site_col, right_on="SITE_NAME", how="left"
        )

        # Drop duplicate SITE_NAME column
        if "SITE_NAME" in df.columns and gsm_site_col != "SITE_NAME":
            df.drop(["SITE_NAME"], axis=1, inplace=True)

        self.gsm_enriched = df
        print(f"✓ GSM data enriched: {len(df)} records")
