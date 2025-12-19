"""Data transformation module - UPDATED WITH PERIOD GROUPING"""

import pandas as pd
import numpy as np
from utils.helpers import (
    extract_tower_id, 
    map_frequency_band, 
    format_date_mmm_yy,
    get_latest_date_and_periods,
    assign_period_to_date
)
from config.settings import LTEColumns, GSMColumns, ClusterColumns


class DataTransformer:
    """Transform and merge data"""

    def __init__(self, data_loader):
        self.loader = data_loader
        self.lte_enriched = None
        self.gsm_enriched = None
        self.ngi_enriched = None
        self.period_info = None  # Store period information

    def transform_all(self):
        """Execute all transformations"""
        print("\n=== Data Transformation ===")
        
        # Calculate periods first based on LTE data
        self._calculate_periods()
        
        self._enrich_lte_data()
        self._enrich_gsm_data()
        self._enrich_ngi_data()
        print("✓ Transformation complete")
        return {
            "lte": self.lte_enriched,
            "gsm": self.gsm_enriched,
            "ngi": self.ngi_enriched,
            "period_info": self.period_info
        }

    def _calculate_periods(self):
        """Calculate 3 periods of 30 days from 90 days before latest date"""
        if self.loader.lte_data is None:
            print("⚠ No LTE data to calculate periods")
            return
        
        dates = pd.to_datetime(self.loader.lte_data.iloc[:, LTEColumns.BEGIN_TIME])
        self.period_info = get_latest_date_and_periods(dates)
        
        if self.period_info:
            print("\n=== Period Information ===")
            print(f"Latest date: {self.period_info['latest_date'].strftime('%Y-%m-%d')}")
            print(f"Analysis period: {self.period_info['start_date'].strftime('%Y-%m-%d')} to {self.period_info['latest_date'].strftime('%Y-%m-%d')}")
            print("\nPeriods:")
            for i in range(1, 4):
                p = self.period_info[f'period_{i}']
                print(f"  Period {i} ({p['label']}): {p['start'].strftime('%Y-%m-%d')} to {p['end'].strftime('%Y-%m-%d')}")
            print("=" * 80)

    def _enrich_ngi_data(self):
        """Map NGI to CLUSTER via LTE_CELL → Cell Name"""
        if self.loader.ngi is None:
            print("⚠ No NGI data loaded, skipping NGI enrichment")
            self.ngi_enriched = None
            return

        cluster_df = self.loader.cluster_data.copy()
        ngi_df = self.loader.ngi.copy()

        cluster_df.columns = [str(c).strip() for c in cluster_df.columns]
        ngi_df.columns = [str(c).strip() for c in ngi_df.columns]

        required_ngi_cols = ["Cell Name", "RSRP", "RSRQ"]
        missing_ngi = [col for col in required_ngi_cols if col not in ngi_df.columns]
        if missing_ngi:
            print(f"✗ NGI file missing required columns: {missing_ngi}")
            self.ngi_enriched = None
            return

        try:
            lte_cell_col = cluster_df.columns[ClusterColumns.LTE_CELL]
            cluster_col = cluster_df.columns[ClusterColumns.CLUSTER]
            towerid_col = cluster_df.columns[ClusterColumns.TOWERID]
            cat_col = cluster_df.columns[ClusterColumns.CAT]
        except IndexError as e:
            print(f"✗ Cluster file column index error: {e}")
            self.ngi_enriched = None
            return

        cluster_view = cluster_df[[lte_cell_col, cluster_col, towerid_col, cat_col]].copy()
        cluster_view.columns = ["LTE_CELL", "CLUSTER", "TOWERID", "CAT"]
        cluster_view["LTE_CELL"] = cluster_view["LTE_CELL"].astype(str).str.strip().str.upper()
        cluster_view["CAT"] = cluster_view["CAT"].astype(str).str.strip().str.upper()

        ngi_df["Cell Name"] = ngi_df["Cell Name"].astype(str).str.strip().str.upper()
        ngi_df = ngi_df[ngi_df["Cell Name"] != "--"].copy()

        merged = cluster_view.merge(ngi_df, left_on="LTE_CELL", right_on="Cell Name", how="inner")

        if len(merged) == 0:
            print("⚠ No matching cells found between NGI and Cluster files")
            self.ngi_enriched = None
            return

        out = merged[["CLUSTER", "TOWERID", "CAT", "LTE_CELL", "RSRP", "RSRQ"]].copy()
        out = out.rename(columns={"LTE_CELL": "CELL_NAME", "TOWERID": "TOWER_ID"})
        out = out.dropna(subset=["CLUSTER", "CAT", "RSRP", "RSRQ"])
        out["RSRP"] = pd.to_numeric(out["RSRP"], errors="coerce")
        out["RSRQ"] = pd.to_numeric(out["RSRQ"], errors="coerce")
        out = out.dropna(subset=["RSRP", "RSRQ"])

        self.ngi_enriched = out
        print(f"✓ NGI data enriched: {len(out)} records")

    def _enrich_lte_data(self):
        """Enrich LTE data with cluster info, band mapping, and period assignment"""
        df = self.loader.lte_data.copy()
        cluster_df = self.loader.cluster_data.copy()

        df["TOWER_ID"] = df.iloc[:, LTEColumns.ME_NAME].apply(extract_tower_id)
        df["LTE_BAND"] = df.iloc[:, LTEColumns.FREQ_BAND].apply(map_frequency_band)
        
        # Assign PERIOD instead of MONTH
        if self.period_info:
            df["PERIOD"] = pd.to_datetime(df.iloc[:, LTEColumns.BEGIN_TIME]).apply(
                lambda x: assign_period_to_date(x, self.period_info)
            )
            # Keep MONTH for backward compatibility (charts, etc)
            df["MONTH"] = pd.to_datetime(df.iloc[:, LTEColumns.BEGIN_TIME]).apply(format_date_mmm_yy)
        else:
            df["PERIOD"] = None
            df["MONTH"] = pd.to_datetime(df.iloc[:, LTEColumns.BEGIN_TIME]).apply(format_date_mmm_yy)

        cols = [
            cluster_df.columns[ClusterColumns.LTE_CELL],
            cluster_df.columns[ClusterColumns.CLUSTER],
            cluster_df.columns[ClusterColumns.TX],
        ]

        if len(cluster_df.columns) > ClusterColumns.CAT:
            cols.append(cluster_df.columns[ClusterColumns.CAT])

        cluster_df_merge = cluster_df[cols].copy()
        rename_map = {
            cluster_df.columns[ClusterColumns.LTE_CELL]: "CELL_NAME",
            cluster_df.columns[ClusterColumns.CLUSTER]: "CLUSTER",
            cluster_df.columns[ClusterColumns.TX]: "TX",
        }

        if len(cluster_df.columns) > ClusterColumns.CAT:
            rename_map[cluster_df.columns[ClusterColumns.CAT]] = "CAT"

        cluster_df_merge.columns = [rename_map.get(c, c) for c in cluster_df_merge.columns]
        cluster_df_merge = cluster_df_merge.drop_duplicates(subset=["CELL_NAME"], keep="first")

        lte_cell_col = df.columns[LTEColumns.CELL_NAME]
        df = df.merge(cluster_df_merge, left_on=lte_cell_col, right_on="CELL_NAME", how="left")

        if "CELL_NAME" in df.columns and lte_cell_col != "CELL_NAME":
            df.drop(["CELL_NAME"], axis=1, inplace=True)

        self.lte_enriched = df
        print(f"✓ LTE data enriched: {len(df)} records")

    def _enrich_gsm_data(self):
        """Enrich GSM data with cluster info and period assignment"""
        df = self.loader.gsm_data.copy()
        cluster_df = self.loader.cluster_data.copy()

        # Assign PERIOD instead of MONTH
        if self.period_info:
            df["PERIOD"] = pd.to_datetime(df.iloc[:, GSMColumns.BEGIN_TIME]).apply(
                lambda x: assign_period_to_date(x, self.period_info)
            )
            df["MONTH"] = pd.to_datetime(df.iloc[:, GSMColumns.BEGIN_TIME]).apply(format_date_mmm_yy)
        else:
            df["PERIOD"] = None
            df["MONTH"] = pd.to_datetime(df.iloc[:, GSMColumns.BEGIN_TIME]).apply(format_date_mmm_yy)

        cluster_df_merge = cluster_df[
            [
                cluster_df.columns[ClusterColumns.SITENAME],
                cluster_df.columns[ClusterColumns.CLUSTER],
                cluster_df.columns[ClusterColumns.TOWERID],
            ]
        ].copy()

        cluster_df_merge.columns = ["SITE_NAME", "CLUSTER", "TOWER_ID"]
        cluster_df_merge = cluster_df_merge.drop_duplicates(subset=["SITE_NAME"], keep="first")

        gsm_site_col = df.columns[GSMColumns.SITE_NAME]
        df = df.merge(cluster_df_merge, left_on=gsm_site_col, right_on="SITE_NAME", how="left")

        if "SITE_NAME" in df.columns and gsm_site_col != "SITE_NAME":
            df.drop(["SITE_NAME"], axis=1, inplace=True)

        self.gsm_enriched = df
        print(f"✓ GSM data enriched: {len(df)} records")