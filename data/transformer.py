"""Data transformation module - FIXED NGI ENRICHMENT"""

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
        self.ngi_enriched = None

    def transform_all(self):
        """Execute all transformations"""
        print("\n=== Data Transformation ===")
        self._enrich_lte_data()
        self._enrich_gsm_data()
        self._enrich_ngi_data()
        print("✓ Transformation complete")
        return {
            "lte": self.lte_enriched,
            "gsm": self.gsm_enriched,
            "ngi": self.ngi_enriched,
        }

    def _enrich_ngi_data(self):
        """
        Map NGI (NVE Grid) ke CLUSTER via LTE_CELL ↔ Cell Name.

        NGI file format:
        - eNodeB ID, Cell ID, Cell Name, RSRP, RSRQ, etc.

        CLUSTER file format:
        - CLUSTER, TOWERID, LTE_CELL, TX, SITENAME, CAT

        Output format:
        - CLUSTER, TOWER_ID, CELL_NAME, RSRP, RSRQ, CAT
        """
        if self.loader.ngi is None:
            print("⚠ No NGI data loaded, skipping NGI enrichment")
            self.ngi_enriched = None
            return

        cluster_df = self.loader.cluster_data.copy()
        ngi_df = self.loader.ngi.copy()

        # 🔍 DEBUG: Check raw columns
        print("\n=== NGI Raw Columns ===")
        print(f"NGI columns: {list(ngi_df.columns)}")
        print(f"Cluster columns: {list(cluster_df.columns)}")

        # Clean column names (strip whitespace)
        cluster_df.columns = [str(c).strip() for c in cluster_df.columns]
        ngi_df.columns = [str(c).strip() for c in ngi_df.columns]

        # ===== VALIDATE NGI COLUMNS =====
        required_ngi_cols = ["Cell Name", "RSRP", "RSRQ"]
        missing_ngi = [
            col for col in required_ngi_cols if col not in ngi_df.columns]
        if missing_ngi:
            print(f"❌ NGI file missing required columns: {missing_ngi}")
            print(f"Available columns: {list(ngi_df.columns)}")
            self.ngi_enriched = None
            return

        # ===== VALIDATE CLUSTER COLUMNS =====
        # Get column by index from ClusterColumns
        try:
            lte_cell_col = cluster_df.columns[ClusterColumns.LTE_CELL]
            cluster_col = cluster_df.columns[ClusterColumns.CLUSTER]
            towerid_col = cluster_df.columns[ClusterColumns.TOWERID]
            cat_col = cluster_df.columns[ClusterColumns.CAT]
        except IndexError as e:
            print(f"❌ Cluster file column index error: {e}")
            print(f"Available columns: {list(cluster_df.columns)}")
            self.ngi_enriched = None
            return

        # Check if CAT column exists
        if cat_col not in cluster_df.columns:
            print(
                f"❌ Cluster file missing CAT column at index {ClusterColumns.CAT}")
            print(f"Available columns: {list(cluster_df.columns)}")
            self.ngi_enriched = None
            return

        # ===== PREPARE CLUSTER DATA =====
        cluster_view = cluster_df[
            [lte_cell_col, cluster_col, towerid_col, cat_col]
        ].copy()

        cluster_view.columns = ["LTE_CELL", "CLUSTER", "TOWERID", "CAT"]

        # 🔍 DEBUG: Check CAT values
        print(f"\n=== Cluster CAT Values ===")
        print(cluster_view["CAT"].value_counts())

        # Normalize for matching
        cluster_view["LTE_CELL"] = (
            cluster_view["LTE_CELL"].astype(str).str.strip().str.upper()
        )
        cluster_view["CAT"] = (
            cluster_view["CAT"].astype(str).str.strip().str.upper()
        )

        # ===== PREPARE NGI DATA =====
        ngi_df["Cell Name"] = ngi_df["Cell Name"].astype(
            str).str.strip().str.upper()

        # Skip "All" summary row
        ngi_df = ngi_df[ngi_df["Cell Name"] != "--"].copy()

        print(f"\n=== NGI Data Before Merge ===")
        print(f"Total NGI rows: {len(ngi_df)}")
        print(f"Unique Cell Names: {ngi_df['Cell Name'].nunique()}")
        print(f"Sample Cell Names: {ngi_df['Cell Name'].head(10).tolist()}")

        # ===== MERGE NGI WITH CLUSTER =====
        merged = cluster_view.merge(
            ngi_df,
            left_on="LTE_CELL",
            right_on="Cell Name",
            how="inner",  # Only keep matching cells
        )

        print(f"\n=== After Merge ===")
        print(f"Merged rows: {len(merged)}")

        if len(merged) == 0:
            print("⚠ No matching cells found between NGI and Cluster files")
            print("\n=== Sample LTE_CELL from Cluster ===")
            print(cluster_view["LTE_CELL"].head(20).tolist())
            print("\n=== Sample Cell Name from NGI ===")
            print(ngi_df["Cell Name"].head(20).tolist())
            self.ngi_enriched = None
            return

        # ===== VALIDATE REQUIRED COLUMNS =====
        if "RSRP" not in merged.columns or "RSRQ" not in merged.columns:
            print("❌ Merged data missing RSRP or RSRQ columns")
            print(f"Available columns: {list(merged.columns)}")
            self.ngi_enriched = None
            return

        # ===== SELECT AND RENAME COLUMNS =====
        out = merged[["CLUSTER", "TOWERID", "CAT",
                      "LTE_CELL", "RSRP", "RSRQ"]].copy()
        out = out.rename(columns={
            "LTE_CELL": "CELL_NAME",
            "TOWERID": "TOWER_ID"
        })

        # ===== CLEAN DATA =====
        # Drop rows with missing critical data
        out = out.dropna(subset=["CLUSTER", "CAT", "RSRP", "RSRQ"])

        # Convert RSRP and RSRQ to numeric
        out["RSRP"] = pd.to_numeric(out["RSRP"], errors="coerce")
        out["RSRQ"] = pd.to_numeric(out["RSRQ"], errors="coerce")

        # Drop rows where conversion failed
        out = out.dropna(subset=["RSRP", "RSRQ"])

        self.ngi_enriched = out
        print(f"✓ NGI data enriched: {len(out)} records")

        # ===== DEBUG SAMPLE =====
        if len(out) > 0:
            print("\n=== NGI Enriched Sample (first 10 records) ===")
            print(out.head(10).to_string(index=False))
            print("\n=== NGI Summary by Cluster ===")
            print(out.groupby(["CLUSTER", "CAT"]).size())
            print("=" * 80)
        else:
            print("⚠ NGI enriched data is empty after filtering")

    def _enrich_lte_data(self):
        """Enrich LTE data with cluster info and band mapping"""
        df = self.loader.lte_data.copy()
        cluster_df = self.loader.cluster_data.copy()

        # TOWER_ID dari ME_NAME
        df["TOWER_ID"] = df.iloc[:, LTEColumns.ME_NAME].apply(extract_tower_id)

        # Mapping band
        df["LTE_BAND"] = df.iloc[:, LTEColumns.FREQ_BAND].apply(
            map_frequency_band)

        # MONTH (Sep-25, dst)
        df["MONTH"] = pd.to_datetime(df.iloc[:, LTEColumns.BEGIN_TIME]).apply(
            format_date_mmm_yy
        )

        # Ambil kolom cluster yang dibutuhkan (LTE_CELL, CLUSTER, TX, CAT)
        cols = [
            cluster_df.columns[ClusterColumns.LTE_CELL],
            cluster_df.columns[ClusterColumns.CLUSTER],
            cluster_df.columns[ClusterColumns.TX],
        ]

        # CAT column
        if len(cluster_df.columns) > ClusterColumns.CAT:
            cols.append(cluster_df.columns[ClusterColumns.CAT])

        cluster_df_merge = cluster_df[cols].copy()

        # Rename agar generik
        rename_map = {
            cluster_df.columns[ClusterColumns.LTE_CELL]: "CELL_NAME",
            cluster_df.columns[ClusterColumns.CLUSTER]: "CLUSTER",
            cluster_df.columns[ClusterColumns.TX]: "TX",
        }

        if len(cluster_df.columns) > ClusterColumns.CAT:
            rename_map[cluster_df.columns[ClusterColumns.CAT]] = "CAT"

        cluster_df_merge.columns = [
            rename_map.get(c, c) for c in cluster_df_merge.columns
        ]

        cluster_df_merge = cluster_df_merge.drop_duplicates(
            subset=["CELL_NAME"], keep="first"
        )

        lte_cell_col = df.columns[LTEColumns.CELL_NAME]
        df = df.merge(
            cluster_df_merge,
            left_on=lte_cell_col,
            right_on="CELL_NAME",
            how="left",
        )

        if "CELL_NAME" in df.columns and lte_cell_col != "CELL_NAME":
            df.drop(["CELL_NAME"], axis=1, inplace=True)

        print("\n=== DEBUG: Sample Data After Merge ===")
        sample_cols = [lte_cell_col, "TOWER_ID", "LTE_BAND", "CLUSTER"]
        if "TX" in df.columns:
            sample_cols.append("TX")
        if "CAT" in df.columns:
            sample_cols.append("CAT")

        sample = df[sample_cols].drop_duplicates().head(20)
        print(sample.to_string(index=False))
        print("=" * 80)

        total_cells = df[lte_cell_col].nunique()
        cells_with_tx = (
            df[df.get("TX").notna()][lte_cell_col].nunique()
            if "TX" in df.columns
            else 0
        )
        cells_without_tx = total_cells - cells_with_tx

        print(f"\nMapping Summary:")
        print(f"  Total unique cells: {total_cells}")
        print(f"  Cells with TX mapping: {cells_with_tx}")
        print(f"  Cells without TX mapping: {cells_without_tx}")

        if cells_without_tx > 0:
            print(
                f"\n  ⚠ {cells_without_tx} cells not found in cluster file "
                f"(will be excluded from SE validation)"
            )

        self.lte_enriched = df
        print(f"✓ LTE data enriched: {len(df)} records")

    def _enrich_gsm_data(self):
        """Enrich GSM data with cluster info"""
        df = self.loader.gsm_data.copy()
        cluster_df = self.loader.cluster_data.copy()

        df["MONTH"] = pd.to_datetime(df.iloc[:, GSMColumns.BEGIN_TIME]).apply(
            format_date_mmm_yy
        )

        cluster_df_merge = cluster_df[
            [
                cluster_df.columns[ClusterColumns.SITENAME],
                cluster_df.columns[ClusterColumns.CLUSTER],
                cluster_df.columns[ClusterColumns.TOWERID],
            ]
        ].copy()

        cluster_df_merge.columns = ["SITE_NAME", "CLUSTER", "TOWER_ID"]
        cluster_df_merge = cluster_df_merge.drop_duplicates(
            subset=["SITE_NAME"], keep="first"
        )

        gsm_site_col = df.columns[GSMColumns.SITE_NAME]
        df = df.merge(
            cluster_df_merge,
            left_on=gsm_site_col,
            right_on="SITE_NAME",
            how="left",
        )

        if "SITE_NAME" in df.columns and gsm_site_col != "SITE_NAME":
            df.drop(["SITE_NAME"], axis=1, inplace=True)

        self.gsm_enriched = df
        print(f"✓ GSM data enriched: {len(df)} records")
