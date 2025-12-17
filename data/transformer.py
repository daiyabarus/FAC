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
        self.ngi_enriched = None

    def transform_all(self):
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
        """Map NGI (NVE Grid) ke CLUSTER via LTE_CELL ↔ Cell Name."""
        if self.loader.ngi is None:
            print("⚠ No NGI data loaded, skipping NGI enrichment")
            self.ngi_enriched = None
            return

        cluster_df = self.loader.cluster_data.copy()
        ngi_df = self.loader.ngi.copy()

        cluster_df.columns = [str(c).strip() for c in cluster_df.columns]
        ngi_df.columns = [str(c).strip() for c in ngi_df.columns]

        # Pastikan CAT ada
        if "CAT" not in cluster_df.columns:
            cluster_df["CAT"] = np.nan

        # Ambil LTE_CELL, CLUSTER, TOWERID, CAT
        cluster_view = cluster_df[
            [
                cluster_df.columns[ClusterColumns.LTE_CELL],
                cluster_df.columns[ClusterColumns.CLUSTER],
                cluster_df.columns[ClusterColumns.TOWERID],
                "CAT",
            ]
        ].copy()
        cluster_view["LTE_CELL"] = (
            cluster_view["LTE_CELL"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        ngi_df["Cell Name"] = (
            ngi_df["Cell Name"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        if "Cell Name" not in ngi_df.columns:
            raise ValueError("NGI file must contain 'Cell Name' column")

        merged = cluster_view.merge(
            ngi_df,
            left_on="LTE_CELL",
            right_on="Cell Name",
            how="left",
        )

        if "RSRP" not in merged.columns or "RSRQ" not in merged.columns:
            raise ValueError("NGI file must contain 'RSRP' and 'RSRQ' columns")

        out = merged[["CLUSTER", "TOWERID", "CAT", "LTE_CELL", "RSRP", "RSRQ"]].copy()
        out = out.rename(columns={"LTE_CELL": "CELL_NAME"})

        self.ngi_enriched = out
        print(f"✓ NGI data enriched: {len(out)} records")
    def _enrich_lte_data(self):
        """Enrich LTE data with cluster info and band mapping"""
        df = self.loader.lte_data.copy()
        cluster_df = self.loader.cluster_data.copy()

        # TOWER_ID dari ME_NAME
        df["TOWER_ID"] = df.iloc[:, LTEColumns.ME_NAME].apply(extract_tower_id)
        # Mapping band
        df["LTE_BAND"] = df.iloc[:, LTEColumns.FREQ_BAND].apply(map_frequency_band)
        # MONTH (Sep-25, dst)
        df["MONTH"] = pd.to_datetime(
            df.iloc[:, LTEColumns.BEGIN_TIME]
        ).apply(format_date_mmm_yy)

        # Ambil kolom cluster yang dibutuhkan (LTE_CELL, CLUSTER, TX, CAT)
        cols = [
            cluster_df.columns[ClusterColumns.LTE_CELL],
            cluster_df.columns[ClusterColumns.CLUSTER],
            cluster_df.columns[ClusterColumns.TX],
        ]
        # CAT mungkin belum ada di semua file, jadi cek dulu
        if hasattr(ClusterColumns, "CAT"):
            cols.append(cluster_df.columns[ClusterColumns.CAT])

        cluster_df_merge = cluster_df[cols].copy()

        # Rename agar generik
        rename_map = {
            cluster_df.columns[ClusterColumns.LTE_CELL]: "CELL_NAME",
            cluster_df.columns[ClusterColumns.CLUSTER]: "CLUSTER",
            cluster_df.columns[ClusterColumns.TX]: "TX",
        }
        if hasattr(ClusterColumns, "CAT"):
            rename_map[cluster_df.columns[ClusterColumns.CAT]] = "CAT"

        cluster_df_merge.columns = [rename_map.get(c, c) for c in cluster_df_merge.columns]

        cluster_df_merge = cluster_df_merge.drop_duplicates(subset=["CELL_NAME"], keep="first")

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
        cells_with_tx = df[df.get("TX").notna()][lte_cell_col].nunique() if "TX" in df.columns else 0
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

        df["MONTH"] = pd.to_datetime(
            df.iloc[:, GSMColumns.BEGIN_TIME]
        ).apply(format_date_mmm_yy)

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

