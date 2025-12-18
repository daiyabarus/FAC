"""Data loading module for Excel files - FIXED NGI LOADING"""

import pandas as pd
from pathlib import Path
from utils.helpers import clean_numeric


class DataLoader:
    """Load and parse Excel data files"""

    def __init__(self):
        self.lte_data = None
        self.gsm_data = None
        self.cluster_data = None
        self.ngi = None

    def load_lte_file(self, file_path):
        """Load LTE Excel file"""
        print(f"Loading LTE file: {file_path}")

        try:
            df = pd.read_excel(file_path, sheet_name="Sheet0")

            numeric_cols = list(range(19, 61))

            for col_idx in numeric_cols:
                if col_idx < len(df.columns):
                    df.iloc[:, col_idx] = df.iloc[:,
                                                  col_idx].apply(clean_numeric)

            self.lte_data = df
            print(f"✓ Loaded {len(df)} LTE records")
            return df

        except Exception as e:
            print(f"✗ Error loading LTE file: {e}")
            raise

    def load_gsm_file(self, file_path):
        """Load GSM Excel file"""
        print(f"Loading GSM file: {file_path}")

        try:
            df = pd.read_excel(file_path, sheet_name="Sheet0")

            numeric_cols = [13, 14, 15, 16, 17, 18]

            for col_idx in numeric_cols:
                if col_idx < len(df.columns):
                    df.iloc[:, col_idx] = df.iloc[:,
                                                  col_idx].apply(clean_numeric)

            self.gsm_data = df
            print(f"✓ Loaded {len(df)} GSM records")
            return df

        except Exception as e:
            print(f"✗ Error loading GSM file: {e}")
            raise

    def load_cluster_file(self, file_path):
        """Load Cluster Excel file with CAT column"""
        print(f"Loading Cluster file: {file_path}")

        try:
            df = pd.read_excel(file_path, sheet_name="CLUSTER")

            # 🔍 DEBUG: Show columns
            print(f"  Cluster columns: {list(df.columns)}")
            print(f"  Expected columns: CLUSTER, TOWERID, LTE_CELL, TX, SITENAME, CAT")

            # Validate required columns exist
            required_cols = ["CLUSTER", "TOWERID",
                             "LTE_CELL", "TX", "SITENAME", "CAT"]
            missing = [col for col in required_cols if col not in df.columns]
            if missing:
                print(f"  ⚠ Missing columns in Cluster file: {missing}")
                print(f"  Note: CAT column is required for NGI validation")

            self.cluster_data = df
            print(f"✓ Loaded {len(df)} Cluster records")
            return df

        except Exception as e:
            print(f"✗ Error loading Cluster file: {e}")
            raise

    def load_ngi_file(self, path: str):
        """
        Load NVE Grid / NGI file.
        Sheet name: 'NVE Grid'

        Expected columns:
        - eNodeB ID
        - Cell ID
        - Cell Name
        - Total Sampling Points
        - RSRP
        - RSRQ
        - GoodRatio(%)
        - etc.
        """
        if not path or path.strip() == "":
            print("⚠ NGI file not provided, skipping...")
            self.ngi = None
            return

        print(f"Loading NGI file: {path}")

        try:
            df = pd.read_excel(path, sheet_name="NVE Grid")

            # Clean column names (strip whitespace)
            df.columns = [str(c).strip() for c in df.columns]

            # 🔍 DEBUG: Show actual columns
            print(f"  NGI columns found: {list(df.columns)}")

            # Validate required columns
            required = ["Cell Name", "RSRP", "RSRQ"]
            missing = [c for c in required if c not in df.columns]
            if missing:
                print(f"  ❌ NGI file missing required columns: {missing}")
                print(f"  Available columns: {list(df.columns)}")
                raise ValueError(f"NGI file missing columns: {missing}")

            # Skip summary row (where Cell Name is "--" or "All")
            df = df[~df["Cell Name"].astype(
                str).str.upper().isin(["--", "ALL"])].copy()

            # Convert RSRP and RSRQ to numeric
            df["RSRP"] = pd.to_numeric(df["RSRP"], errors="coerce")
            df["RSRQ"] = pd.to_numeric(df["RSRQ"], errors="coerce")

            # Show sample data
            print(f"\n  === NGI Sample Data ===")
            sample_cols = ["Cell Name", "RSRP", "RSRQ"]
            if "eNodeB ID" in df.columns:
                sample_cols.insert(0, "eNodeB ID")
            print(df[sample_cols].head(5).to_string(index=False))

            self.ngi = df
            print(f"✓ Loaded {len(df)} NGI records (excluding summary rows)")
            return df

        except Exception as e:
            print(f"✗ Error loading NGI file: {e}")
            raise

    def get_data(self):
        """Return all loaded data"""
        return {
            "lte": self.lte_data,
            "gsm": self.gsm_data,
            "cluster": self.cluster_data,
            "ngi": self.ngi,
        }
