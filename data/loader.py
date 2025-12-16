"""Data loading module for Excel files"""
import pandas as pd
from pathlib import Path
from utils.helpers import clean_numeric


class DataLoader:
    """Load and parse Excel data files"""

    def __init__(self):
        self.lte_data = None
        self.gsm_data = None
        self.cluster_data = None

    def load_lte_file(self, file_path):
        """Load LTE Excel file"""
        print(f"Loading LTE file: {file_path}")

        try:
            df = pd.read_excel(file_path, sheet_name='Sheet0')

            # Clean numeric columns
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
            df = pd.read_excel(file_path, sheet_name='Sheet0')

            # Clean numeric columns
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
        """Load Cluster Excel file"""
        print(f"Loading Cluster file: {file_path}")

        try:
            df = pd.read_excel(file_path, sheet_name='CLUSTER')

            self.cluster_data = df
            print(f"✓ Loaded {len(df)} Cluster records")
            return df

        except Exception as e:
            print(f"✗ Error loading Cluster file: {e}")
            raise

    def get_data(self):
        """Return all loaded data"""
        return {
            'lte': self.lte_data,
            'gsm': self.gsm_data,
            'cluster': self.cluster_data
        }
