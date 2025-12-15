"""
Data Loader Service
Load data from Excel files with proper handling of comma separators and percent signs
"""
import pandas as pd
import re
from pathlib import Path
from models.column_enums import LTECol, GSMCol, ClusterCol


class DataLoader:
    """Load and preprocess Excel data files"""

    def __init__(self):
        self.lte_data = None
        self.gsm_data = None
        self.cluster_data = None

    def load_lte_data(self, file_path: str) -> pd.DataFrame:
        """Load LTE data from Excel file"""
        try:
            # Read with all columns as object first to handle mixed types
            df = pd.read_excel(file_path, sheet_name='Sheet0', dtype=str)

            # Convert BEGIN_TIME to datetime
            df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')

            # Clean and convert numeric columns (from column 19 onwards)
            for col_idx in range(19, len(df.columns)):
                df.iloc[:, col_idx] = self._clean_and_convert_numeric(
                    df.iloc[:, col_idx])

            self.lte_data = df
            return df

        except Exception as e:
            raise Exception(f"Error loading LTE data: {str(e)}")

    def load_gsm_data(self, file_path: str) -> pd.DataFrame:
        """Load GSM data from Excel file"""
        try:
            # Read with all columns as object first
            df = pd.read_excel(file_path, sheet_name='Sheet0', dtype=str)

            # Convert BEGIN_TIME to datetime
            df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')

            # Clean and convert numeric columns (from column 13 onwards)
            for col_idx in range(13, len(df.columns)):
                df.iloc[:, col_idx] = self._clean_and_convert_numeric(
                    df.iloc[:, col_idx])

            self.gsm_data = df
            return df

        except Exception as e:
            raise Exception(f"Error loading GSM data: {str(e)}")

    def load_cluster_data(self, file_path: str) -> pd.DataFrame:
        """Load Cluster mapping data from Excel file"""
        try:
            df = pd.read_excel(file_path, sheet_name='CLUSTER')
            self.cluster_data = df
            return df

        except Exception as e:
            raise Exception(f"Error loading Cluster data: {str(e)}")

    def _clean_and_convert_numeric(self, series: pd.Series) -> pd.Series:
        """
        Clean and convert column to numeric:
        - Remove commas (1,234 -> 1234)
        - Remove percent signs (50% -> 50)
        - Convert to float64
        """
        def clean_value(val):
            if pd.isna(val) or val == '':
                return None

            val_str = str(val).strip()
            # Remove commas and percent signs
            val_str = val_str.replace(',', '').replace('%', '')

            try:
                return float(val_str)
            except:
                return None

        # Apply cleaning and return as float64
        cleaned = series.apply(clean_value)
        return pd.to_numeric(cleaned, errors='coerce')

    def get_tower_id_from_me_name(self, me_name: str) -> str:
        """
        Extract TOWER ID from LTE ME_NAME using regex pattern #([^#]+)#
        Example: "4217035E_LTE_SEUNEUBOK PIDIE ACEH TAMIANG#SUM-AC-LGS-0248#NR(213068)"
        Returns: "SUM-AC-LGS-0248"
        """
        pattern = r"#([^#]+)#"
        match = re.search(pattern, str(me_name))
        if match:
            return match.group(1)
        return None

    def validate_data(self) -> dict:
        """Validate loaded data"""
        validation = {
            'lte_valid': False,
            'gsm_valid': False,
            'cluster_valid': False,
            'errors': []
        }

        if self.lte_data is not None and len(self.lte_data) > 0:
            validation['lte_valid'] = True
            validation['lte_rows'] = len(self.lte_data)
        else:
            validation['errors'].append("LTE data is empty or not loaded")

        if self.gsm_data is not None and len(self.gsm_data) > 0:
            validation['gsm_valid'] = True
            validation['gsm_rows'] = len(self.gsm_data)
        else:
            validation['errors'].append("GSM data is empty or not loaded")

        if self.cluster_data is not None and len(self.cluster_data) > 0:
            validation['cluster_valid'] = True
            validation['cluster_rows'] = len(self.cluster_data)
        else:
            validation['errors'].append("Cluster data is empty or not loaded")

        validation['all_valid'] = (
            validation['lte_valid'] and
            validation['gsm_valid'] and
            validation['cluster_valid']
        )

        return validation
