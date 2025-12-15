# Run this in your Python console
import pandas as pd

# Load your LTE CSV
lte_df = pd.read_csv("data/CLUSTER.csv")

# Print column info
print("=" * 80)
print("COLUMN NAMES:")
print(lte_df.columns.tolist())
print("\n" + "=" * 80)
print("FIRST 3 ROWS:")
print(lte_df.head(3))
print("\n" + "=" * 80)
print("COLUMN INDEX FOR TX (if exists):")
if "TX" in lte_df.columns:
    print(f"TX column index: {lte_df.columns.get_loc('TX')}")
    print(f"TX unique values: {lte_df['TX'].unique()}")
else:
    print("TX column NOT FOUND")

print("\n" + "=" * 80)
print("COLUMN INDEX FOR BAND (if exists):")
if "BAND" in lte_df.columns:
    print(f"BAND column index: {lte_df.columns.get_loc('BAND')}")
    print(f"BAND unique values: {lte_df['BAND'].unique()}")
else:
    print("BAND column NOT FOUND")

print("\n" + "=" * 80)
print("Check column_enums.py:")
print(f"LTECol.LTE_TX = ? (should be column index for TX)")
print(f"LTECol.LTE_BAND = ? (should be column index for BAND or mapped from SITE_NAME)")
