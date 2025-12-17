"""Debug Spectral Efficiency data"""

import pandas as pd
from data.loader import DataLoader
from data.transformer import DataTransformer

# Load data
loader = DataLoader()
# Ganti dengan path file anda
loader.load_lte_file(
    "D:/Git/FAC/files/Performance_Management_History_Query_4G_FAC_MULTINETYPE_ZTE_DaiyaBarus.xlsx"
)
loader.load_cluster_file("D:/Git/FAC/files/CLUSTER_OK.xlsx")

# Transform
transformer = DataTransformer(loader)
transformed = transformer.transform_all()

lte = transformed["lte"]

# Filter September untuk cluster pertama
cluster = "Z_11.74_CL01_Kota Langsa"
month = "Sep-25"

lte_filtered = lte[(lte["CLUSTER"] == cluster) & (lte["MONTH"] == month)]

print(f"\n=== SE Data for {cluster} - {month} ===")
print(f"Total cells: {len(lte_filtered)}")

# Check each SE condition
conditions = [
    ("2T2R", 850, 1.1, "Row 36"),
    ("2T2R", 900, 1.1, "Row 37"),
    ("2T2R", 2100, 1.3, "Row 38"),
    ("2T2R", 1800, 1.25, "Row 39"),
    (["4T4R", "8T8R"], 1800, 1.5, "Row 40"),
    (["4T4R", "8T8R"], 2100, 1.7, "Row 41"),
    ("32T32R", 2300, 2.1, "Row 43"),
]

for tx, band, baseline, row_label in conditions:
    if isinstance(tx, list):
        mask_tx = lte_filtered["TX"].isin(tx)
        tx_str = " or ".join(tx)
    else:
        mask_tx = lte_filtered["TX"] == tx
        tx_str = tx

    if isinstance(band, list):
        mask_band = lte_filtered["LTE_BAND"].isin(band)
        band_str = " or ".join([str(b) for b in band])
    else:
        mask_band = lte_filtered["LTE_BAND"] == band
        band_str = str(band)

    filtered = lte_filtered[mask_tx & mask_band]
    se_vals = filtered["SPECTRAL_EFF"].dropna()

    if len(se_vals) > 0:
        pass_count = (se_vals >= baseline).sum()
        pass_pct = pass_count / len(se_vals) * 100
        status = "PASS" if pass_pct >= 90 else "FAIL"

        print(f"\n{row_label}: TX={tx_str}, Band={band_str}")
        print(f"  Total cells: {len(se_vals)}")
        print(f"  Pass count (>= {baseline}): {pass_count}")
        print(f"  Pass %: {pass_pct:.2f}%")
        print(f"  Status: {status}")

        # Sample values
        print(f"  SE values (first 5): {list(se_vals.head())}")
        print(f"  SE mean: {se_vals.mean():.2f}")
        print(f"  SE min: {se_vals.min():.2f}")
        print(f"  SE max: {se_vals.max():.2f}")
    else:
        print(f"\n{row_label}: TX={tx_str}, Band={band_str}")
        print(f"  NO DATA")
