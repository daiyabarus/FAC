"""Chart generation module"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime
import numpy as np
from io import BytesIO
import base64


class ChartGenerator:
    """Generate charts for KPI trends"""

    def __init__(self, kpi_data, cluster):
        self.kpi_data = kpi_data
        self.cluster = cluster

    def generate_all_charts(self):
        """Generate all KPI trend charts"""
        print(f"\n=== Generating charts for {self.cluster} ===")

        charts = {}

        # Filter data for cluster
        lte_data = self.kpi_data["lte"]
        lte_cluster = lte_data[lte_data["CLUSTER"] == self.cluster].copy()

        if len(lte_cluster) == 0:
            print("⚠ No data for charts")
            return charts

        # Generate charts for key KPIs
        kpi_list = [
            ("SESSION_SSR", "Session Setup Success Rate", 99),
            ("RACH_SR", "RACH Success Rate", 85),
            ("HO_SR", "Handover Success Rate", 97),
            ("ERAB_DROP", "E-RAB Drop Rate", 2),
            ("DL_THP", "Downlink Throughput (Mbps)", 3),
            ("UL_THP", "Uplink Throughput (Mbps)", 1),
        ]

        for kpi_col, kpi_name, baseline in kpi_list:
            try:
                chart_img = self._generate_kpi_chart(
                    lte_cluster, kpi_col, kpi_name, baseline
                )
                if chart_img:
                    charts[kpi_col] = chart_img
            except Exception as e:
                print(f"⚠ Could not generate chart for {kpi_name}: {e}")

        print(f"✓ Generated {len(charts)} charts")
        return charts

    def _generate_kpi_chart(self, df, kpi_col, kpi_name, baseline):
        """Generate a single KPI trend chart"""
        # Group by date
        df_chart = df.copy()
        df_chart["DATE"] = pd.to_datetime(df_chart.iloc[:, 0])  # BEGIN_TIME

        # Calculate daily average
        daily_avg = df_chart.groupby("DATE")[kpi_col].mean().reset_index()
        daily_avg = daily_avg.dropna()

        if len(daily_avg) == 0:
            return None

        # Create figure
        fig, ax = plt.subplots(figsize=(12, 6))

        # Plot KPI
        ax.plot(
            daily_avg["DATE"],
            daily_avg[kpi_col],
            marker="o",
            linewidth=2,
            markersize=4,
            label=kpi_name,
        )

        # Plot baseline
        ax.axhline(
            y=baseline,
            color="r",
            linestyle="--",
            linewidth=1.5,
            label=f"Baseline ({baseline})",
        )

        # Formatting
        ax.set_xlabel("Date", fontsize=12)
        ax.set_ylabel(kpi_name, fontsize=12)
        ax.set_title(f"{kpi_name} - {self.cluster}", fontsize=14, fontweight="bold")
        ax.legend(loc="best")
        ax.grid(True, alpha=0.3)

        # Format x-axis
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b"))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=5))
        plt.xticks(rotation=45)

        plt.tight_layout()

        # Save to BytesIO
        buf = BytesIO()
        plt.savefig(buf, format="png", dpi=100, bbox_inches="tight")
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode()
        plt.close(fig)

        return img_base64
