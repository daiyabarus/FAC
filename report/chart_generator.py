"""Chart generation module"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime
import numpy as np
from io import BytesIO
import base64
from config.settings import LTEColumns, GSMColumns


class ChartGenerator:
    """Generate charts for KPI trends"""

    def __init__(self, kpi_data, transformed_data, cluster):
        self.kpi_data = kpi_data
        self.transformed_data = transformed_data
        self.cluster = cluster
        plt.style.use("seaborn-v0_8-darkgrid")

    def generate_all_charts(self):
        """Generate all KPI trend charts"""
        print(f"\n=== Generating charts for {self.cluster} ===")

        charts = {}

        lte_data = self.kpi_data["lte"]
        gsm_data = self.kpi_data["gsm"]

        lte_cluster = lte_data[lte_data["CLUSTER"] == self.cluster].copy()
        gsm_cluster = gsm_data[gsm_data["CLUSTER"] == self.cluster].copy()

        if len(lte_cluster) == 0 and len(gsm_cluster) == 0:
            print("⚠ No data for charts")
            return charts

        gsm_kpis = [
            ("CSSR", "Call Setup Success Rate (%)", 98.5, "higher"),
            ("SDCCH_SR", "SDCCH Success Rate (%)", 98.5, "higher"),
            ("DROP_RATE", "Perceive Drop Rate (%)", 2, "lower"),
        ]

        for kpi_col, kpi_name, baseline, direction in gsm_kpis:
            try:
                chart_img = self._generate_chart(
                    gsm_cluster, kpi_col, kpi_name, baseline, "2G RAN", is_ratio=False
                )
                if chart_img:
                    charts[f"2G_{kpi_col}"] = chart_img
                    print(f"✓ Generated chart: {kpi_name}")
            except Exception as e:
                print(f"⚠ Could not generate chart for {kpi_name}: {e}")

         lte_kpis = [
            ("SESSION_SSR", "Session Setup Success Rate (%)", 99, True),
            ("RACH_SR", "RACH Success Rate (%)", 85, True),
            ("HO_SR", "Handover Success Rate (%)", 97, True),
            ("ERAB_DROP", "E-RAB Drop Rate (%)", 2, True),
            ("DL_THP", "Downlink User Throughput (Mbps)", 3, True),
            ("UL_THP", "Uplink User Throughput (Mbps)", 1, True),
            ("UL_PLOSS", "UL Packet Loss (PDCP) (%)", 0.85, False),
            ("DL_PLOSS", "DL Packet Loss (PDCP) (%)", 0.10, False),
            ("CQI", "CQI", 7, True),
            ("MIMO_RANK2", "MIMO Transmission Rank2 Rate (%)", 35, True),
            ("UL_RSSI", "UL RSSI (dBm)", -105, True),
            ("LATENCY", "Packet Latency (ms)", 30, True),
            ("LTC_NON_CAP", "LTC Non Capacity (%)", 3, False),
            ("OVERLAP_RATE", "Coverage Overlapping Ratio (%)", 35, False),
            ("VOLTE_CSSR", "VoLTE Call Success Rate (%)", 97, True),
            ("VOLTE_DROP", "VoLTE Call Drop Rate (%)", 2, True),
            ("SRVCC_SR", "SRVCC Success Rate (%)", 97, True),
        ]

        for kpi_col, kpi_name, baseline, is_ratio in lte_kpis:
            try:
                chart_img = self._generate_chart(
                    lte_cluster,
                    kpi_col,
                    kpi_name,
                    baseline,
                    "4G RAN",
                    is_ratio=is_ratio,
                )
                if chart_img:
                    charts[f"4G_{kpi_col}"] = chart_img
                    print(f"✓ Generated chart: {kpi_name}")
            except Exception as e:
                print(f"⚠ Could not generate chart for {kpi_name}: {e}")
        se_configs = [
            ("2T2R", 850, 1.1, "SE 2T2R 850MHz"),
            ("2T2R", 900, 1.1, "SE 2T2R 900MHz"),
            ("2T2R", 2100, 1.3, "SE 2T2R 2100MHz"),
            ("2T2R", 1800, 1.25, "SE 2T2R 1800MHz"),
            (["4T4R", "8T8R"], 1800, 1.5, "SE 4T4R/8T8R 1800MHz"),
            (["4T4R", "8T8R"], [2100, 2300], 1.7, "SE 4T4R/8T8R 2100/2300MHz"),
            ("32T32R", 2300, 2.1, "SE 32T32R 2300MHz"),
        ]

        for tx_cond, band_cond, baseline, chart_name in se_configs:
            try:
                chart_img = self._generate_se_chart(
                    lte_cluster, tx_cond, band_cond, baseline, chart_name
                )
                if chart_img:
                    safe_name = chart_name.replace(" ", "_").replace("/", "_")
                    charts[f"4G_SE_{safe_name}"] = chart_img
                    print(f"✓ Generated chart: {chart_name}")
            except Exception as e:
                print(f"⚠ Could not generate chart for {chart_name}: {e}")

        print(f"✓ Generated {len(charts)} charts total")
        return charts

    def _generate_chart(self, df, kpi_col, kpi_name, baseline, tech, is_ratio=False):
        """Generate a single KPI trend chart"""
        if len(df) == 0:
            return None

        if tech == "2G RAN":
            time_col = df.columns[GSMColumns.BEGIN_TIME]
        else:
            time_col = df.columns[LTEColumns.BEGIN_TIME]

        df_chart = df.copy()
        df_chart["DATE"] = pd.to_datetime(df_chart[time_col])


        if is_ratio:
            daily_agg = df_chart.groupby("DATE")[kpi_col].mean().reset_index()
        else:
            daily_agg = df_chart.groupby("DATE")[kpi_col].mean().reset_index()

        daily_agg = daily_agg.dropna()

        if len(daily_agg) == 0:
            return None

        daily_agg = daily_agg.sort_values("DATE")

        fig, ax = plt.subplots(figsize=(14, 7))
        fig.patch.set_edgecolor("black")
        fig.patch.set_linewidth(2)

        ax.plot(
            daily_agg["DATE"],
            daily_agg[kpi_col],
            marker="o",
            linewidth=2,
            markersize=5,
            label=kpi_name,
            color="#2E86AB",
            alpha=0.8,
        )

        ax.axhline(
            y=baseline,
            color="red",
            linestyle="dashdot",
            linewidth=2,
            label=f"Baseline ({baseline})",
            alpha=0.7,
        )

        ax.set_xlabel("Date", fontsize=13, fontweight="bold")
        ax.set_ylabel(kpi_name, fontsize=13, fontweight="bold")
        ax.set_title(
            f"{kpi_name} - {self.cluster} ({tech})",
            fontsize=15,
            fontweight="bold",
            pad=20,
        )

        ax.grid(True, alpha=0.3, linestyle="--", linewidth=0.5)
        ax.legend(loc="best", fontsize=11, framealpha=0.9)
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b-%y"))
        ax.xaxis.set_major_locator(
            mdates.DayLocator(interval=max(1, len(daily_agg) // 15))
        )
        plt.xticks(rotation=45, ha="right")
        if len(daily_agg) > 0:
            first_val = daily_agg.iloc[0][kpi_col]
            ax.annotate(
                f"{first_val:.2f}",
                xy=(daily_agg.iloc[0]["DATE"], first_val),
                xytext=(10, 10),
                textcoords="offset points",
                fontsize=9,
                alpha=0.7,
            )
            last_val = daily_agg.iloc[-1][kpi_col]
            ax.annotate(
                f"{last_val:.2f}",
                xy=(daily_agg.iloc[-1]["DATE"], last_val),
                xytext=(10, -15),
                textcoords="offset points",
                fontsize=9,
                alpha=0.7,
            )

        plt.tight_layout()
        buf = BytesIO()
        plt.savefig(
            buf,
            format="png",
            dpi=100,
            bbox_inches="tight",
            facecolor="white",
            edgecolor="black",
        )
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode()
        plt.close(fig)

        return img_base64

    def _generate_chart_with_numden(
        self, df, num_col, den_col, kpi_name, baseline, tech
    ):
        """Generate chart with NUM/DEN aggregation"""
        if len(df) == 0:
            return None
        if tech == "2G RAN":
            time_col = df.columns[GSMColumns.BEGIN_TIME]
        else:
            time_col = df.columns[LTEColumns.BEGIN_TIME]
        df_chart = df.copy()
        df_chart["DATE"] = pd.to_datetime(df_chart[time_col])
        if tech == "2G RAN":
            num_idx = getattr(GSMColumns, num_col)
            den_idx = getattr(GSMColumns, den_col)
        else:
            num_idx = getattr(LTEColumns, num_col)
            den_idx = getattr(LTEColumns, den_col)
        daily_agg = (
            df_chart.groupby("DATE")
            .agg({df_chart.columns[num_idx]: "sum", df_chart.columns[den_idx]: "sum"})
            .reset_index()
        )
        daily_agg["KPI_VALUE"] = np.where(
            daily_agg[df_chart.columns[den_idx]] > 0,
            (
                daily_agg[df_chart.columns[num_idx]]
                / daily_agg[df_chart.columns[den_idx]]
            )
            * 100,
            None,
        )

        daily_agg = daily_agg.dropna(subset=["KPI_VALUE"])

        if len(daily_agg) == 0:
            return None

        daily_agg = daily_agg.sort_values("DATE")

        fig, ax = plt.subplots(figsize=(14, 7))
        ax.plot(
            daily_agg["DATE"],
            daily_agg["KPI_VALUE"],
            marker="o",
            linewidth=2,
            markersize=5,
            label=kpi_name,
            color="#2E86AB",
            alpha=0.8,
        )
        ax.axhline(
            y=baseline,
            color="red",
            linestyle="dashdot",
            linewidth=2,
            label=f"Baseline ({baseline})",
            alpha=0.7,
        )

        ax.set_xlabel("Date", fontsize=13, fontweight="bold")
        ax.set_ylabel(kpi_name, fontsize=13, fontweight="bold")
        ax.set_title(
            f"{kpi_name} - {self.cluster} ({tech})",
            fontsize=15,
            fontweight="bold",
            pad=20,
        )

        ax.grid(True, alpha=0.3, linestyle="--", linewidth=0.5)
        ax.legend(loc="best", fontsize=11, framealpha=0.9)
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b-%y"))
        ax.xaxis.set_major_locator(
            mdates.DayLocator(interval=max(1, len(daily_agg) // 15))
        )
        plt.xticks(rotation=45, ha="right")
        if len(daily_agg) > 0:
            first_val = daily_agg.iloc[0]["KPI_VALUE"]
            ax.annotate(
                f"{first_val:.2f}",
                xy=(daily_agg.iloc[0]["DATE"], first_val),
                xytext=(10, 10),
                textcoords="offset points",
                fontsize=9,
                alpha=0.7,
            )

            last_val = daily_agg.iloc[-1]["KPI_VALUE"]
            ax.annotate(
                f"{last_val:.2f}",
                xy=(daily_agg.iloc[-1]["DATE"], last_val),
                xytext=(10, -15),
                textcoords="offset points",
                fontsize=9,
                alpha=0.7,
            )

        plt.tight_layout()
        buf = BytesIO()
        plt.savefig(buf, format="png", dpi=100, bbox_inches="tight")
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode()
        plt.close(fig)

        return img_base64

    def _generate_se_chart(self, df, tx_cond, band_cond, baseline, chart_name):
        """Generate Spectral Efficiency chart for specific TX/Band combination"""
        if len(df) == 0:
            return None
        if isinstance(tx_cond, list):
            mask_tx = df["TX"].isin(tx_cond)
        else:
            mask_tx = df["TX"] == tx_cond
        if isinstance(band_cond, list):
            mask_band = df["LTE_BAND"].isin(band_cond)
        else:
            mask_band = df["LTE_BAND"] == band_cond
        filtered_df = df[mask_tx & mask_band].copy()

        if len(filtered_df) == 0:
            print(f"⚠ No data for {chart_name}")
            return None

        time_col = filtered_df.columns[LTEColumns.BEGIN_TIME]
        filtered_df["DATE"] = pd.to_datetime(filtered_df[time_col])
        daily_agg = filtered_df.groupby("DATE")["SPECTRAL_EFF"].mean().reset_index()
        daily_agg = daily_agg.dropna()

        if len(daily_agg) == 0:
            return None
        daily_agg = daily_agg.sort_values("DATE")
        fig, ax = plt.subplots(figsize=(14, 7))
        fig.patch.set_edgecolor("black")
        fig.patch.set_linewidth(2)
        ax.plot(
            daily_agg["DATE"],
            daily_agg["SPECTRAL_EFF"],
            marker="o",
            linewidth=2,
            markersize=5,
            label=f"{chart_name}",
            color="#2E86AB",
            alpha=0.8,
        )
        ax.axhline(
            y=baseline,
            color="red",
            linestyle="dashdot",
            linewidth=2,
            label=f"Baseline ({baseline})",
            alpha=0.7,
        )

        ax.set_xlabel("Date", fontsize=13, fontweight="bold")
        ax.set_ylabel("Spectral Efficiency (bps/Hz)", fontsize=13, fontweight="bold")
        if isinstance(tx_cond, list):
            tx_str = "/".join(tx_cond)
        else:
            tx_str = tx_cond

        if isinstance(band_cond, list):
            band_str = "/".join([str(b) for b in band_cond])
        else:
            band_str = str(band_cond)

        ax.set_title(
            f"Spectral Efficiency - {self.cluster}\n{tx_str} - {band_str}MHz (4G RAN)",
            fontsize=15,
            fontweight="bold",
            pad=20,
        )

        # Grid
        ax.grid(True, alpha=0.3, linestyle="--", linewidth=0.5)

        # Legend
        ax.legend(loc="best", fontsize=11, framealpha=0.9)

        # Format x-axis dates
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b-%y"))
        ax.xaxis.set_major_locator(
            mdates.DayLocator(interval=max(1, len(daily_agg) // 15))
        )
        plt.xticks(rotation=45, ha="right")
        if len(daily_agg) > 0:
            first_val = daily_agg.iloc[0]["SPECTRAL_EFF"]
            ax.annotate(
                f"{first_val:.2f}",
                xy=(daily_agg.iloc[0]["DATE"], first_val),
                xytext=(10, 10),
                textcoords="offset points",
                fontsize=9,
                alpha=0.7,
            )
            last_val = daily_agg.iloc[-1]["SPECTRAL_EFF"]
            ax.annotate(
                f"{last_val:.2f}",
                xy=(daily_agg.iloc[-1]["DATE"], last_val),
                xytext=(10, -15),
                textcoords="offset points",
                fontsize=9,
                alpha=0.7,
            )

        plt.tight_layout()
        buf = BytesIO()
        plt.savefig(
            buf,
            format="png",
            dpi=100,
            bbox_inches="tight",
            facecolor="white",
            edgecolor="black",
        )
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode()
        plt.close(fig)

        return img_base64
