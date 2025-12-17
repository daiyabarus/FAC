"""KPI validation module"""

import pandas as pd
import numpy as np
import json
from config.settings import CONFIG_DIR


class KPIValidator:
    """Validate KPIs against baselines and calculate pass/fail"""

    def __init__(self, kpi_data):
        self.lte_kpis = kpi_data["lte"]
        self.gsm_kpis = kpi_data["gsm"]

        # Load KPI config
        config_path = CONFIG_DIR / "kpi_config.json"
        with open(config_path, "r") as f:
            self.config = json.load(f)

        self.monthly_results = {}

    def validate_all(self):
        """Validate all KPIs and calculate monthly achievements"""
        print("\n=== KPI Validation ===")

        # Get unique months
        lte_months = sorted(self.lte_kpis["MONTH"].unique())
        gsm_months = sorted(self.gsm_kpis["MONTH"].unique())
        all_months = sorted(list(set(lte_months + gsm_months)))

        # PERBAIKAN 2: Sort months properly by date (ascending - oldest first)
        month_dates = []
        for month in all_months:
            try:
                # Parse "Sep-25" format to datetime
                dt = pd.to_datetime(month, format="%b-%y")
                month_dates.append((month, dt))
            except:
                # Fallback if parsing fails
                month_dates.append((month, pd.Timestamp("1900-01-01")))

        # Sort by date (ascending - September, October, November)
        month_dates.sort(key=lambda x: x[1])
        all_months = [m[0] for m in month_dates]

        print(f"Processing months (sorted oldest to newest): {all_months}")

        # Get unique clusters
        clusters = self.lte_kpis["CLUSTER"].dropna().unique()

        results = {}

        for cluster in clusters:
            print(f"\nProcessing cluster: {cluster}")
            cluster_results = {}

            for month in all_months:
                month_results = self._validate_month(cluster, month)
                cluster_results[month] = month_results

            results[cluster] = cluster_results

        self.monthly_results = results
        print("\n✓ Validation complete")

        return results

    def _validate_month(self, cluster, month):
        """Validate KPIs for a specific cluster and month"""
        results = {}

        # Filter data
        lte_month = self.lte_kpis[
            (self.lte_kpis["CLUSTER"] == cluster) & (self.lte_kpis["MONTH"] == month)
        ]
        gsm_month = self.gsm_kpis[
            (self.gsm_kpis["CLUSTER"] == cluster) & (self.gsm_kpis["MONTH"] == month)
        ]

        # Validate GSM KPIs
        results["gsm"] = self._validate_gsm_kpis(gsm_month)

        # Validate LTE KPIs
        results["lte"] = self._validate_lte_kpis(lte_month)

        # Overall pass/fail
        all_results = list(results["gsm"].values()) + list(results["lte"].values())
        results["overall_pass"] = all(
            r.get("pass", False) for r in all_results if r is not None
        )

        return results

    def _validate_gsm_kpis(self, df):
        """Validate GSM KPIs"""
        if len(df) == 0:
            return {}

        results = {}

        # CSSR
        cssr_vals = df["CSSR"].dropna()
        if len(cssr_vals) > 0:
            pass_pct = (cssr_vals >= 98.5).sum() / len(cssr_vals) * 100
            results["cssr"] = {
                "value": pass_pct,
                "pass": pass_pct >= 95,
                "target": 95,
                "baseline": 98.5,
            }

        # SDCCH SR
        sdcch_vals = df["SDCCH_SR"].dropna()
        if len(sdcch_vals) > 0:
            pass_pct = (sdcch_vals >= 98.5).sum() / len(sdcch_vals) * 100
            results["sdcch_sr"] = {
                "value": pass_pct,
                "pass": pass_pct >= 95,
                "target": 95,
                "baseline": 98.5,
            }

        # Drop Rate
        drop_vals = df["DROP_RATE"].dropna()
        if len(drop_vals) > 0:
            pass_pct = (drop_vals < 2).sum() / len(drop_vals) * 100
            results["drop_rate"] = {
                "value": pass_pct,
                "pass": pass_pct >= 95,
                "target": 95,
                "baseline": 2,
            }

        return results

    def _validate_lte_kpis(self, df):
        """Validate LTE KPIs"""
        if len(df) == 0:
            return {}

        results = {}

        # Session SSR
        session_vals = df["SESSION_SSR"].dropna()
        if len(session_vals) > 0:
            pass_pct = (session_vals >= 99).sum() / len(session_vals) * 100
            results["session_ssr"] = {
                "value": pass_pct,
                "pass": pass_pct >= 97,
                "target": 97,
                "baseline": 99,
            }

        # RACH SR - High
        rach_vals = df["RACH_SR"].dropna()
        if len(rach_vals) > 0:
            pass_pct = (rach_vals >= 85).sum() / len(rach_vals) * 100
            results["rach_sr_high"] = {
                "value": pass_pct,
                "pass": pass_pct >= 60,
                "target": 60,
                "baseline": 85,
            }

        # RACH SR - Low
        if len(rach_vals) > 0:
            fail_pct = (rach_vals < 55).sum() / len(rach_vals) * 100
            results["rach_sr_low"] = {
                "value": fail_pct,
                "pass": fail_pct < 3,
                "target": 3,
                "baseline": 55,
            }

        # Handover SR
        ho_vals = df["HO_SR"].dropna()
        if len(ho_vals) > 0:
            pass_pct = (ho_vals >= 97).sum() / len(ho_vals) * 100
            results["ho_sr"] = {
                "value": pass_pct,
                "pass": pass_pct >= 95,
                "target": 95,
                "baseline": 97,
            }

        # E-RAB Drop
        erab_vals = df["ERAB_DROP"].dropna()
        if len(erab_vals) > 0:
            pass_pct = (erab_vals < 2).sum() / len(erab_vals) * 100
            results["erab_drop"] = {
                "value": pass_pct,
                "pass": pass_pct >= 95,
                "target": 95,
                "baseline": 2,
            }

        # DL Throughput - High
        dl_vals = df["DL_THP"].dropna()
        if len(dl_vals) > 0:
            pass_pct = (dl_vals >= 3).sum() / len(dl_vals) * 100
            results["dl_thp_high"] = {
                "value": pass_pct,
                "pass": pass_pct >= 85,
                "target": 85,
                "baseline": 3,
            }

        # DL Throughput - Low
        if len(dl_vals) > 0:
            fail_pct = (dl_vals < 1).sum() / len(dl_vals) * 100
            results["dl_thp_low"] = {
                "value": fail_pct,
                "pass": fail_pct < 2,
                "target": 2,
                "baseline": 1,
            }

        # UL Throughput - High
        ul_vals = df["UL_THP"].dropna()
        if len(ul_vals) > 0:
            pass_pct = (ul_vals >= 1).sum() / len(ul_vals) * 100
            results["ul_thp_high"] = {
                "value": pass_pct,
                "pass": pass_pct >= 65,
                "target": 65,
                "baseline": 1,
            }

        # UL Throughput - Low
        if len(ul_vals) > 0:
            fail_pct = (ul_vals < 0.256).sum() / len(ul_vals) * 100
            results["ul_thp_low"] = {
                "value": fail_pct,
                "pass": fail_pct < 2,
                "target": 2,
                "baseline": 0.256,
            }

        # PERBAIKAN 3: UL Packet Loss - INCLUDE 0 (0 is GOOD!)
        ul_ploss_vals = df["UL_PLOSS"].dropna()
        # TIDAK exclude 0 - karena 0 berarti packet loss = 0% (BAIK)
        if len(ul_ploss_vals) > 0:
            pass_pct = (ul_ploss_vals < 0.85).sum() / len(ul_ploss_vals) * 100
            results["ul_ploss"] = {
                "value": pass_pct,
                "pass": pass_pct >= 97,
                "target": 97,
                "baseline": 0.85,
            }

        # PERBAIKAN 3: DL Packet Loss - INCLUDE 0 (0 is GOOD!)
        dl_ploss_vals = df["DL_PLOSS"].dropna()
        # TIDAK exclude 0 - karena 0 berarti packet loss = 0% (BAIK)
        if len(dl_ploss_vals) > 0:
            pass_pct = (dl_ploss_vals < 0.10).sum() / len(dl_ploss_vals) * 100
            results["dl_ploss"] = {
                "value": pass_pct,
                "pass": pass_pct >= 97,
                "target": 97,
                "baseline": 0.10,
            }

        # CQI
        cqi_vals = df["CQI"].dropna()
        if len(cqi_vals) > 0:
            pass_pct = (cqi_vals >= 7).sum() / len(cqi_vals) * 100
            results["cqi"] = {
                "value": pass_pct,
                "pass": pass_pct >= 95,
                "target": 95,
                "baseline": 7,
            }

        # MIMO Rank2 - High
        mimo_vals = df["MIMO_RANK2"].dropna()
        if len(mimo_vals) > 0:
            pass_pct = (mimo_vals >= 35).sum() / len(mimo_vals) * 100
            results["mimo_rank2_high"] = {
                "value": pass_pct,
                "pass": pass_pct >= 70,
                "target": 70,
                "baseline": 35,
            }

        # MIMO Rank2 - Low
        if len(mimo_vals) > 0:
            fail_pct = (mimo_vals < 20).sum() / len(mimo_vals) * 100
            results["mimo_rank2_low"] = {
                "value": fail_pct,
                "pass": fail_pct < 5,
                "target": 5,
                "baseline": 20,
            }

        # UL RSSI
        rssi_vals = df["UL_RSSI"].dropna()
        if len(rssi_vals) > 0:
            pass_pct = (rssi_vals < -105).sum() / len(rssi_vals) * 100
            results["ul_rssi"] = {
                "value": pass_pct,
                "pass": pass_pct >= 97,
                "target": 97,
                "baseline": -105,
            }

        # Latency - Low
        lat_vals = df["LATENCY"].dropna()
        if len(lat_vals) > 0:
            pass_pct = (lat_vals < 30).sum() / len(lat_vals) * 100
            results["latency_low"] = {
                "value": pass_pct,
                "pass": pass_pct >= 95,
                "target": 95,
                "baseline": 30,
            }

        # Latency - Medium
        if len(lat_vals) > 0:
            between_pct = (
                ((lat_vals > 30) & (lat_vals < 40)).sum() / len(lat_vals) * 100
            )
            results["latency_medium"] = {
                "value": between_pct,
                "pass": between_pct < 5,
                "target": 5,
                "baseline": [30, 40],
            }

        # LTC Non Capacity
        ltc_vals = df["LTC_NON_CAP"].dropna()
        if len(ltc_vals) > 0:
            fail_pct = (ltc_vals < 3).sum() / len(ltc_vals) * 100
            results["ltc_non_cap"] = {
                "value": fail_pct,
                "pass": fail_pct < 5,
                "target": 5,
                "baseline": 3,
            }

        # Overlap Rate
        overlap_vals = df["OVERLAP_RATE"].dropna()
        if len(overlap_vals) > 0:
            pass_pct = (overlap_vals < 35).sum() / len(overlap_vals) * 100
            results["overlap_rate"] = {
                "value": pass_pct,
                # Pass jika >= 80% cells punya overlap < 35%
                "pass": pass_pct >= 80,
                "target": 80,
                "baseline": 35,
            }

        # Spectral Efficiency (multiple conditions)
        results.update(self._validate_spectral_efficiency(df))

        # VoLTE CSSR
        volte_cssr_vals = df["VOLTE_CSSR"].dropna()
        if len(volte_cssr_vals) > 0:
            pass_pct = (volte_cssr_vals > 97).sum() / len(volte_cssr_vals) * 100
            results["volte_cssr"] = {
                "value": pass_pct,
                "pass": pass_pct > 95,
                "target": 95,
                "baseline": 97,
            }

        # VoLTE Drop
        volte_drop_vals = df["VOLTE_DROP"].dropna()
        if len(volte_drop_vals) > 0:
            pass_pct = (volte_drop_vals < 2).sum() / len(volte_drop_vals) * 100
            results["volte_drop"] = {
                "value": pass_pct,
                "pass": pass_pct > 95,
                "target": 95,
                "baseline": 2,
            }

        # SRVCC SR
        srvcc_vals = df["SRVCC_SR"].dropna()
        if len(srvcc_vals) > 0:
            pass_pct = (srvcc_vals >= 97).sum() / len(srvcc_vals) * 100
            results["srvcc_sr"] = {
                "value": pass_pct,
                "pass": pass_pct > 95,
                "target": 95,
                "baseline": 97,
            }

        return results

    def _validate_spectral_efficiency(self, df):
        """Validate spectral efficiency with multiple conditions"""
        results = {}

        # FILTER: Only process cells with TX mapping
        df_with_tx = df[df["TX"].notna()].copy()

        if len(df_with_tx) == 0:
            print("  No cells with TX mapping for SE validation")
            return results

        # Different baselines based on TX and Band
        se_configs = [
            ("se_850_2t2r", "2T2R", 850, 1.1, 90, 36),
            ("se_900_2t2r", "2T2R", 900, 1.1, 90, 37),
            ("se_2100_2t2r", "2T2R", 2100, 1.3, 90, 38),
            ("se_1800_2t2r", "2T2R", 1800, 1.25, 90, 39),
            ("se_1800_4t4r", ["4T4R", "8T8R"], 1800, 1.5, 90, 40),
            ("se_2100_4t4r", ["4T4R", "8T8R"], 2100, 1.7, 90, 41),
            ("se_2300_32t32r", "32T32R", 2300, 2.1, 90, 43),
        ]

        for key, tx_cond, band_cond, baseline, target, row_num in se_configs:
            # Filter by TX
            if isinstance(tx_cond, list):
                mask_tx = df_with_tx["TX"].isin(tx_cond)
            else:
                mask_tx = df_with_tx["TX"] == tx_cond

            # Filter by Band
            if isinstance(band_cond, list):
                mask_band = df_with_tx["LTE_BAND"].isin(band_cond)
            else:
                mask_band = df_with_tx["LTE_BAND"] == band_cond

            # Apply filters
            filtered = df_with_tx[mask_tx & mask_band]
            se_vals = filtered["SPECTRAL_EFF"].dropna()

            # Only add results if data exists
            if len(se_vals) > 0:
                pass_pct = (se_vals >= baseline).sum() / len(se_vals) * 100
                results[key] = {
                    "value": pass_pct,
                    "pass": pass_pct >= 90,
                    "target": 90,
                    "baseline": baseline,
                    "tx": tx_cond,
                    "band": band_cond,
                    "row": row_num,
                    "cell_count": len(se_vals),
                }
                print(
                    f"  SE {key}: {len(se_vals)} cells, {pass_pct:.2f}% pass (baseline >= {baseline})"
                )
            else:
                print(f"  SE {key}: No data (TX={tx_cond}, Band={band_cond})")

        return results
