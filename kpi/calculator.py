"""KPI calculation module"""

import pandas as pd
import numpy as np
import json
from pathlib import Path
from config.settings import LTEColumns, GSMColumns, CONFIG_DIR


class KPICalculator:
    """Calculate KPIs for LTE and GSM data"""

    def __init__(self, transformed_data):
        self.lte_data = transformed_data["lte"]
        self.gsm_data = transformed_data["gsm"]
        self.ngi = transformed_data.get("ngi")

        # Load KPI config
        config_path = CONFIG_DIR / "kpi_config.json"
        with open(config_path, "r") as f:
            self.config = json.load(f)

        self.lte_kpis = None
        self.gsm_kpis = None

    def calculate_all(self):
        """Calculate all KPIs"""
        print("\n=== KPI Calculation ===")

        self._calculate_gsm_kpis()
        self._calculate_lte_kpis()

        return {
            "lte": self.lte_kpis,
            "gsm": self.gsm_kpis,
            "ngi": self.ngi
        }

    def _calculate_gsm_kpis(self):
        """Calculate GSM KPIs"""
        df = self.gsm_data.copy()

        # CSSR
        df["CSSR"] = np.where(
            df.iloc[:, GSMColumns.CSSR_DEN] > 0,
            (df.iloc[:, GSMColumns.CSSR_NUM] / df.iloc[:, GSMColumns.CSSR_DEN]) * 100,
            None,
        )

        # SDCCH Success Rate
        df["SDCCH_SR"] = np.where(
            df.iloc[:, GSMColumns.SDCCH_SR_DEN] > 0,
            (df.iloc[:, GSMColumns.SDCCH_SR_NUM] / df.iloc[:, GSMColumns.SDCCH_SR_DEN])
            * 100,
            None,
        )

        # Drop Rate
        df["DROP_RATE"] = np.where(
            df.iloc[:, GSMColumns.DROP_DEN] > 0,
            (df.iloc[:, GSMColumns.DROP_NUM] / df.iloc[:, GSMColumns.DROP_DEN]) * 100,
            None,
        )

        self.gsm_kpis = df
        print(f"✓ Calculated GSM KPIs: {len(df)} records")

    def _calculate_lte_kpis(self):
        """Calculate LTE KPIs"""
        df = self.lte_data.copy()

        rrc_ssr = np.where(
            df.iloc[:, LTEColumns.RRC_SSR_DEN] > 0,
            df.iloc[:, LTEColumns.RRC_SSR_NUM] / df.iloc[:, LTEColumns.RRC_SSR_DEN],
            0,
        )
        erab_ssr = np.where(
            df.iloc[:, LTEColumns.ERAB_SSR_DEN] > 0,
            df.iloc[:, LTEColumns.ERAB_SSR_NUM] / df.iloc[:, LTEColumns.ERAB_SSR_DEN],
            0,
        )
        s1_ssr = np.where(
            df.iloc[:, LTEColumns.S1_SSR_DEN] > 0,
            df.iloc[:, LTEColumns.S1_SSR_NUM] / df.iloc[:, LTEColumns.S1_SSR_DEN],
            0,
        )
        df["SESSION_SSR"] = rrc_ssr * erab_ssr * s1_ssr * 100
        df["SESSION_SSR"] = df["SESSION_SSR"].replace(0, None)

        # RACH Success Rate
        df["RACH_SR"] = np.where(
            df.iloc[:, LTEColumns.RACH_SETUP_DEN] > 0,
            (
                df.iloc[:, LTEColumns.RACH_SETUP_NUM]
                / df.iloc[:, LTEColumns.RACH_SETUP_DEN]
            )
            * 100,
            None,
        )

        # Handover Success Rate
        df["HO_SR"] = np.where(
            df.iloc[:, LTEColumns.HO_SR_DEN] > 0,
            (df.iloc[:, LTEColumns.HO_SR_NUM] / df.iloc[:, LTEColumns.HO_SR_DEN]) * 100,
            None,
        )

        # E-RAB Drop Rate
        df["ERAB_DROP"] = np.where(
            df.iloc[:, LTEColumns.ERAB_DROP_DEN] > 0,
            (
                df.iloc[:, LTEColumns.ERAB_DROP_NUM]
                / df.iloc[:, LTEColumns.ERAB_DROP_DEN]
            )
            * 100,
            None,
        )

        # Downlink Throughput
        df["DL_THP"] = np.where(
            df.iloc[:, LTEColumns.DL_THP_DEN] > 0,
            df.iloc[:, LTEColumns.DL_THP_NUM] / df.iloc[:, LTEColumns.DL_THP_DEN],
            None,
        )

        # Uplink Throughput
        df["UL_THP"] = np.where(
            df.iloc[:, LTEColumns.UL_THP_DEN] > 0,
            df.iloc[:, LTEColumns.UL_THP_NUM] / df.iloc[:, LTEColumns.UL_THP_DEN],
            None,
        )

        # Packet Loss
        # df['UL_PLOSS'] = df.iloc[:, LTEColumns.UL_PLOSS]
        # df['DL_PLOSS'] = df.iloc[:, LTEColumns.DL_PLOSS]
        df["UL_PLOSS"] = pd.to_numeric(df.iloc[:, LTEColumns.UL_PLOSS], errors="coerce")
        df["DL_PLOSS"] = pd.to_numeric(df.iloc[:, LTEColumns.DL_PLOSS], errors="coerce")

        # CQI
        df["CQI"] = np.where(
            df.iloc[:, LTEColumns.CQI_DEN] > 0,
            df.iloc[:, LTEColumns.CQI_NUM] / df.iloc[:, LTEColumns.CQI_DEN],
            None,
        )

        # MIMO Rank2 Rate
        df["MIMO_RANK2"] = np.where(
            df.iloc[:, LTEColumns.RANK_GT2_DEN] > 0,
            (df.iloc[:, LTEColumns.RANK_GT2_NUM] / df.iloc[:, LTEColumns.RANK_GT2_DEN])
            * 100,
            None,
        )

        # UL RSSI
        df["UL_RSSI"] = np.where(
            df.iloc[:, LTEColumns.RSSI_PUSCH_DEN] > 0,
            df.iloc[:, LTEColumns.RSSI_PUSCH_NUM]
            / df.iloc[:, LTEColumns.RSSI_PUSCH_DEN],
            None,
        )

        # Packet Latency
        df["LATENCY"] = np.where(
            df.iloc[:, LTEColumns.RAN_LAT_DEN] > 0,
            df.iloc[:, LTEColumns.RAN_LAT_NUM] / df.iloc[:, LTEColumns.RAN_LAT_DEN],
            None,
        )

        # LTC Non Capacity
        df["LTC_NON_CAP"] = df.iloc[:, LTEColumns.LTC_NON_CAP]

        # Overlap Rate
        df["OVERLAP_RATE"] = df.iloc[:, LTEColumns.OVERLAP_RATE]

        # Spectral Efficiency
        df["SPECTRAL_EFF"] = np.where(
            df.iloc[:, LTEColumns.DL_SE_DEN] > 0,
            df.iloc[:, LTEColumns.DL_SE_NUM] / df.iloc[:, LTEColumns.DL_SE_DEN],
            None,
        )

        # VoLTE CSSR
        df["VOLTE_CSSR"] = np.where(
            df.iloc[:, LTEColumns.VOLTE_CSSR_DEN] > 0,
            (
                df.iloc[:, LTEColumns.VOLTE_CSSR_NUM]
                / df.iloc[:, LTEColumns.VOLTE_CSSR_DEN]
            )
            * 100,
            None,
        )

        # VoLTE Drop
        df["VOLTE_DROP"] = np.where(
            df.iloc[:, LTEColumns.VOLTE_DROP_DEN] > 0,
            (
                df.iloc[:, LTEColumns.VOLTE_DROP_NUM]
                / df.iloc[:, LTEColumns.VOLTE_DROP_DEN]
            )
            * 100,
            None,
        )

        # SRVCC Success Rate
        df["SRVCC_SR"] = np.where(
            df.iloc[:, LTEColumns.SRVCC_SR_DEN] > 0,
            (df.iloc[:, LTEColumns.SRVCC_SR_NUM] / df.iloc[:, LTEColumns.SRVCC_SR_DEN])
            * 100,
            None,
        )

        self.lte_kpis = df
        print(f"✓ Calculated LTE KPIs: {len(df)} records")
