# fac_kpi_reporter.py
"""
FAC KPI Reporter
Single-file application implementing:
- CSV parsing (FDD, TDD, GSM, TOWERID)
- KPI evaluation per cell grouped by month & cluster
- Excel output per cluster (Summary + Cell contributors)
- Matplotlib charts (displayed and optionally embedded)
- PyQt6 GUI for selecting input/output and running

Author: ChatGPT (structured for extendability)
"""

from __future__ import annotations
import re
import sys
import os
import json
import math
import traceback
from dataclasses import dataclass
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# PyQt6 imports
from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QTextEdit,
    QProgressBar,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

# Excel writer engines
from pandas import ExcelWriter

# -----------------------------
# Domain models & constants
# -----------------------------

# Cell ID mapping from problem statement (truncated into dict)
CELL_ID_MAPPING: Dict[str, List[str]] = {
    "131": ["1", "850"],
    "132": ["2", "850"],
    "133": ["3", "850"],
    "134": ["4", "850"],
    "4": ["1", "1800"],
    "5": ["2", "1800"],
    "6": ["3", "1800"],
    "24": ["4", "1800"],
    "51": ["11", "1800"],
    "52": ["12", "1800"],
    "53": ["13", "1800"],
    "54": ["14", "1800"],
    "55": ["15", "1800"],
    "56": ["16", "1800"],
    "14": ["M1", "1800"],
    "15": ["M2", "1800"],
    "16": ["M3", "1800"],
    "64": ["M4", "1800"],
    "1": ["1", "2100"],
    "2": ["2", "2100"],
    "3": ["3", "2100"],
    "7": ["1", "2100"],
    "8": ["2", "2100"],
    "9": ["3", "2100"],
    "97": ["11", "2100"],
    "27": ["4", "2100"],
    "91": ["11", "2100"],
    "92": ["12", "2100"],
    "93": ["13", "2100"],
    "94": ["14", "2100"],
    "95": ["15", "2100"],
    "96": ["16", "2100"],
    "17": ["M1", "2100"],
    "18": ["M2", "2100"],
    "19": ["M3", "2100"],
    "67": ["M4", "2100"],
    "111": ["1", "2300"],
    "112": ["2", "2300"],
    "113": ["3", "2300"],
    "114": ["4", "2300"],
    "141": ["11", "2300"],
    "142": ["12", "2300"],
    "143": ["13", "2300"],
    "144": ["14", "2300"],
    "145": ["15", "2300"],
    "146": ["16", "2300"],
    "121": ["1", "2300"],
    "122": ["2", "2300"],
    "123": ["3", "2300"],
    "124": ["4", "2300"],
    "151": ["11", "2300"],
    "152": ["12", "2300"],
    "153": ["13", "2300"],
    "154": ["14", "2300"],
    "155": ["15", "2300"],
    "156": ["16", "2300"],
}

# Default MIMO by band for spectral efficiency
DEFAULT_MIMO_BY_BAND = {"850": "2T2R", "1800": "4T4R", "2100": "4T4R", "2300": "4T4R"}

# Regex to extract cluster key from element_name
CLUSTER_REGEX = re.compile(r"#([^#]+)#")

# Index mappings (0-based)
FDD_TDD_INDEX_MAP = {
    "begin_time": 0,
    "end_time": 1,
    "granularity": 2,
    "subnet_id": 3,
    "subnet_name": 4,
    "element_id": 5,
    "element_name": 6,
    "enodeb_cu_id": 7,
    "enodeb_cu_name": 8,
    "lte_id": 9,
    "lte_name": 10,
    "eutran_cell_id": 11,
    "eutran_cell_name": 12,
    "cell_id": 13,
    "enodeb_id": 14,
    # metrics (numerators/denominators) ...
    "sssr_num": 15,
    "sssr_den": 16,
    "rach_setup_sr_num": 17,
    "rach_setup_sr_den": 18,
    "ho_sr_num": 19,
    "ho_sr_den": 20,
    "erab_drop_num": 21,
    "erab_drop_den": 22,
    "dl_thp_num": 23,
    "dl_thp_den": 24,
    "ul_thp_num": 25,
    "ul_thp_den": 26,
    "ul_loss_num": 27,
    "dl_loss_num": 28,
    "cqi_num": 29,
    "cqi_den": 30,
    "rank_gt2_num": 31,
    "rank_gt2_den": 32,
    "rssi_avg_num": 33,
    "rssi_avg_den": 34,
    "ran_latency_num": 35,
    "ran_latency_den": 36,
    "dl_se_num": 37,
    "dl_se_den": 38,
    "volte_setup_num": 39,
    "volte_setup_den": 40,
    "volte_drop_num": 41,
    "volte_drop_den": 42,
    "srvcc_success_num": 43,
    "srvcc_success_den": 44,
}

GSM_INDEX_MAP = {
    "begin_time": 0,
    "end_time": 1,
    "granularity": 2,
    "subnet_id": 3,
    "subnet_name": 4,
    "element_id": 5,
    "element_name": 6,
    "site_id": 7,
    "site_name": 8,
    "bts_id": 9,
    "bts_name": 10,
    "freq_band": 11,
    "call_setup_sr_num": 12,
    "call_setup_sr_den": 13,
    "sdcch_sr_num": 14,
    "sdcch_sr_den": 15,
    "drop_rate_num": 16,
    "drop_rate_den": 17,
}

TOWERID_MAP_INDICES = {"cluster": 0, "towerid": 1, "site_name": 2}


# -----------------------------
# Utility & parsing functions
# -----------------------------


def cleanse_value(v: Any) -> Any:
    """
    Remove thousands separators and percent signs from numeric-like strings.
    Return numeric types (int/float) where possible, otherwise return original stripped string.
    """
    if pd.isna(v):
        return v
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if s == "":
        return np.nan
    # Remove commas used as thousand separators
    s = s.replace(",", "")
    # Remove percent
    s = s.replace("%", "")
    # convert to numeric if possible
    try:
        if "." in s or "e" in s.lower():
            return float(s)
        else:
            return int(s)
    except Exception:
        # if can't convert, return original cleaned string
        return s


def parse_csv_file(path: str) -> pd.DataFrame:
    """
    Read CSV robustly with pandas, do minimal cleansing.
    """
    df = pd.read_csv(path, header=None, dtype=str, low_memory=False)
    # Cleanse every cell
    df = df.applymap(cleanse_value)
    return df


def extract_cluster_from_element_name(element_name: str) -> Optional[str]:
    if not isinstance(element_name, str):
        return None
    m = CLUSTER_REGEX.search(element_name)
    if m:
        return m.group(1)
    return None


def extract_band_sector_from_cell_id(
    cell_id_value: Any,
) -> Tuple[Optional[str], Optional[str]]:
    # cell_id might be numeric or string; safely convert to string
    if pd.isna(cell_id_value):
        return None, None
    s = str(cell_id_value).strip()
    # often the mapping is found by the last 3 chars or so; we will search for any mapping key as suffix/prefix
    # simplest: try exact match, then try last 3/2/1 characters
    if s in CELL_ID_MAPPING:
        return CELL_ID_MAPPING[s][0], CELL_ID_MAPPING[s][1]
    for l in (3, 2, 1):
        if len(s) >= l:
            k = s[-l:]
            if k in CELL_ID_MAPPING:
                return CELL_ID_MAPPING[k][0], CELL_ID_MAPPING[k][1]
    # not found
    return None, None


def month_label_from_time(ts: Any) -> Optional[str]:
    if pd.isna(ts):
        return None
    s = str(ts)
    # try parse common formats
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y-%m")  # group by year-month
        except Exception:
            continue
    # fallback: try using first 7 chars YYYY-MM
    if len(s) >= 7 and s[4] == "-":
        return s[:7]
    return None


# -----------------------------
# KPI classes & engine
# -----------------------------


@dataclass
class KPIResult:
    cell_key: str
    cluster: str
    site: str
    band: Optional[str]
    month: str
    kpi_name: str
    kpi_value: Optional[float]
    meets: Optional[bool]
    raw: Dict[str, Any]


class KPI:
    """
    Base KPI class. Subclass to implement kpi_value(row) and threshold_evaluator(value) methods.
    """

    name: str
    description: str
    unit: str

    def __init__(self, name: str, description: str, unit: str):
        self.name = name
        self.description = description
        self.unit = unit

    def compute(self, row: Dict[str, Any]) -> Optional[float]:
        """
        Compute KPI metric from a row dictionary (common schema). Return numeric or None.
        """
        raise NotImplementedError

    def evaluate(self, value: Optional[float]) -> Optional[bool]:
        """
        Return True if meets KPI target, False if not, None if unknown.
        """
        raise NotImplementedError


# Implement some KPIs using the numerator/denominator patterns
class RatioKPI(KPI):
    def __init__(
        self,
        name: str,
        description: str,
        unit: str,
        num_field: str,
        den_field: str,
        compare: str,
        threshold: float,
    ):
        super().__init__(name, description, unit)
        self.num_field = num_field
        self.den_field = den_field
        self.compare = compare
        self.threshold = threshold

    def compute(self, row: Dict[str, Any]) -> Optional[float]:
        n = row.get(self.num_field)
        d = row.get(self.den_field)
        try:
            if d is None or d == 0 or pd.isna(d):
                return None
            return float(n) / float(d) * (100.0 if self.unit == "%" else 1.0)
        except Exception:
            return None

    def evaluate(self, value: Optional[float]) -> Optional[bool]:
        if value is None:
            return None
        if self.compare == ">=":
            return value >= self.threshold
        elif self.compare == "<":
            return value < self.threshold
        elif self.compare == "<=":
            return value <= self.threshold
        elif self.compare == ">":
            return value > self.threshold
        else:
            return None


# Specific KPIs


# 2G Call Setup Success Rate
class CallSetup2GKPI(RatioKPI):
    pass


# SDCCH
class SDCCHKPI(RatioKPI):
    pass


# Perceive Drop Rate (call drop)
class PerceiveDropKPI(RatioKPI):
    pass


# 4G Session Setup Success Rate (SSSR)
class SessionSetupKPI(RatioKPI):
    pass


# RACH Success Rate
class RACHKPI(RatioKPI):
    pass


# E-RAB drop rate
class ERABDropKPI(RatioKPI):
    pass


# DL/UL Throughput KPIs: use numerator/denominator if available else use a throughput field
class ThroughputKPI(KPI):
    def __init__(
        self,
        name: str,
        description: str,
        unit: str,
        num_field: Optional[str],
        den_field: Optional[str],
        threshold: float,
        comparator: str,
        pct_target: Optional[float] = None,
    ):
        super().__init__(name, description, unit)
        self.num_field = num_field
        self.den_field = den_field
        self.threshold = threshold
        self.comparator = comparator
        self.pct_target = pct_target

    def compute(self, row: Dict[str, Any]) -> Optional[float]:
        # prefer direct throughput numerator/denominator to compute average throughput: num/den
        n = row.get(self.num_field) if self.num_field else None
        d = row.get(self.den_field) if self.den_field else None
        try:
            if n is None or d is None or pd.isna(n) or pd.isna(d):
                # attempt use dl_thp_num directly as Mbps if present
                if self.num_field and row.get(self.num_field) is not None:
                    return float(row.get(self.num_field))
                return None
            if float(d) == 0:
                return None
            return float(n) / float(d)  # assumes units match threshold
        except Exception:
            return None

    def evaluate(self, value: Optional[float]) -> Optional[bool]:
        if value is None:
            return None
        if self.comparator == ">=":
            return value >= self.threshold
        if self.comparator == "<":
            return value < self.threshold
        return None


# CQI KPI uses cqi_num/cqi_den to compute average
class CQIKPI(RatioKPI):
    def compute(self, row: Dict[str, Any]) -> Optional[float]:
        n = row.get("cqi_num")
        d = row.get("cqi_den")
        try:
            if d is None or d == 0 or pd.isna(d):
                return None
            return float(n) / float(d)
        except Exception:
            return None


# UL RSSI KPI (value is dBm)
class ULRSSIKPI(KPI):
    def __init__(
        self,
        name: str,
        description: str,
        unit: str,
        field: str,
        threshold: float,
        compare: str,
    ):
        super().__init__(name, description, unit)
        self.field = field
        self.threshold = threshold
        self.compare = compare

    def compute(self, row: Dict[str, Any]) -> Optional[float]:
        v = row.get(self.field)
        try:
            if pd.isna(v):
                return None
            return float(v)
        except Exception:
            return None

    def evaluate(self, value: Optional[float]) -> Optional[bool]:
        if value is None:
            return None
        if self.compare == "<":
            return value < self.threshold
        if self.compare == ">=":
            return value >= self.threshold
        return None


# Packet latency KPI
class LatencyKPI(KPI):
    def __init__(
        self,
        name: str,
        description: str,
        unit: str,
        field_num: str,
        field_den: str,
        threshold: float,
        compare: str,
    ):
        super().__init__(name, description, unit)
        self.field_num = field_num
        self.field_den = field_den
        self.threshold = threshold
        self.compare = compare

    def compute(self, row: Dict[str, Any]) -> Optional[float]:
        n = row.get(self.field_num)
        d = row.get(self.field_den)
        try:
            if d is None or d == 0 or pd.isna(d):
                return None
            return float(n) / float(d)
        except Exception:
            return None

    def evaluate(self, value: Optional[float]) -> Optional[bool]:
        if value is None:
            return None
        if self.compare == "<":
            return value < self.threshold
        if self.compare == ">=":
            return value >= self.threshold
        return None


# Spectral Efficiency KPI: uses dl_se_num / dl_se_den or estimated mapping + default MIMO thresholds
class SpectralEfficiencyKPI(KPI):
    """
    Compute spectral efficiency if dl_se_num/dl_se_den available,
    otherwise use a band+MIMO default target mapping.
    """

    # threshold map (band + mimo -> target)
    THRESHOLD_MAP = {
        ("850", "2T2R"): 1.1,
        ("900", "2T2R"): 1.1,
        ("2100", "2T2R"): 1.3,
        ("1800", "2T2R"): 1.25,
        ("1800", "4T4R"): 1.5,
        ("1800", "8T8R"): 1.5,
        ("2100", "4T4R"): 1.7,
        ("2300", "4T4R"): 1.7,
        ("1800", "MM"): 1.25,
        ("2300", "MM"): 2.1,
    }

    def compute(self, row: Dict[str, Any]) -> Optional[float]:
        n = row.get("dl_se_num")
        d = row.get("dl_se_den")
        try:
            if n is not None and d is not None and not pd.isna(d) and float(d) != 0:
                return float(n) / float(d)
        except Exception:
            pass
        # otherwise can't compute actual SE from data; return None (we'll use threshold evaluator that may infer)
        return None

    def evaluate(
        self, value: Optional[float], band: Optional[str], mimo: Optional[str]
    ) -> Optional[bool]:
        # If we computed actual SE, compare to threshold based on band+MIMO mapping
        target = None
        if band is None:
            return None
        if mimo is None:
            mimo = DEFAULT_MIMO_BY_BAND.get(band, None)
        key = (str(band), mimo)
        target = self.THRESHOLD_MAP.get(key)
        if value is not None and target is not None:
            return value >= target
        # If we don't have computed value, we consider unknown (or could compute using other proxies)
        return None


# -----------------------------
# Application / Use-case layer
# -----------------------------


class DataLoader:
    """
    Reads CSV files (FDD, TDD, GSM, TOWERID) from a folder and produces a unified DataFrame
    with normalized columns.
    """

    def __init__(self, input_folder: str):
        self.input_folder = input_folder
        self.fdd_dfs: List[pd.DataFrame] = []
        self.tdd_dfs: List[pd.DataFrame] = []
        self.gsm_dfs: List[pd.DataFrame] = []
        self.tower_df: Optional[pd.DataFrame] = None

    def discover_files(self) -> Dict[str, List[str]]:
        files = os.listdir(self.input_folder)
        result = {"fdd": [], "tdd": [], "gsm": [], "tower": []}
        for f in files:
            lp = f.lower()
            if lp.endswith(".csv"):
                if "fdd" in lp:
                    result["fdd"].append(os.path.join(self.input_folder, f))
                elif "tdd" in lp:
                    result["tdd"].append(os.path.join(self.input_folder, f))
                elif "gsm" in lp:
                    result["gsm"].append(os.path.join(self.input_folder, f))
                elif "tower" in lp or "towerid" in lp:
                    result["tower"].append(os.path.join(self.input_folder, f))
                else:
                    # try to inspect header heuristically
                    result["fdd"].append(os.path.join(self.input_folder, f))
        return result

    def load(self) -> Dict[str, pd.DataFrame]:
        files = self.discover_files()
        # Load towerid first
        if files["tower"]:
            # take first
            tpath = files["tower"][0]
            self.tower_df = parse_csv_file(tpath)
            self.tower_df.columns = [f"c{idx}" for idx in range(self.tower_df.shape[1])]
        else:
            self.tower_df = pd.DataFrame()
        # load FDD/TDD/GSM
        fdd_concat = []
        for path in files["fdd"]:
            df = parse_csv_file(path)
            df.columns = [f"c{idx}" for idx in range(df.shape[1])]
            fdd_concat.append(df)
        tdd_concat = []
        for path in files["tdd"]:
            df = parse_csv_file(path)
            df.columns = [f"c{idx}" for idx in range(df.shape[1])]
            tdd_concat.append(df)
        gsm_concat = []
        for path in files["gsm"]:
            df = parse_csv_file(path)
            df.columns = [f"c{idx}" for idx in range(df.shape[1])]
            gsm_concat.append(df)
        fdd_df = (
            pd.concat(fdd_concat, ignore_index=True) if fdd_concat else pd.DataFrame()
        )
        tdd_df = (
            pd.concat(tdd_concat, ignore_index=True) if tdd_concat else pd.DataFrame()
        )
        gsm_df = (
            pd.concat(gsm_concat, ignore_index=True) if gsm_concat else pd.DataFrame()
        )

        return {"fdd": fdd_df, "tdd": tdd_df, "gsm": gsm_df, "tower": self.tower_df}


class Normalizer:
    """
    Normalize raw CSV rows into a common schema (dictionary per row) for KPI engine.
    """

    def __init__(self, tower_df: pd.DataFrame):
        # tower_df uses raw columns c0, c1, ...; map by TOWERID_MAP_INDICES
        self.tower_map: Dict[str, str] = {}
        if tower_df is not None and not tower_df.empty:
            for _, row in tower_df.iterrows():
                try:
                    site_name = row.get(f"c{TOWERID_MAP_INDICES['site_name']}")
                    cluster = row.get(f"c{TOWERID_MAP_INDICES['cluster']}")
                    if site_name is not None:
                        self.tower_map[str(site_name).strip()] = (
                            str(cluster).strip() if cluster is not None else ""
                        )
                except Exception:
                    pass

    def normalize_fdd_tdd_row(self, row: pd.Series) -> Dict[str, Any]:
        # convert row with columns c0.. to fields using FDD_TDD_INDEX_MAP
        d: Dict[str, Any] = {}
        for field, idx in FDD_TDD_INDEX_MAP.items():
            key = f"c{idx}"
            d[field] = row.get(key)
        # additional computed fields
        d["technology"] = "4G"
        # cluster extraction
        d["cluster"] = (
            extract_cluster_from_element_name(d.get("element_name", "")) or ""
        )
        # site derive (element_name or enodeb_cu_name)
        d["site_name"] = d.get("enodeb_cu_name") or d.get("element_name")
        d["cell_key"] = f"{d.get('element_id') or ''}_{d.get('eutran_cell_id') or ''}"
        # month grouping
        d["month"] = (
            month_label_from_time(d.get("begin_time") or d.get("end_time")) or "unknown"
        )
        # band/sector
        sector, band = extract_band_sector_from_cell_id(d.get("cell_id"))
        d["band"] = band
        d["sector"] = sector
        # keep raw row
        d["raw_row"] = row.to_dict()
        return d

    def normalize_gsm_row(self, row: pd.Series) -> Dict[str, Any]:
        d: Dict[str, Any] = {}
        for field, idx in GSM_INDEX_MAP.items():
            key = f"c{idx}"
            d[field] = row.get(key)
        d["technology"] = "2G"
        # map cluster using site_name via tower map
        site = d.get("site_name")
        d["cluster"] = self.tower_map.get(str(site).strip(), "")
        d["cell_key"] = f"{d.get('site_id') or ''}_{d.get('bts_id') or ''}"
        d["month"] = (
            month_label_from_time(d.get("begin_time") or d.get("end_time")) or "unknown"
        )
        d["band"] = d.get("freq_band")
        d["sector"] = None
        d["raw_row"] = row.to_dict()
        return d


class KPIEngine:
    """
    Given normalized rows, compute KPIs and produce KPIResults aggregated.
    """

    def __init__(self):
        # register KPIs here
        self.kpis: List[KPI] = []
        # 2G KPIs
        self.kpis.append(
            CallSetup2GKPI(
                "2G Call Setup Success Rate",
                "Call setup success",
                "%",
                "call_setup_sr_num",
                "call_setup_sr_den",
                ">=",
                98.5,
            )
        )
        self.kpis.append(
            SDCCHKPI(
                "2G SDCCH Success Rate",
                "SDCCH success",
                "%",
                "sdcch_sr_num",
                "sdcch_sr_den",
                ">=",
                98.5,
            )
        )
        self.kpis.append(
            PerceiveDropKPI(
                "2G Perceive Drop Rate",
                "Drop rate",
                "%",
                "drop_rate_num",
                "drop_rate_den",
                "<",
                2.0,
            )
        )
        # 4G KPIs (fields use FDD_TDD_INDEX_MAP names)
        self.kpis.append(
            SessionSetupKPI(
                "4G Session Setup Success Rate",
                "Session setup",
                "%",
                "sssr_num",
                "sssr_den",
                ">=",
                99.0,
            )
        )
        self.kpis.append(
            RACHKPI(
                "4G RACH Success Rate",
                "RACH success",
                "%",
                "rach_setup_sr_num",
                "rach_setup_sr_den",
                ">=",
                85.0,
            )
        )
        self.kpis.append(
            ERABDropKPI(
                "4G E-RAB Drop Rate",
                "E-RAB drop",
                "%",
                "erab_drop_num",
                "erab_drop_den",
                "<",
                2.0,
            )
        )
        # Throughput KPIs (DL >=3 Mbps; DL <1 Mbps; UL >=1 Mbps; UL <0.256 Mbps)
        self.kpis.append(
            ThroughputKPI(
                "DL Throughput >= 3Mbps",
                "Downlink throughput",
                "Mbps",
                "dl_thp_num",
                "dl_thp_den",
                3.0,
                ">=",
            )
        )
        self.kpis.append(
            ThroughputKPI(
                "DL Throughput < 1Mbps",
                "Downlink low throughput",
                "Mbps",
                "dl_thp_num",
                "dl_thp_den",
                1.0,
                "<",
            )
        )
        self.kpis.append(
            ThroughputKPI(
                "UL Throughput >= 1Mbps",
                "Uplink throughput",
                "Mbps",
                "ul_thp_num",
                "ul_thp_den",
                1.0,
                ">=",
            )
        )
        self.kpis.append(
            ThroughputKPI(
                "UL Throughput < 0.256Mbps",
                "Uplink low throughput",
                "Mbps",
                "ul_thp_num",
                "ul_thp_den",
                0.256,
                "<",
            )
        )
        # UL/DL packet loss PDCP: UL loss field is ul_loss_num? We only have ul_loss_num/dl_loss_num in mapping (denominators missing)
        self.kpis.append(
            RatioKPI(
                "4G UL Packet Loss (PDCP)",
                "UL PDCP loss",
                "%",
                "ul_loss_num",
                "sssr_den",
                "<",
                0.85,
            )
        )  # fallback denominator chosen; adjust as needed
        self.kpis.append(
            RatioKPI(
                "4G DL Packet Loss (PDCP)",
                "DL PDCP loss",
                "%",
                "dl_loss_num",
                "sssr_den",
                "<",
                0.10,
            )
        )
        self.kpis.append(CQIKPI("4G CQI", "CQI", "num", "cqi_num", "cqi_den", ">=", 7))
        self.kpis.append(
            RatioKPI(
                "4G MIMO Rank2 Rate >=35%",
                "MIMO rank>=2 pct",
                "%",
                "rank_gt2_num",
                "rank_gt2_den",
                ">=",
                35.0,
            )
        )
        self.kpis.append(
            RatioKPI(
                "4G MIMO Rank2 Rate <20%",
                "MIMO low rank pct",
                "%",
                "rank_gt2_num",
                "rank_gt2_den",
                "<",
                20.0,
            )
        )
        # UL RSSI
        self.kpis.append(
            ULRSSIKPI("4G UL RSSI", "UL RSSI dBm", "dBm", "rssi_avg_num", -105.0, "<")
        )
        # Packet latency
        self.kpis.append(
            LatencyKPI(
                "4G Packet Latency <30ms",
                "Latency ms",
                "ms",
                "ran_latency_num",
                "ran_latency_den",
                30.0,
                "<",
            )
        )
        self.kpis.append(
            LatencyKPI(
                "4G Packet Latency >=40ms",
                "Latency ms",
                "ms",
                "ran_latency_num",
                "ran_latency_den",
                40.0,
                ">=",
            )
        )
        # Spectral Efficiency KPI
        self.spectral_kpi = SpectralEfficiencyKPI("Spectral Efficiency", "SE", "num")
        self.kpis.append(self.spectral_kpi)
        # VoLTE (placeholders)
        self.kpis.append(
            RatioKPI(
                "VoLTE Voice Call Success Rate",
                "VoLTE call setup",
                "%",
                "volte_setup_num",
                "volte_setup_den",
                ">=",
                97.0,
            )
        )
        self.kpis.append(
            RatioKPI(
                "VoLTE Voice Call Drop Rate",
                "VoLTE drop",
                "%",
                "volte_drop_num",
                "volte_drop_den",
                "<",
                2.0,
            )
        )
        self.kpis.append(
            RatioKPI(
                "VoLTE SRVCC Success Rate",
                "SRVCC",
                "%",
                "srvcc_success_num",
                "srvcc_success_den",
                ">=",
                97.0,
            )
        )

    def compute_for_rows(
        self, normalized_rows: List[Dict[str, Any]]
    ) -> List[KPIResult]:
        results: List[KPIResult] = []
        for row in normalized_rows:
            for kpi in self.kpis:
                try:
                    if isinstance(kpi, SpectralEfficiencyKPI):
                        val = kpi.compute(row)
                        # attempt evaluate with band/mimo
                        mimo = None
                        band = row.get("band")
                        if band:
                            mimo = DEFAULT_MIMO_BY_BAND.get(str(band))
                        meets = kpi.evaluate(val, band, mimo)
                        results.append(
                            KPIResult(
                                cell_key=row.get("cell_key"),
                                cluster=row.get("cluster") or "",
                                site=row.get("site_name") or "",
                                band=row.get("band"),
                                month=row.get("month") or "unknown",
                                kpi_name=kpi.name,
                                kpi_value=val,
                                meets=meets,
                                raw=row,
                            )
                        )
                    else:
                        val = kpi.compute(row)
                        meets = kpi.evaluate(val)
                        results.append(
                            KPIResult(
                                cell_key=row.get("cell_key"),
                                cluster=row.get("cluster") or "",
                                site=row.get("site_name") or "",
                                band=row.get("band"),
                                month=row.get("month") or "unknown",
                                kpi_name=kpi.name,
                                kpi_value=val,
                                meets=meets,
                                raw=row,
                            )
                        )
                except Exception:
                    # don't fail entire process for single KPI
                    results.append(
                        KPIResult(
                            cell_key=row.get("cell_key"),
                            cluster=row.get("cluster") or "",
                            site=row.get("site_name") or "",
                            band=row.get("band"),
                            month=row.get("month") or "unknown",
                            kpi_name=kpi.name,
                            kpi_value=None,
                            meets=None,
                            raw=row,
                        )
                    )
        return results


# -----------------------------
# Reporting / Excel generation
# -----------------------------


class ExcelReporter:
    """
    Builds Excel workbooks per cluster with Summary and Cell Contributors.
    """

    def __init__(self, output_path: str):
        self.output_path = output_path

    def report(self, kpi_results: List[KPIResult], output_file: str):
        # Group by cluster
        if not kpi_results:
            raise ValueError("No KPI results to write")
        df = pd.DataFrame(
            [
                {
                    "cell_key": r.cell_key,
                    "cluster": r.cluster,
                    "site": r.site,
                    "band": r.band,
                    "month": r.month,
                    "kpi_name": r.kpi_name,
                    "kpi_value": r.kpi_value,
                    "meets": r.meets,
                }
                for r in kpi_results
            ]
        )
        clusters = df["cluster"].fillna("").unique().tolist()
        # If no cluster, one workbook
        wb_path = os.path.join(self.output_path, output_file)
        with pd.ExcelWriter(
            wb_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd"
        ) as writer:
            # Global summary sheet build: pivot to get percentage pass per KPI per month
            summary_rows = []
            for cluster in clusters:
                dfc = df[df["cluster"] == cluster]
                months = sorted(dfc["month"].fillna("unknown").unique())
                kpi_names = sorted(dfc["kpi_name"].unique())
                for k in kpi_names:
                    row = {"Cluster": cluster, "KPI": k}
                    for m in months:
                        sub = dfc[(dfc["kpi_name"] == k) & (dfc["month"] == m)]
                        total = len(sub)
                        passed = len(sub[sub["meets"] == True])
                        if total == 0:
                            pct = None
                        else:
                            pct = passed / total * 100.0
                        row[m] = pct
                    summary_rows.append(row)
            summary_df = pd.DataFrame(summary_rows)
            summary_df.to_excel(
                writer, sheet_name="FAC KPI Achievement Summary", index=False
            )
            # Cell contributors: list failing cells (meets == False)
            fail_df = df[df["meets"] == False].copy()
            if not fail_df.empty:
                fail_df = fail_df.sort_values(["cluster", "kpi_name", "month"])
                fail_df.to_excel(writer, sheet_name="Cell Contributors", index=False)
            else:
                pd.DataFrame(
                    [], columns=["cluster", "kpi_name", "cell_key", "month"]
                ).to_excel(writer, sheet_name="Cell Contributors", index=False)

            # optionally write a raw sheet
            df.to_excel(writer, sheet_name="Raw KPI Results", index=False)
        return wb_path


# -----------------------------
# GUI & Worker Thread
# -----------------------------


class WorkerThread(QThread):
    progress = pyqtSignal(int)
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)

    def __init__(self, input_folder: str, output_folder: str, output_filename: str):
        super().__init__()
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.output_filename = output_filename

    def run(self):
        try:
            # Step 1: load data
            self.progress.emit(5)
            loader = DataLoader(self.input_folder)
            raw = loader.load()
            self.progress.emit(15)
            normalizer = Normalizer(raw.get("tower"))
            # build normalized rows for all technologies
            normalized_rows: List[Dict[str, Any]] = []
            # FDD
            fdd = raw.get("fdd")
            if fdd is not None and not fdd.empty:
                for _, r in fdd.iterrows():
                    normalized_rows.append(normalizer.normalize_fdd_tdd_row(r))
            # TDD
            tdd = raw.get("tdd")
            if tdd is not None and not tdd.empty:
                for _, r in tdd.iterrows():
                    normalized_rows.append(normalizer.normalize_fdd_tdd_row(r))
            # GSM
            gsm = raw.get("gsm")
            if gsm is not None and not gsm.empty:
                for _, r in gsm.iterrows():
                    normalized_rows.append(normalizer.normalize_gsm_row(r))
            self.progress.emit(40)
            # KPI compute
            engine = KPIEngine()
            kpi_results = engine.compute_for_rows(normalized_rows)
            self.progress.emit(80)
            # Excel report
            reporter = ExcelReporter(self.output_folder)
            outpath = reporter.report(kpi_results, self.output_filename)
            self.progress.emit(95)
            # make charts (simple example: KPI achievement per month for top KPIs)
            try:
                self.make_charts(kpi_results)
                self.progress.emit(100)
            except Exception:
                pass
            self.finished_signal.emit(outpath)
        except Exception as ex:
            tb = traceback.format_exc()
            self.error_signal.emit(f"Error: {ex}\n{tb}")

    def make_charts(self, kpi_results: List[KPIResult]):
        if not kpi_results:
            return
        df = pd.DataFrame(
            [
                {
                    "kpi": r.kpi_name,
                    "month": r.month,
                    "meets": (1 if r.meets else 0),
                    "cluster": r.cluster,
                }
                for r in kpi_results
            ]
        )
        # select a few KPIs for charting
        top_kpis = df["kpi"].value_counts().index.tolist()[:6]
        chart_df = df[df["kpi"].isin(top_kpis)]
        # pivot percent pass
        pivot = (
            chart_df.groupby(["kpi", "month"])["meets"]
            .mean()
            .unstack(level=1)
            .fillna(0)
            * 100
        )
        pivot.plot(kind="bar", rot=45)
        plt.ylabel("Percent Cells Meeting KPI (%)")
        plt.title("KPI Achievement by Month")
        plt.tight_layout()
        plt.show()


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FAC KPI Reporter")
        self.resize(800, 400)
        self.setup_ui()
        self.worker: Optional[WorkerThread] = None

    def setup_ui(self):
        layout = QVBoxLayout()

        # Input folder
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Input folder (CSV files):"))
        self.input_edit = QLineEdit()
        h1.addWidget(self.input_edit)
        btn_in = QPushButton("Browse")
        btn_in.clicked.connect(self.browse_input)
        h1.addWidget(btn_in)
        layout.addLayout(h1)

        # Output folder + filename
        h2 = QHBoxLayout()
        h2.addWidget(QLabel("Output folder:"))
        self.output_edit = QLineEdit()
        h2.addWidget(self.output_edit)
        btn_out = QPushButton("Browse")
        btn_out.clicked.connect(self.browse_output)
        h2.addWidget(btn_out)
        layout.addLayout(h2)

        h3 = QHBoxLayout()
        h3.addWidget(QLabel("Output filename (e.g. FAC_report.xlsx):"))
        self.filename_edit = QLineEdit("FAC_report.xlsx")
        h3.addWidget(self.filename_edit)
        layout.addLayout(h3)

        # Buttons
        btn_layout = QHBoxLayout()
        self.run_btn = QPushButton("Generate Report")
        self.run_btn.clicked.connect(self.run_report)
        btn_layout.addWidget(self.run_btn)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.cancel_run)
        btn_layout.addWidget(self.cancel_btn)
        layout.addLayout(btn_layout)

        # Progress and log
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        layout.addWidget(self.progress)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        self.setLayout(layout)

    def browse_input(self):
        dirpath = QFileDialog.getExistingDirectory(self, "Select input folder")
        if dirpath:
            self.input_edit.setText(dirpath)

    def browse_output(self):
        dirpath = QFileDialog.getExistingDirectory(self, "Select output folder")
        if dirpath:
            self.output_edit.setText(dirpath)

    def log_message(self, s: str):
        self.log.append(s)

    def run_report(self):
        input_folder = self.input_edit.text().strip()
        output_folder = self.output_edit.text().strip()
        filename = self.filename_edit.text().strip()
        if not input_folder or not os.path.isdir(input_folder):
            QMessageBox.warning(
                self, "Input required", "Please select a valid input folder."
            )
            return
        if not output_folder or not os.path.isdir(output_folder):
            QMessageBox.warning(
                self, "Output required", "Please select a valid output folder."
            )
            return
        if not filename.endswith(".xlsx"):
            QMessageBox.warning(
                self, "Filename", "Output filename should end with .xlsx"
            )
            return
        self.run_btn.setEnabled(False)
        self.worker = WorkerThread(input_folder, output_folder, filename)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.error_signal.connect(self.on_error)
        self.worker.start()
        self.log_message("Started processing...")

    def cancel_run(self):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.log_message("Worker terminated by user.")
            self.run_btn.setEnabled(True)
        else:
            self.log_message("No running process to cancel.")

    def on_finished(self, outpath: str):
        self.log_message(f"Report generated: {outpath}")
        QMessageBox.information(self, "Done", f"Report written to:\n{outpath}")
        self.run_btn.setEnabled(True)

    def on_error(self, error: str):
        self.log_message(error)
        QMessageBox.critical(self, "Error", error)
        self.run_btn.setEnabled(True)


# -----------------------------
# Main entry
# -----------------------------
def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
