"""
Domain Layer - Core Business Models
File: domain/models.py
"""

from dataclasses import dataclass
from typing import Optional
from datetime import datetime
from enum import Enum


class Technology(Enum):
    """Technology types for cells."""
    GSM = "2G"
    LTE_FDD = "4G_FDD"
    LTE_TDD = "4G_TDD"
    NR = "5G"


class KPIDomain(Enum):
    """KPI Domain categories."""
    ACCESSIBILITY = "Accessibility"
    RETAINABILITY = "Retainability"
    INTEGRITY = "Integrity"
    MOBILITY = "Mobility"
    VOLTE = "VoLTE"


@dataclass
class CellData:
    """Represents a single cell's raw data."""
    technology: Technology
    begin_time: datetime
    end_time: datetime
    cluster: str
    site_name: str
    cell_name: str
    cell_id: str
    band: Optional[str]
    sector: Optional[str]
    
    sssr_num: float = 0.0
    sssr_den: float = 0.0
    rach_setup_sr_num: float = 0.0
    rach_setup_sr_den: float = 0.0
    ho_sr_num: float = 0.0
    ho_sr_den: float = 0.0
    erab_drop_num: float = 0.0
    erab_drop_den: float = 0.0
    dl_thp_num: float = 0.0
    dl_thp_den: float = 0.0
    ul_thp_num: float = 0.0
    ul_thp_den: float = 0.0
    ul_loss_num: float = 0.0
    dl_loss_num: float = 0.0
    cqi_num: float = 0.0
    cqi_den: float = 0.0
    rank_gt2_num: float = 0.0
    rank_gt2_den: float = 0.0
    rssi_avg_num: float = 0.0
    rssi_avg_den: float = 0.0
    ran_latency_num: float = 0.0
    ran_latency_den: float = 0.0
    dl_se_num: float = 0.0
    dl_se_den: float = 0.0
    volte_setup_num: float = 0.0
    volte_setup_den: float = 0.0
    volte_drop_num: float = 0.0
    volte_drop_den: float = 0.0
    srvcc_success_num: float = 0.0
    srvcc_success_den: float = 0.0
    
    call_setup_sr_num: float = 0.0
    call_setup_sr_den: float = 0.0
    sdcch_sr_num: float = 0.0
    sdcch_sr_den: float = 0.0
    drop_rate_num: float = 0.0
    drop_rate_den: float = 0.0


@dataclass
class KPITarget:
    """Defines a KPI target threshold."""
    name: str
    domain: KPIDomain
    technology: str
    target_percentage: float
    operator: str 
    threshold_value: float
    unit: str
    measurement_method: str = "NMS/OSS"


@dataclass
class KPIResult:
    """Results for a single KPI across all cells."""
    kpi_name: str
    target: KPITarget
    month: int
    cluster: str
    total_cells: int
    cells_meeting_target: int
    achievement_percentage: float
    passed: bool
    failing_cells: list[tuple[str, str, str, float]] 


@dataclass
class ClusterReport:
    """Complete report for a cluster."""
    cluster_name: str
    months: list[int]
    site_count: int
    cell_count: int
    kpi_results: list[KPIResult]
    last_update: datetime
    month_ranges: dict[int, tuple[datetime, datetime]] | None = None
