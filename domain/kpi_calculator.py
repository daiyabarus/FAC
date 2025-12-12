"""
Domain Layer - KPI Calculation Logic
File: domain/kpi_calculator.py
"""

from typing import Optional
from domain.models import CellData, KPIResult, KPITarget, KPIDomain


class KPICalculator:
    """Calculates KPI values for cells."""
    
    @staticmethod
    def calculate_cell_kpi(cell: CellData, kpi_name: str) -> Optional[float]:
        """Calculate a specific KPI value for a cell."""
        
        if kpi_name == "Call Setup Success Rate":
            return KPICalculator._safe_divide(cell.call_setup_sr_num, cell.call_setup_sr_den) * 100
        
        elif kpi_name == "SDCCH Success rate":
            return KPICalculator._safe_divide(cell.sdcch_sr_num, cell.sdcch_sr_den) * 100
        
        elif kpi_name == "Perceive Drop Rate":
            return KPICalculator._safe_divide(cell.drop_rate_num, cell.drop_rate_den) * 100
        
        elif kpi_name == "Session Setup Success Rate":
            return KPICalculator._safe_divide(cell.sssr_num, cell.sssr_den) * 100
        
        elif kpi_name == "RACH Success Rate":
            return KPICalculator._safe_divide(cell.rach_setup_sr_num, cell.rach_setup_sr_den) * 100
        
        elif kpi_name == "Handover Success Rate Inter and Intra-Frequency":
            return KPICalculator._safe_divide(cell.ho_sr_num, cell.ho_sr_den) * 100
        
        elif kpi_name == "E-RAB Drop Rate":
            return KPICalculator._safe_divide(cell.erab_drop_num, cell.erab_drop_den) * 100
        
        elif kpi_name == "Downlink User Throughput":
            return KPICalculator._safe_divide(cell.dl_thp_num, cell.dl_thp_den) 
        
        elif kpi_name == "Uplink User Throughput":
            return KPICalculator._safe_divide(cell.ul_thp_num, cell.ul_thp_den)
        
        elif kpi_name == "UL Packet Loss (PDCP )":
            return KPICalculator._safe_divide(cell.ul_loss_num, cell.ul_thp_den) * 100
        
        elif kpi_name == "DL Packet Loss (PDCP )":
            return KPICalculator._safe_divide(cell.dl_loss_num, cell.dl_thp_den) * 100
        
        elif kpi_name == "CQI":
            return KPICalculator._safe_divide(cell.cqi_num, cell.cqi_den)
        
        elif kpi_name == "MIMO Transmission Rank2 Rate":
            return KPICalculator._safe_divide(cell.rank_gt2_num, cell.rank_gt2_den) * 100
        
        elif kpi_name == "UL RSSI":
            return KPICalculator._safe_divide(cell.rssi_avg_num, cell.rssi_avg_den)
        
        elif kpi_name == "Packet Latency":
            return KPICalculator._safe_divide(cell.ran_latency_num, cell.ran_latency_den)
        
        elif kpi_name == "Spectral Efficiency":
            return KPICalculator._calculate_spectral_efficiency(cell)
        
        elif kpi_name == "Voice Call Success Rate (VoLTE)":
            return KPICalculator._safe_divide(cell.volte_setup_num, cell.volte_setup_den) * 100
        
        elif kpi_name == "Voice Call Drop Rate (VoLTE)":
            return KPICalculator._safe_divide(cell.volte_drop_num, cell.volte_drop_den) * 100
        
        elif kpi_name == "SRVCC Success Rate":
            return KPICalculator._safe_divide(cell.srvcc_success_num, cell.srvcc_success_den) * 100
        
        return None
    
    @staticmethod
    def _safe_divide(numerator: float, denominator: float) -> float:
        """Safely divide two numbers, returning 0 if denominator is 0."""
        return numerator / denominator if denominator > 0 else 0.0
    
    @staticmethod
    def _calculate_spectral_efficiency(cell: CellData) -> Optional[float]:
        """Calculate spectral efficiency based on band and MIMO configuration."""
        if cell.dl_se_den == 0:
            return None
        
        se_value = cell.dl_se_num / cell.dl_se_den
        return se_value
    
    @staticmethod
    def check_kpi_threshold(value: Optional[float], target: KPITarget) -> bool:
        """Check if a KPI value meets its target threshold."""
        if value is None:
            return False
        
        if target.operator == ">=":
            return value >= target.threshold_value
        elif target.operator == "<":
            return value < target.threshold_value
        elif target.operator == ">":
            return value > target.threshold_value
        elif target.operator == "<=":
            return value <= target.threshold_value
        
        return False
    
    @staticmethod
    def aggregate_kpi_results(
        cells: list[CellData],
        target: KPITarget,
        month: int,
        cluster: str
    ) -> KPIResult:
        """Aggregate KPI results for a list of cells."""
        total_cells = len(cells)
        cells_meeting_target = 0
        failing_cells = []
        
        for cell in cells:
            kpi_value = KPICalculator.calculate_cell_kpi(cell, target.name)
            
            if kpi_value is not None:
                meets_target = KPICalculator.check_kpi_threshold(kpi_value, target)
                
                if meets_target:
                    cells_meeting_target += 1
                else:
                    failing_cells.append((
                        cell.site_name,
                        cell.band or "N/A",
                        cell.cell_name,
                        kpi_value
                    ))
        
        achievement_percentage = (cells_meeting_target / total_cells * 100) if total_cells > 0 else 0.0
        passed = achievement_percentage >= target.target_percentage
        
        return KPIResult(
            kpi_name=target.name,
            target=target,
            month=month,
            cluster=cluster,
            total_cells=total_cells,
            cells_meeting_target=cells_meeting_target,
            achievement_percentage=achievement_percentage,
            passed=passed,
            failing_cells=failing_cells
        )
