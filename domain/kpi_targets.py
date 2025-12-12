"""
Domain Layer - KPI Targets Configuration
File: domain/kpi_targets.py

Updated to include:
- RACH Success Rate secondary target (3% of cells < 55%)
- Packet Latency secondary target (5% of cells >= 40ms)
- Downlink/Uplink Throughput secondary targets
"""

from domain.models import KPITarget, KPIDomain


class KPITargets:
    """Centralized KPI target definitions."""
    
    @staticmethod
    def get_all_targets() -> list[KPITarget]:
        """Return all KPI targets as defined in requirements."""
        return [
            # 2G KPIs
            KPITarget(
                name="Call Setup Success Rate",
                domain=KPIDomain.ACCESSIBILITY,
                technology="2G RAN",
                target_percentage=95.0,
                operator=">=",
                threshold_value=98.5,
                unit="%"
            ),
            KPITarget(
                name="SDCCH Success rate",
                domain=KPIDomain.ACCESSIBILITY,
                technology="2G RAN",
                target_percentage=95.0,
                operator=">=",
                threshold_value=98.5,
                unit="%"
            ),
            KPITarget(
                name="Perceive Drop Rate",
                domain=KPIDomain.RETAINABILITY,
                technology="2G RAN",
                target_percentage=95.0,
                operator="<",
                threshold_value=2.0,
                unit="%"
            ),
            
            # 4G Accessibility
            KPITarget(
                name="Session Setup Success Rate",
                domain=KPIDomain.ACCESSIBILITY,
                technology="4G RAN",
                target_percentage=97.0,
                operator=">=",
                threshold_value=99.0,
                unit="%"
            ),
            # RACH Success Rate - Primary target
            KPITarget(
                name="RACH Success Rate",
                domain=KPIDomain.ACCESSIBILITY,
                technology="4G RAN",
                target_percentage=60.0,
                operator=">=",
                threshold_value=85.0,
                unit="%"
            ),
            # RACH Success Rate - Secondary target (at most 3% below 55%)
            KPITarget(
                name="RACH Success Rate",
                domain=KPIDomain.ACCESSIBILITY,
                technology="4G RAN",
                target_percentage=55,
                operator="<",
                threshold_value=3,
                unit="%"
            ),
            
            # 4G Mobility
            KPITarget(
                name="Handover Success Rate Inter and Intra-Frequency",
                domain=KPIDomain.MOBILITY,
                technology="4G RAN",
                target_percentage=95.0,
                operator=">=",
                threshold_value=97.0,
                unit="%"
            ),
            
            # 4G Retainability
            KPITarget(
                name="E-RAB Drop Rate",
                domain=KPIDomain.RETAINABILITY,
                technology="4G RAN",
                target_percentage=95.0,
                operator="<",
                threshold_value=2.0,
                unit="%"
            ),
            
            # 4G Integrity - Throughput (Primary targets)
            KPITarget(
                name="Downlink User Throughput",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=85.0,
                operator=">=",
                threshold_value=3.0,
                unit="Mbps @ 10 Mhz 15 Mhz 20 Mhz"
            ),
            # Downlink User Throughput - Secondary target (at most 2% below 1 Mbps)
            KPITarget(
                name="Downlink User Throughput",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=2.0,
                operator="<",
                threshold_value=1.0,
                unit="Mbps @ 10 Mhz 15 Mhz 20 Mhz"
            ),
            KPITarget(
                name="Uplink User Throughput",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=65.0,
                operator=">=",
                threshold_value=1.0,
                unit="Mbps @ 10 Mhz 15 Mhz 20 Mhz"
            ),
            # Uplink User Throughput - Secondary target (at most 2% below 0.256 Mbps)
            KPITarget(
                name="Uplink User Throughput",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=2.0,
                operator="<",
                threshold_value=0.256,
                unit="Mbps @ 10 Mhz 15 Mhz 20 Mhz"
            ),
            
            # 4G Integrity - Packet Loss
            KPITarget(
                name="UL Packet Loss (PDCP )",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=97.0,
                operator="<",
                threshold_value=0.85,
                unit="%"
            ),
            KPITarget(
                name="DL Packet Loss (PDCP )",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=97.0,
                operator="<",
                threshold_value=0.10,
                unit="%"
            ),
            
            # 4G Integrity - CQI
            KPITarget(
                name="CQI",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=95.0,
                operator=">=",
                threshold_value=7.0,
                unit="num"
            ),
            
            # 4G Integrity - MIMO (Primary and Secondary)
            KPITarget(
                name="MIMO Transmission Rank2 Rate",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=70.0,
                operator=">=",
                threshold_value=35.0,
                unit="%"
            ),
            KPITarget(
                name="MIMO Transmission Rank2 Rate",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=5.0,
                operator="<",
                threshold_value=20.0,
                unit="%"
            ),
            
            # 4G Integrity - RSSI
            KPITarget(
                name="UL RSSI",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=97.0,
                operator="<",
                threshold_value=-105.0,
                unit="dBm"
            ),
            
            # 4G Integrity - Latency (Primary and Secondary)
            KPITarget(
                name="Packet Latency",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=95.0,
                operator="<",
                threshold_value=30.0,
                unit="ms, @PDCP\nms, @VoLTE"
            ),
            # Packet Latency - Secondary (at most 5% >= 40ms)
            KPITarget(
                name="Packet Latency",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=5.0,
                operator=">=",
                threshold_value=40.0,
                unit="ms, @PDCP\nms, @VoLTE"
            ),
            
            # 4G Integrity - Spectral Efficiency
            KPITarget(
                name="Spectral Efficiency",
                domain=KPIDomain.INTEGRITY,
                technology="4G RAN",
                target_percentage=90.0,
                operator=">=",
                threshold_value=1.1,  # This varies by band, handled in calculator
                unit="num"
            ),
            
            # 4G VoLTE
            KPITarget(
                name="Voice Call Success Rate (VoLTE)",
                domain=KPIDomain.VOLTE,
                technology="4G RAN",
                target_percentage=95.0,
                operator=">=",
                threshold_value=97.0,
                unit="%"
            ),
            KPITarget(
                name="Voice Call Drop Rate (VoLTE)",
                domain=KPIDomain.VOLTE,
                technology="4G RAN",
                target_percentage=95.0,
                operator="<",
                threshold_value=2.0,
                unit="%"
            ),
            KPITarget(
                name="SRVCC Success Rate",
                domain=KPIDomain.VOLTE,
                technology="4G RAN",
                target_percentage=95.0,
                operator=">=",
                threshold_value=97.0,
                unit="%"
            ),
            
            # 5G Integrity KPIs
            KPITarget(
                name="CQI",
                domain=KPIDomain.INTEGRITY,
                technology="5G RAN",
                target_percentage=30.0,
                operator=">=",
                threshold_value=10.0,
                unit="num"
            ),
            KPITarget(
                name="UL RSSI",
                domain=KPIDomain.INTEGRITY,
                technology="5G RAN",
                target_percentage=95.0,
                operator="<",
                threshold_value=-105.0,
                unit="dBm"
            ),
            KPITarget(
                name="Packet Latency",
                domain=KPIDomain.INTEGRITY,
                technology="5G RAN",
                target_percentage=95.0,
                operator="<",
                threshold_value=20.0,
                unit="ms"
            ),
        ]
    
    @staticmethod
    def get_spectral_efficiency_thresholds() -> dict[str, float]:
        """Return SE thresholds by band and MIMO configuration."""
        return {
            "2T2R_850": 1.1,
            "2T2R_900": 1.1,
            "2T2R_1800": 1.25,
            "2T2R_2100": 1.3,
            "4T4R_1800": 1.5,
            "8T8R_1800": 1.5,
            "4T4R_2100": 1.7,
            "8T8R_2100": 1.7,
            "4T4R_2300": 1.7,
            "8T8R_2300": 1.7,
            "MM_1800": 1.25,
            "MM_2300": 2.1,
        }