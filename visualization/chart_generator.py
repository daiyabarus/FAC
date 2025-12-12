"""
Visualization Layer - Chart Generation
File: visualization/chart_generator.py
"""

import matplotlib.pyplot as plt
import numpy as np
from typing import List
from domain.models import ClusterReport, KPIResult, KPIDomain
import calendar


class ChartGenerator:
    """Generates visualization charts for KPI data."""
    
    @staticmethod
    def generate_kpi_charts(report: ClusterReport) -> None:
        """Generate and display comprehensive KPI charts."""
        kpi_groups = {}
        for result in report.kpi_results:
            if result.kpi_name not in kpi_groups:
                kpi_groups[result.kpi_name] = []
            kpi_groups[result.kpi_name].append(result)
        
        ChartGenerator._create_summary_chart(report, kpi_groups)
        ChartGenerator._create_domain_charts(report, kpi_groups)
        plt.show()
    
    @staticmethod
    def _create_summary_chart(
        report: ClusterReport,
        kpi_groups: dict[str, list[KPIResult]]
    ) -> None:
        """Create overall KPI achievement summary chart."""
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))
        fig.suptitle(f'FAC KPI Achievement Summary - {report.cluster_name}', 
                     fontsize=16, fontweight='bold')
        
        months = sorted(report.months)
        pass_counts = []
        fail_counts = []
        
        for month in months:
            month_results = [r for r in report.kpi_results if r.month == month]
            passed = sum(1 for r in month_results if r.passed)
            failed = len(month_results) - passed
            pass_counts.append(passed)
            fail_counts.append(failed)
        
        x = np.arange(len(months))
        width = 0.35
        
        ax1.bar(x - width/2, pass_counts, width, label='Passed')
        ax1.bar(x + width/2, fail_counts, width, label='Failed')
        ax1.set_xlabel('Month', fontweight='bold')
        ax1.set_ylabel('Number of KPIs', fontweight='bold')
        ax1.set_title('KPI Pass/Fail Count by Month')
        ax1.set_xticks(x)
        ax1.set_xticklabels([ChartGenerator._month_label(report, m) for m in months])
        ax1.legend()
        ax1.grid(axis='y', alpha=0.3)
        
        avg_achievements = []
        for month in months:
            month_results = [r for r in report.kpi_results if r.month == month]
            if month_results:
                avg = np.mean([r.achievement_percentage for r in month_results])
            else:
                avg = 0.0
            avg_achievements.append(avg)
        
        ax2.plot(x, avg_achievements, marker='o', linewidth=2)
        ax2.axhline(y=95, linestyle='--', label='Target (95%)', linewidth=2)
        ax2.set_xlabel('Month', fontweight='bold')
        ax2.set_ylabel('Average Achievement (%)', fontweight='bold')
        ax2.set_title('Average KPI Achievement by Month')
        ax2.set_xticks(x)
        ax2.set_xticklabels([ChartGenerator._month_label(report, m) for m in months])
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        ax2.set_ylim([0, 105])
        
        plt.tight_layout()
    
    @staticmethod
    def _create_domain_charts(
        report: ClusterReport,
        kpi_groups: dict[str, list[KPIResult]]
    ) -> None:
        """Create charts grouped by KPI domain."""
        domain_results = {}
        for kpi_name, results in kpi_groups.items():
            domain = results[0].target.domain
            if domain not in domain_results:
                domain_results[domain] = {}
            domain_results[domain][kpi_name] = results
        
        for domain, kpis in domain_results.items():
            ChartGenerator._create_single_domain_chart(report, domain, kpis)
    
    @staticmethod
    def _create_single_domain_chart(
        report: ClusterReport,
        domain: KPIDomain,
        kpis: dict[str, list[KPIResult]]
    ) -> None:
        """Create a chart for a single KPI domain."""
        fig, ax = plt.subplots(figsize=(14, 8))
        fig.suptitle(f'{domain.value} KPIs - {report.cluster_name}', 
                     fontsize=16, fontweight='bold')
        
        months = sorted(report.months)
        x = np.arange(len(months))
        width = 0.8 / len(kpis) if len(kpis) > 0 else 0.8
        
        colors = plt.cm.Set3(np.linspace(0, 1, len(kpis)))
        
        for idx, (kpi_name, results) in enumerate(kpis.items()):
            achievements = []
            for month in months:
                month_result = next((r for r in results if r.month == month), None)
                if month_result:
                    achievements.append(month_result.achievement_percentage)
                else:
                    achievements.append(0)
            
            offset = (idx - len(kpis)/2) * width + width/2
            bars = ax.bar(x + offset, achievements, width, label=kpi_name[:30])
            
            for bar in bars:
                height = bar.get_height()
                if height > 0:
                    ax.text(bar.get_x() + bar.get_width()/2., height,
                           f'{height:.1f}%',
                           ha='center', va='bottom', fontsize=8, rotation=90)
        
        ax.axhline(y=95, linestyle='--', label='Target (95%)', linewidth=2, alpha=0.7)
        ax.set_xlabel('Month', fontweight='bold')
        ax.set_ylabel('Achievement (%)', fontweight='bold')
        ax.set_title(f'{domain.value} KPI Achievement by Month')
        ax.set_xticks(x)
        ax.set_xticklabels([ChartGenerator._month_label(report, m) for m in months])
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=9)
        ax.grid(axis='y', alpha=0.3)
        ax.set_ylim([0, 105])
        
        plt.tight_layout()
    
    @staticmethod
    def _month_label(report: ClusterReport, month: int) -> str:
        """Return a compact month label using month_ranges if available."""
        if report.month_ranges and month in report.month_ranges:
            start_dt, end_dt = report.month_ranges[month]
            # e.g. "Sep 2025"
            return start_dt.strftime("%b %Y")
        else:
            # fallback
            return f"Month {month}"
    
    @staticmethod
    def save_charts_to_files(report: ClusterReport, output_folder: str) -> list[str]:
        """Save charts to image files instead of displaying them."""
        from pathlib import Path
        output_path = Path(output_folder)
        output_path.mkdir(parents=True, exist_ok=True)
        
        saved_files = []
        
        plt.ioff() 
        # Implementation left intentionally minimal; can be extended when required.
        
        return saved_files
