"""
Infrastructure Layer - CSV File Parsing
File: infrastructure/csv_parser.py

Behavior:
- Use the explicit regex pattern "#([^#]+)#" to extract the towerid from element_name (row[6]) for FDD/TDD.
- If regex finds a match, lookup towerid -> cluster. If not found or no mapping, SKIP row (strict).
- GSM: use site_name at row[8] to lookup cluster; if missing, SKIP.
"""

import csv
import re
from pathlib import Path
from datetime import datetime
from typing import Optional
from domain.models import CellData, Technology


class CSVParser:
    """Parses CSV files and extracts cell data with strict cluster mapping using a tower-id regex."""

    # mapping for cell id -> (sector, band)
    CELL_ID_MAPPING = {
        "131": ["1", "850"], "132": ["2", "850"], "133": ["3", "850"], "134": ["4", "850"],
        "4": ["1", "1800"], "5": ["2", "1800"], "6": ["3", "1800"], "24": ["4", "1800"],
        "51": ["11", "1800"], "52": ["12", "1800"], "53": ["13", "1800"], "54": ["14", "1800"],
        "55": ["15", "1800"], "56": ["16", "1800"], "14": ["M1", "1800"], "15": ["M2", "1800"],
        "16": ["M3", "1800"], "64": ["M4", "1800"],
        "1": ["1", "2100"], "2": ["2", "2100"], "3": ["3", "2100"], "7": ["1", "2100"],
        "8": ["2", "2100"], "9": ["3", "2100"], "97": ["11", "2100"], "27": ["4", "2100"],
        "91": ["11", "2100"], "92": ["12", "2100"], "93": ["13", "2100"], "94": ["14", "2100"],
        "95": ["15", "2100"], "96": ["16", "2100"], "17": ["M1", "2100"], "18": ["M2", "2100"],
        "19": ["M3", "2100"], "67": ["M4", "2100"],
        "111": ["1", "2300"], "112": ["2", "2300"], "113": ["3", "2300"], "114": ["4", "2300"],
        "141": ["11", "2300"], "142": ["12", "2300"], "143": ["13", "2300"], "144": ["14", "2300"],
        "145": ["15", "2300"], "146": ["16", "2300"], "121": ["1", "2300"], "122": ["2", "2300"],
        "123": ["3", "2300"], "124": ["4", "2300"], "151": ["11", "2300"], "152": ["12", "2300"],
        "153": ["13", "2300"], "154": ["14", "2300"], "155": ["15", "2300"], "156": ["16", "2300"]
    }

    # regex pattern to extract towerid from element_name per your spec
    tower_id_pattern = re.compile(r"#([^#]+)#")

    def __init__(self):
        # normalized maps (lowercase)
        self.towerid_to_cluster: dict[str, str] = {}
        self.site_name_to_cluster: dict[str, str] = {}
        self.site_name_to_towerid: dict[str, str] = {}

    def load_tower_mapping(self, tower_csv_path: Path) -> None:
        """
        Load TOWERID CSV. Expected columns: CLUSTER (0), TOWERID (1), SITE_NAME (2)
        Builds:
          - towerid_to_cluster[towerid.lower()] = cluster
          - site_name_to_cluster[site_name.lower()] = cluster
          - site_name_to_towerid[site_name.lower()] = towerid (if provided)
        """
        with open(tower_csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            next(reader, None)  # skip header if present
            for row in reader:
                if len(row) >= 3:
                    cluster = self._clean_value(row[0])
                    towerid = self._clean_value(row[1])
                    site_name = self._clean_value(row[2])
                    if towerid:
                        self.towerid_to_cluster[towerid.lower()] = cluster
                    if site_name:
                        self.site_name_to_cluster[site_name.lower()] = cluster
                        if towerid:
                            self.site_name_to_towerid[site_name.lower()] = towerid

        print(f"[CSVParser] Loaded {len(self.towerid_to_cluster)} towerid mappings and {len(self.site_name_to_cluster)} site-name mappings")

    # --------------------------
    # FDD/TDD parsing
    # --------------------------
    def parse_lte_csv(self, csv_path: Path, technology: Technology) -> list[CellData]:
        """Parse FDD/TDD CSV rows (0..44 columns expected) using fixed indexes and strict regex mapping."""
        parsed = []
        skipped = 0
        total = 0
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                total += 1
                if len(row) < 45:
                    skipped += 1
                    continue
                cell = self._parse_lte_row(row, technology)
                if cell:
                    parsed.append(cell)
                else:
                    skipped += 1

        print(f"[CSVParser] {csv_path.name}: Parsed {len(parsed)} mapped LTE rows, Skipped {skipped}")
        return parsed

    def _parse_lte_row(self, row: list[str], technology: Technology) -> Optional[CellData]:
        """
        Parse single LTE row. Use tower_id_pattern to find towerid inside element_name (row[6]).
        If pattern matches and towerid exists in tower mapping => use that cluster.
        If none => SKIP (strict).
        """

        try:
            element_name = self._clean_value(row[6])
            match = self.tower_id_pattern.search(element_name)
            if not match:
                # no towerid token found inside element_name -> skip (strict)
                return None

            extracted_towerid = match.group(1).strip()
            if not extracted_towerid:
                return None

            cluster = self.towerid_to_cluster.get(extracted_towerid.lower())
            if not cluster:
                # towerid found in element_name but not present in tower mapping -> skip
                return None

            # Site name — try lte_name (index 10) else fallback to subnet_name (4) else towerid
            site_name = self._clean_value(row[10]) or self._clean_value(row[4]) or extracted_towerid

            cell_id = self._clean_value(row[13])
            band, sector = self._get_band_sector(cell_id)

            begin_time = self._parse_datetime(row[0])
            end_time = self._parse_datetime(row[1])

            return CellData(
                technology=technology,
                begin_time=begin_time,
                end_time=end_time,
                cluster=cluster,
                site_name=site_name,
                cell_name=self._clean_value(row[12]),
                cell_id=cell_id,
                band=band,
                sector=sector,
                sssr_num=self._parse_float(row[15]),
                sssr_den=self._parse_float(row[16]),
                rach_setup_sr_num=self._parse_float(row[17]),
                rach_setup_sr_den=self._parse_float(row[18]),
                ho_sr_num=self._parse_float(row[19]),
                ho_sr_den=self._parse_float(row[20]),
                erab_drop_num=self._parse_float(row[21]),
                erab_drop_den=self._parse_float(row[22]),
                dl_thp_num=self._parse_float(row[23]),
                dl_thp_den=self._parse_float(row[24]),
                ul_thp_num=self._parse_float(row[25]),
                ul_thp_den=self._parse_float(row[26]),
                ul_loss_num=self._parse_float(row[27]),
                dl_loss_num=self._parse_float(row[28]),
                cqi_num=self._parse_float(row[29]),
                cqi_den=self._parse_float(row[30]),
                rank_gt2_num=self._parse_float(row[31]),
                rank_gt2_den=self._parse_float(row[32]),
                rssi_avg_num=self._parse_float(row[33]),
                rssi_avg_den=self._parse_float(row[34]),
                ran_latency_num=self._parse_float(row[35]),
                ran_latency_den=self._parse_float(row[36]),
                dl_se_num=self._parse_float(row[37]),
                dl_se_den=self._parse_float(row[38]),
                volte_setup_num=self._parse_float(row[39]),
                volte_setup_den=self._parse_float(row[40]),
                volte_drop_num=self._parse_float(row[41]),
                volte_drop_den=self._parse_float(row[42]),
                srvcc_success_num=self._parse_float(row[43]),
                srvcc_success_den=self._parse_float(row[44]),
            )
        except Exception as e:
            print(f"[CSVParser] Error parsing LTE row: {e}")
            return None

    # --------------------------
    # GSM parsing
    # --------------------------
    def parse_gsm_csv(self, csv_path: Path) -> list[CellData]:
        """Parse GSM CSV using fixed indexes and strict site_name -> cluster mapping."""
        parsed = []
        skipped = 0
        total = 0
        with open(csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                total += 1
                if len(row) < 18:
                    skipped += 1
                    continue
                cell = self._parse_gsm_row(row)
                if cell:
                    parsed.append(cell)
                else:
                    skipped += 1

        print(f"[CSVParser] {csv_path.name}: Parsed {len(parsed)} mapped GSM rows, Skipped {skipped}")
        return parsed

    def _parse_gsm_row(self, row: list[str]) -> Optional[CellData]:
        """
        Parse GSM row:
        - site_name at index 8 -> lookup site_name_to_cluster
        - if found -> include, else skip
        """
        try:
            raw_site_name = self._clean_value(row[8])
            cluster = self.site_name_to_cluster.get(raw_site_name.lower())
            if not cluster:
                return None

            begin_time = self._parse_datetime(row[0])
            end_time = self._parse_datetime(row[1])

            return CellData(
                technology=Technology.GSM,
                begin_time=begin_time,
                end_time=end_time,
                cluster=cluster,
                site_name=raw_site_name,
                cell_name=self._clean_value(row[10]),
                cell_id=self._clean_value(row[9]),
                band=self._clean_value(row[11]),
                sector=None,
                call_setup_sr_num=self._parse_float(row[12]),
                call_setup_sr_den=self._parse_float(row[13]),
                sdcch_sr_num=self._parse_float(row[14]),
                sdcch_sr_den=self._parse_float(row[15]),
                drop_rate_num=self._parse_float(row[16]),
                drop_rate_den=self._parse_float(row[17]),
            )
        except Exception as e:
            print(f"[CSVParser] Error parsing GSM row: {e}")
            return None

    # --------------------------
    # Utilities
    # --------------------------
    def _get_band_sector(self, cell_id: str) -> tuple[Optional[str], Optional[str]]:
        mapping = self.CELL_ID_MAPPING.get(cell_id)
        if mapping:
            return mapping[1], mapping[0]
        return None, None

    @staticmethod
    def _clean_value(value: str) -> str:
        try:
            return str(value).replace(',', '').replace('%', '').strip()
        except Exception:
            return str(value).strip()

    @staticmethod
    def _parse_float(value: str) -> float:
        try:
            cleaned = CSVParser._clean_value(value)
            return float(cleaned) if cleaned else 0.0
        except Exception:
            return 0.0

    @staticmethod
    def _parse_datetime(value: str) -> datetime:
        try:
            for fmt in ["%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
                        "%d-%m-%Y %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]:
                try:
                    return datetime.strptime(value, fmt)
                except ValueError:
                    continue
            return datetime.now()
        except Exception:
            return datetime.now()
