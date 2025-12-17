"""Helper utility functions"""

import re
from datetime import datetime
import pandas as pd
import numpy as np


def clean_numeric(value):
    """Clean numeric values - remove commas and percentages"""
    if pd.isna(value) or value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    # Convert to string and clean
    str_val = str(value).strip()

    # Remove percentage sign
    if "%" in str_val:
        str_val = str_val.replace("%", "")
        try:
            return float(str_val) / 100
        except:
            return None

    # Remove commas (thousand separator)
    str_val = str_val.replace(",", "")

    try:
        return float(str_val)
    except:
        return None


def extract_tower_id(me_name):
    """Extract tower ID from ME name using regex #tower_id#"""
    if pd.isna(me_name):
        return None

    pattern = r"#([^#]+)#"
    match = re.search(pattern, str(me_name))

    if match:
        return match.group(1)
    return None


def map_frequency_band(freq_band):
    """Map frequency band number to actual frequency"""
    band_map = {
        5: 850,
        3: 1800,
        1: 2100,
        40: 2300,
        "5": 850,
        "3": 1800,
        "1": 2100,
        "40": 2300,
    }

    try:
        if isinstance(freq_band, str):
            freq_band = int(freq_band)
        return band_map.get(freq_band, None)
    except:
        return None


# def format_date_mdy(date_str):
#     """Convert date from YYYY-MM-DD to M/D/YYYY"""
#     try:
#         if isinstance(date_str, str):
#             dt = pd.to_datetime(date_str)
#         else:
#             dt = date_str
#         return dt.strftime("%-m/%-d/%Y")
#     except:
#         return None
def format_date_mdy(date_str):
    """Convert date from YYYY-MM-DD to M/D/YYYY"""
    try:
        if isinstance(date_str, str):
            dt = pd.to_datetime(date_str)
        else:
            dt = date_str
        # Use strftime without leading zeros
        # For cross-platform compatibility, use replace instead of %-
        result = dt.strftime("%m/%d/%Y")
        # Remove leading zeros
        parts = result.split("/")
        return f"{int(parts[0])}/{int(parts[1])}/{parts[2]}"
    except:
        return None


def format_date_mmm_yy(date_str):
    """Convert date to MMM-YY format (e.g., Sep-25)"""
    try:
        if isinstance(date_str, str):
            dt = pd.to_datetime(date_str)
        else:
            dt = date_str
        return dt.strftime("%b-%y")
    except:
        return None


def format_month_name(date_str):
    """Get full month name (e.g., September)"""
    try:
        if isinstance(date_str, str):
            dt = pd.to_datetime(date_str)
        else:
            dt = date_str
        return dt.strftime("%B")
    except:
        return None


def format_date_range(start_date, end_date):
    """Format date range (e.g., 1 October - 31 October)"""
    try:
        if isinstance(start_date, str):
            start_dt = pd.to_datetime(start_date)
        else:
            start_dt = start_date

        if isinstance(end_date, str):
            end_dt = pd.to_datetime(end_date)
        else:
            end_dt = end_date

        return f"{start_dt.day} {start_dt.strftime('%B')} - {end_dt.day} {end_dt.strftime('%B')}"
    except:
        return None


def get_three_month_range(dates):
    """Get 3-month range string (e.g., 1st October to 31st December)"""
    try:
        sorted_dates = sorted([pd.to_datetime(d) for d in dates])
        start = sorted_dates[0]
        end = sorted_dates[-1]

        # Add ordinal suffix
        def ordinal(n):
            suffix = ["th", "st", "nd", "rd"] + ["th"] * 6
            if 10 <= n % 100 <= 20:
                return f"{n}th"
            return f"{n}{suffix[n % 10]}"

        return f"{ordinal(start.day)} {start.strftime('%B')} to {ordinal(end.day)} {end.strftime('%B')}"
    except:
        return None


def detect_tx_from_cellname(cell_name):
    """Detect TX configuration from LTE cell name"""
    if pd.isna(cell_name):
        return None

    cell_str = str(cell_name).upper()

    # Pattern detection based on cell name prefix
    # Priority order matters!

    # 32T32R: N_AC4G23 or explicit 32T32R
    if "N_AC4G23" in cell_str or "32T32R" in cell_str:
        return "32T32R"

    # 8T8R: explicit marker
    if "8T8R" in cell_str or "_8T8R" in cell_str:
        return "8T8R"

    # 4T4R: AC4G18, AC4G21 (1800/2100 MHz bands)
    if "AC4G18" in cell_str or "AC4G21" in cell_str:
        return "4T4R"

    if "4T4R" in cell_str or "_4T4R" in cell_str:
        return "4T4R"

    # 2T2R: AC4G85 (850 MHz band) or explicit marker
    if "AC4G85" in cell_str or "N_AC4G85" in cell_str:
        return "2T2R"

    if "2T2R" in cell_str or "_2T2R" in cell_str:
        return "2T2R"

    # If no pattern matches, return None
    return None


def extract_band_from_cellname(cell_name):
    """Extract frequency band from LTE cell name - GENERIC pattern"""
    if pd.isna(cell_name):
        return None

    cell_str = str(cell_name).upper()

    # Generic pattern: support AC4G, MD4G, and other variants
    # Pattern: {prefix}4G{code}

    import re

    # Try to find pattern: any prefix + 4G + number
    match = re.search(r"(?:[A-Z_]+)?4G(\d+)", cell_str)
    if match:
        band_code = match.group(1)

        # Map code to frequency
        code_to_freq = {
            "85": 850,
            "18": 1800,
            "21": 2100,
            "23": 2300,
        }

        return code_to_freq.get(band_code, None)

    return None
