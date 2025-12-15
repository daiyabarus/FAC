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
    band_map = {5: 850, 3: 1800, 1: 2100, 40: 2300}

    try:
        band_num = int(freq_band)
        return band_map.get(band_num, None)
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
        result = dt.strftime('%m/%d/%Y')
        # Remove leading zeros
        parts = result.split('/')
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
