"""Helper utilities"""

import re
import pandas as pd
import numpy as np
from datetime import datetime, timedelta


def clean_numeric(value):
    """Clean numeric values from strings"""
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return value
    
    value_str = str(value).strip()
    value_str = value_str.replace(',', '')
    value_str = re.sub(r'[^\d.\-]', '', value_str)
    
    try:
        return float(value_str) if value_str else None
    except ValueError:
        return None


def extract_tower_id(me_name):
    """Extract TOWER_ID from ME_NAME"""
    if pd.isna(me_name):
        return None
    
    me_name_str = str(me_name)
    match = re.search(r'([A-Z]{3}\d{4})', me_name_str)
    
    if match:
        return match.group(1)
    return None


def map_frequency_band(freq_band):
    """Map frequency band to MHz value"""
    if pd.isna(freq_band):
        return None
    
    band_mapping = {
        "5": 850,
        "8": 900,
        "3": 1800,
        "1": 2100,
        "40": 2300,
    }
    
    freq_str = str(freq_band).strip()
    return band_mapping.get(freq_str, None)


def format_date_mmm_yy(date_val):
    """Format date as MMM-YY (e.g., Sep-25)"""
    if pd.isna(date_val):
        return None
    
    try:
        if isinstance(date_val, str):
            dt = pd.to_datetime(date_val)
        else:
            dt = date_val
        
        return dt.strftime("%b-%y")
    except:
        return None


def format_date_range(start_date, end_date):
    """
    Format date range for display in dd-MMMM format
    Example: "19-September to 18-October"
    Works on both Windows and Linux/Mac
    """
    try:
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        
        # Use .day property instead of %-d for Windows compatibility
        start_day = str(start.day)
        start_month = start.strftime("%B")
        
        end_day = str(end.day)
        end_month = end.strftime("%B")
        
        return f"{start_day}-{start_month} to {end_day}-{end_month}"
    except Exception as e:
        print(f"⚠ Error formatting date range: {e}")
        import traceback
        traceback.print_exc()
        return "N/A"


def get_three_month_range(dates):
    """
    Get date range string for the most recent 3 months (90 days)
    Format: dd-MMMM to dd-MMMM
    Example: "19-September to 17-December"
    """
    if len(dates) == 0:
        return "N/A"
    
    try:
        dates_sorted = sorted(pd.to_datetime(dates))
        end_date = dates_sorted[-1]
        start_date = end_date - timedelta(days=89)  # 90 days total including end
        
        return format_date_range(start_date, end_date)
    except Exception as e:
        print(f"⚠ Error getting three month range: {e}")
        import traceback
        traceback.print_exc()
        return "N/A"


def get_latest_date_and_periods(dates):
    """
    Find latest date and create 3 periods of 30 days each from 90 days before latest.
    
    Returns:
        dict: {
            'latest_date': datetime,
            'start_date': datetime,
            'period_1': {'start': datetime, 'end': datetime, 'label': str, 'name': str},
            'period_2': {'start': datetime, 'end': datetime, 'label': str, 'name': str},
            'period_3': {'start': datetime, 'end': datetime, 'label': str, 'name': str}
        }
    """
    if len(dates) == 0:
        return None
    
    dates_clean = pd.to_datetime(dates).dropna()
    if len(dates_clean) == 0:
        return None
    
    latest_date = dates_clean.max()
    start_date = latest_date - timedelta(days=89)  # 90 days including latest
    
    # Period 3 (Most Recent): Days 61-90
    period_3_end = latest_date
    period_3_start = latest_date - timedelta(days=29)
    
    # Period 2 (Middle): Days 31-60
    period_2_end = period_3_start - timedelta(days=1)
    period_2_start = period_2_end - timedelta(days=29)
    
    # Period 1 (Oldest): Days 1-30
    period_1_end = period_2_start - timedelta(days=1)
    period_1_start = start_date
    
    # Get month labels for each period (use the most common month in each period)
    def get_period_label(start, end):
        """Get label as MMM-YY from middle of period"""
        mid_date = start + (end - start) / 2
        return mid_date.strftime("%b-%y")
    
    return {
        'latest_date': latest_date,
        'start_date': start_date,
        'period_1': {
            'start': period_1_start,
            'end': period_1_end,
            'label': get_period_label(period_1_start, period_1_end),
            'name': 'Period 1'
        },
        'period_2': {
            'start': period_2_start,
            'end': period_2_end,
            'label': get_period_label(period_2_start, period_2_end),
            'name': 'Period 2'
        },
        'period_3': {
            'start': period_3_start,
            'end': period_3_end,
            'label': get_period_label(period_3_start, period_3_end),
            'name': 'Period 3'
        }
    }


def assign_period_to_date(date_val, period_info):
    """
    Assign a date to one of the 3 periods.
    
    Returns:
        str: 'Period 1', 'Period 2', or 'Period 3', or None if outside range
    """
    if pd.isna(date_val) or period_info is None:
        return None
    
    date = pd.to_datetime(date_val)
    
    if period_info['period_1']['start'] <= date <= period_info['period_1']['end']:
        return 'Period 1'
    elif period_info['period_2']['start'] <= date <= period_info['period_2']['end']:
        return 'Period 2'
    elif period_info['period_3']['start'] <= date <= period_info['period_3']['end']:
        return 'Period 3'
    else:
        return None


def format_period_date_range(period_dict):
    """
    Format period date range for display in dd-MMMM format
    Example: "19-September to 18-October"
    
    Args:
        period_dict: dict with 'start' and 'end' keys (datetime objects)
    
    Returns:
        str: formatted date range
    """
    try:
        start = period_dict['start']
        end = period_dict['end']
        
        # Convert to datetime if needed
        if isinstance(start, str):
            start = pd.to_datetime(start)
        if isinstance(end, str):
            end = pd.to_datetime(end)
        
        return format_date_range(start, end)
    except Exception as e:
        print(f"⚠ Error formatting period date range: {e}")
        import traceback
        traceback.print_exc()
        return "N/A"
