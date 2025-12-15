"""
Helper utility functions
"""
import re
from datetime import datetime


def clean_cell_name(name: str) -> str:
    """Clean cell name for display"""
    if not name:
        return ""
    return str(name).strip()


def format_percentage(value: float, decimals: int = 2) -> str:
    """Format value as percentage"""
    if value is None:
        return "N/A"
    return f"{round(value, decimals)}%"


def format_date_range(start_date, end_date) -> str:
    """Format date range for display - cross-platform compatible"""
    if not start_date or not end_date:
        return ""

    try:
        # Convert to datetime if string
        start = datetime.strptime(
            str(start_date), '%Y-%m-%d') if isinstance(start_date, str) else start_date
        end = datetime.strptime(
            str(end_date), '%Y-%m-%d') if isinstance(end_date, str) else end_date

        # Safe formatting (works on both Windows and Unix)
        start_str = f"{start.day} {start.strftime('%B')}"
        end_str = f"{end.day} {end.strftime('%B')}"

        return f"{start_str} to {end_str}"
    except Exception as e:
        return f"{start_date} to {end_date}"


def extract_tower_id_from_name(name: str) -> str:
    """Extract TOWER ID using regex pattern"""
    pattern = r"#([^#]+)#"
    match = re.search(pattern, str(name))
    return match.group(1) if match else None
