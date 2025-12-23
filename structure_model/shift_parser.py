import re
from datetime import datetime, timedelta
import pandas as pd

def get_shift_info_from_two_cols(start_val, end_val):
    if pd.isna(start_val) or pd.isna(end_val):
        return None
    try:
        start_str = str(start_val).strip().split()[0]
        end_str = str(end_val).strip().split()[0]
        match_start = re.search(r'(\d{1,2})[:\.\-](\d{2})', start_str)
        match_end = re.search(r'(\d{1,2})[:\.\-](\d{2})', end_str)
        if not match_start or not match_end:
            return None
        h1, m1 = map(int, match_start.groups())
        h2, m2 = map(int, match_end.groups())
        now = datetime.now().replace(second=0, microsecond=0)
        start_dt = now.replace(hour=h1, minute=m1)
        end_dt = now.replace(hour=h2, minute=m2)
        is_next_day = False
        if end_dt < start_dt:
            end_dt += timedelta(days=1)
            is_next_day = True
        duration = (end_dt - start_dt).total_seconds() / 3600
        return {
            'start_dt': start_dt,
            'end_dt': end_dt,
            'start_str': f"{h1:02}:{m1:02}",
            'end_str': f"{h2:02}:{m2:02}",
            'duration': round(duration, 2),
            'is_next_day': is_next_day
        }
    except:
        return None


def calculate_rest_duration(end_str_N, is_next_day_N, start_str_N1):
    try:
        h_end, m_end = map(int, end_str_N.split(':'))
        h_start, m_start = map(int, start_str_N1.split(':'))
        base_dt = datetime.now().replace(second=0, microsecond=0)
        end_dt_N = base_dt.replace(hour=h_end, minute=m_end)
        if not is_next_day_N:
            end_dt_N -= timedelta(days=1)
        start_dt_N1 = base_dt.replace(hour=h_start, minute=m_start)
        while start_dt_N1 <= end_dt_N:
            start_dt_N1 += timedelta(days=1)
        rest_duration = (start_dt_N1 - end_dt_N).total_seconds() / 3600
        return round(rest_duration, 1)
    except:
        return "Н/Д"