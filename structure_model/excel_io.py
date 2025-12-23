import pandas as pd
from structure_model.shift_parser import get_shift_info_from_two_cols
from structure_model.config import ROW_START, STEP, TAB_SHEET_NAME, WEEKEND_CODE


def get_schedule_slots(file_path, start_row, step, col_start, col_end, shift_code, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    slots = []
    current_row = start_row
    while True:
        pd_idx = current_row - 1
        if pd_idx >= len(df):
            break
        start_val = df.iloc[pd_idx, col_start]
        end_val = df.iloc[pd_idx, col_end]
        time_info = get_shift_info_from_two_cols(start_val, end_val)
        if time_info:
            time_info['shift_code'] = shift_code
            slots.append({'excel_row': current_row, 'time_info': time_info})
        current_row += step
    return slots


def get_available_drivers(file_path, day_num, shift_code):
    df = pd.read_excel(file_path, sheet_name=TAB_SHEET_NAME)
    df.columns = df.columns.astype(str)
    day_col = str(day_num)
    if day_col not in df.columns:
        if f'{day_num}.0' in df.columns:
            day_col = f'{day_num}.0'
        else:
            raise ValueError(f"Ошибка: В табеле нет колонки с названием '{day_num}'")
    drivers = []
    for _, row in df.iterrows():
        tab_no = _normalize_tab_no(row.iloc[0])
        if not tab_no:
            continue
        status_clean = str(row[day_col]).strip()
        if status_clean.startswith(str(shift_code)):
            drivers.append(tab_no)
    return drivers


def get_weekend_drivers(file_path, day_num):
    """Возвращает табельные водителей, у которых в табеле стоит выходной (код WEEKEND_CODE)."""
    df = pd.read_excel(file_path, sheet_name=TAB_SHEET_NAME)
    df.columns = df.columns.astype(str)
    day_col = str(day_num)
    if day_col not in df.columns:
        if f'{day_num}.0' in df.columns:
            day_col = f'{day_num}.0'
        else:
            raise ValueError(f"Ошибка: В табеле нет колонки с названием '{day_num}'")
    drivers = []
    for _, row in df.iterrows():
        tab_no = _normalize_tab_no(row.iloc[0])
        if not tab_no:
            continue
        status_clean = str(row[day_col]).strip().upper()
        if status_clean.startswith(str(WEEKEND_CODE)):
            drivers.append(tab_no)
    return drivers


def _normalize_tab_no(value):
    if pd.isna(value):
        return None
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        text = str(value).strip()
        return text or None
