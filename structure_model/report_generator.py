import pandas as pd
import json
import os
from structure_model.config import HISTORY_JSON_DIR, OUTPUT_DIR, TAB_SHEET_NAME
from structure_model.shift_parser import calculate_rest_duration


def _normalize_driver_ids(df, driver_id_col):
    """Привести табельные к строкам без .0, как в new_model.generate_work_summary."""
    all_ids = []
    for x in df[driver_id_col].dropna():
        try:
            all_ids.append(str(int(float(x))))
        except (ValueError, TypeError):
            # если не число — просто приводим к строке
            all_ids.append(str(x).strip())
    return all_ids


def generate_work_summary(target_day, file_path):
    print("\n\n=== ГЕНЕРАЦИЯ СВОДНОГО ОТЧЕТА ===")
    try:
        # Берём тот же лист "Весь_табель", где описаны все типы графиков
        df_tab = pd.read_excel(file_path, sheet_name=TAB_SHEET_NAME)
        driver_id_col = df_tab.columns[0]
        graphik_col = df_tab.columns[1]

        # Карта: таб. номер -> тип графика (4*2, 3*3 и т.п.)
        df_temp = df_tab.set_index(df_tab[driver_id_col].astype(str))
        driver_graphik = df_temp[graphik_col].to_dict()

        # Список всех табельных номеров в нормализованном виде
        all_drivers = _normalize_driver_ids(df_tab, driver_id_col)
    except Exception as e:
        print(f"Ошибка чтения листа '{TAB_SHEET_NAME}' для отчёта: {e}")
        return

    history_data = {}
    for day in range(1, target_day + 2):
        filename = os.path.join(HISTORY_JSON_DIR, f"history_{day}.json")
        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                history_data[day] = json.load(f)
        else:
            history_data[day] = {}

    report_data = []
    for driver_id in all_drivers:
        row = {'График': driver_graphik.get(driver_id, 'N/A')}
        for day in range(1, target_day + 1):
            day_hist = history_data.get(day, {})
            if driver_id in day_hist:
                work = day_hist[driver_id]
                duration = work['duration']
                end_str = work['end_str']
                is_next = work.get('is_next_day', False)
                next_hist = history_data.get(day + 1, {})
                if driver_id in next_hist:
                    rest = calculate_rest_duration(end_str, is_next, next_hist[driver_id]['start_str'])
                    rest_rep = f"Отдых: {rest} ч."
                elif day == target_day:
                    rest_rep = "Отдых: Н/Д"
                else:
                    rest_rep = "Отдых: Полный"
                row[f'День {day}'] = f"Работа: {duration} ч. | {rest_rep}"
            else:
                row[f'День {day}'] = "--- Выходной"
        report_data.append(row)

    df_report = pd.DataFrame(report_data, index=pd.Index(all_drivers, name='Таб. №'))
    report_file = f"{OUTPUT_DIR}/Отчет_Нагрузки_Дни_1_по_{target_day}.xlsx"
    df_report.to_excel(report_file)
    print(f"\n[Отчет] Создан: **{report_file}**")
    return report_file