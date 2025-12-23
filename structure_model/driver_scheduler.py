# structure_model/driver_scheduler.py
from datetime import datetime, timedelta
from openpyxl import load_workbook
import json
import os

from structure_model.config import (
    FILE_PATH, ROW_START, STEP,
    COL_SHIFT_1_START, COL_SHIFT_1_END, COL_SHIFT_2_START, COL_SHIFT_2_END,
    COL_SHIFT_1_INSERT, COL_SHIFT_2_INSERT, REST_HOURS,
    ROUTE_NUMBER, SCHEDULE_SHEETS, ALLOW_WEEKEND_EXTRA_WORK,
    ABSENCES_FILE, OUTPUT_DIR
)
from structure_model.excel_io import get_schedule_slots, get_available_drivers, get_weekend_drivers
from structure_model.history_manager import load_history, save_history

def get_rest_hours_for_driver(driver_id, history_data, target_shift_start_dt):
    drv_str = str(driver_id)
    if drv_str not in history_data:
        return REST_HOURS

    last_work = history_data[drv_str]
    last_end_str = last_work.get('end_str')
    was_next_day = last_work.get('is_next_day', False)

    try:
        base_date = target_shift_start_dt.date() - timedelta(days=1)
        if was_next_day:
            base_date = base_date + timedelta(days=1)

        h, m = map(int, last_end_str.split(':'))
        last_end_dt = datetime(base_date.year, base_date.month, base_date.day, h, m)
        rest_hours = (target_shift_start_dt - last_end_dt).total_seconds() / 3600
        return round(rest_hours, 1)
    except Exception:
        return -9999.0


def worked_same_shift_yesterday(driver_id, history_data, shift_code):
    drv_str = str(driver_id)
    if drv_str not in history_data:
        return False
    return history_data[drv_str].get('shift_code') == shift_code


def _load_absent_drivers_for_day_shift(day: int, shift_code: int):
    """
    Сначала пробуем получить отсутствующих из БД (Absence), если доступен контекст Flask.
    Если DB/контекст недоступен, откатываемся к чтению ABSENCES_FILE (как раньше).
    Возвращаем set таб. номеров (строк).
    """
    # Попытка: получить через SQLAlchemy-модель Absence
    try:
        from flask import current_app
        # импорт модели через серверный модуль
        from structure_model.server import Absence
        # Если нет контекста приложения, вызовет RuntimeError в query
        try:
            with current_app.app_context():
                items = Absence.query.filter_by(day=day, shift=shift_code).all()
                absent = set()
                for a in items:
                    if a.tab_no:
                        absent.add(str(a.tab_no).strip())
                return absent
        except RuntimeError:
            # нет контекста — падаем на файловую логику ниже
            pass
        except Exception:
            # любая другая ошибка — откат к файловому варианту
            pass
    except Exception:
        # не удалось подключиться к DB/Flask — используем ABSENCES_FILE
        pass

    # Fallback: чтение ABSENCES_FILE (старый путь)
    if not ABSENCES_FILE or not os.path.exists(ABSENCES_FILE):
        return set()

    try:
        with open(ABSENCES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return set()

    absent = set()
    for entry in data:
        try:
            if int(entry.get("day")) == day and int(entry.get("shift")) == shift_code:
                tab_no = str(entry.get("tab_no")).strip()
                if tab_no:
                    absent.add(tab_no)
        except (TypeError, ValueError):
            continue

    return absent


def choose_driver_for_slot(candidates, history_data, target_shift_start_dt, shift_code, assigned_today_set):
    scored = []
    for drv in list(candidates):
        drv_str = str(drv)
        if drv_str in assigned_today_set:
            continue

        rest_h = get_rest_hours_for_driver(drv, history_data, target_shift_start_dt)
        if rest_h == -9999.0:
            continue
        if rest_h < -0.5:
            continue

        same_shift = 0 if worked_same_shift_yesterday(drv, history_data, shift_code) else 1

        if rest_h >= REST_HOURS:
            score = (abs(rest_h - REST_HOURS), same_shift, -rest_h)
        else:
            deficit = REST_HOURS - rest_h
            score = (1000 + deficit, same_shift, -rest_h)

        scored.append((score, drv_str))

    if not scored:
        return None

    scored.sort(key=lambda item: item[0])
    return scored[0][1]


def run_planner(target_day, prev_day, route_number=None):
    """
    Планировщик на один день для одного маршрута
    """
    if route_number is None:
        route_number = ROUTE_NUMBER

    print(f"ЗАПУСК ПЛАНИРОВЩИКА: ДЕНЬ {target_day}, маршрут {route_number} (история дня {prev_day})")

    # ---------------- ДАТА ----------------
    current_year = datetime.now().year
    current_month = datetime.now().month
    try:
        target_date = datetime(current_year, current_month, target_day)
    except Exception as e:
        print(f"Неверный день: {e}")
        return

    is_weekend = target_date.weekday() >= 5

    # ---------------- ЛИСТ РАСПИСАНИЯ ----------------
    sheet_name = SCHEDULE_SHEETS.get((route_number, is_weekend))
    if not sheet_name:
        print(f"Не найден лист расписания для маршрута {route_number}, is_weekend={is_weekend}")
        return

    # ---------------- ИСТОРИЯ ----------------
    history = load_history(prev_day)
    today_history_log = {}

    # ---------------- СЛОТЫ ----------------
    slots_s1 = get_schedule_slots(
        FILE_PATH, ROW_START, STEP,
        COL_SHIFT_1_START, COL_SHIFT_1_END,
        1, sheet_name
    )
    slots_s2 = get_schedule_slots(
        FILE_PATH, ROW_START, STEP,
        COL_SHIFT_2_START, COL_SHIFT_2_END,
        2, sheet_name
    )

    try:
        candidates_s1 = get_available_drivers(FILE_PATH, target_day, 1)
        candidates_s2 = get_available_drivers(FILE_PATH, target_day, 2)
    except ValueError as e:
        print(e)
        return

    # ---------------- ОТСУТСТВИЯ ----------------
    absent_s1 = _load_absent_drivers_for_day_shift(target_day, 1)
    absent_s2 = _load_absent_drivers_for_day_shift(target_day, 2)

    candidates_s1 = [d for d in candidates_s1 if str(d) not in absent_s1]
    candidates_s2 = [d for d in candidates_s2 if str(d) not in absent_s2]

    weekend_pool = []
    if ALLOW_WEEKEND_EXTRA_WORK:
        try:
            weekend_pool = get_weekend_drivers(FILE_PATH, target_day)
        except Exception as e:
            print(f"[WARN] weekend drivers error: {e}")

    print(f"[Спрос] маршрут {route_number}: 1 смена={len(slots_s1)}, 2 смена={len(slots_s2)}")
    print(f"[Табель] доступно: 1 смена={len(candidates_s1)}, 2 смена={len(candidates_s2)}")

    # ---------------- ПАПКА И ФАЙЛ ----------------
    route_dir = os.path.join(OUTPUT_DIR, str(route_number))
    os.makedirs(route_dir, exist_ok=True)

    output_file = os.path.join(route_dir, f"Расписание_Итог_{target_day}.xlsx")

    # template sheet name (тот, из которого мы ранее получали slots)
    template_sheet_name = sheet_name  # sheet_name взят выше из SCHEDULE_SHEETS

    # ЕСЛИ ФАЙЛА НЕТ — СОЗДАЁМ КОПИЮ ТОЛЬКО НУЖНОГО ЛИСТА И ПЕРЕИМЕНОВЫВАЕМ ЕГО В "Расписание"
    if not os.path.exists(output_file):
        wb_template = load_workbook(FILE_PATH)
        if template_sheet_name not in wb_template.sheetnames:
            print(f"Шаблонный лист '{template_sheet_name}' не найден в {FILE_PATH}")
            return

        # удаляем все листы кроме нужного
        for name in wb_template.sheetnames[:]:
            if name != template_sheet_name:
                del wb_template[name]

        # переименовываем оставшийся лист в "Расписание"
        ws_temp = wb_template[template_sheet_name]
        ws_temp.title = "Расписание"

        wb_template.save(output_file)

    # ОТКРЫВАЕМ ФАЙЛ РАСПИСАНИЯ (уже с единственным листом "Расписание")
    wb = load_workbook(output_file)
    ws = wb["Расписание"]

    assigned_today = set()
    assigned_s1, assigned_s2 = [], []

    # ================== 1 СМЕНА ==================
    for slot in slots_s1:
        slot_start = slot['time_info']['start_dt']
        chosen = choose_driver_for_slot(
            candidates_s1, history, slot_start, 1, assigned_today
        )

        if not chosen and ALLOW_WEEKEND_EXTRA_WORK:
            chosen = choose_driver_for_slot(
                weekend_pool, history, slot_start, 1, assigned_today
            )

        if not chosen:
            ws.cell(row=slot['excel_row'], column=COL_SHIFT_1_INSERT).value = "НЕТ_РЕЗЕРВА"
            continue

        ws.cell(row=slot['excel_row'], column=COL_SHIFT_1_INSERT).value = chosen
        info = slot['time_info'].copy()
        info['shift_code'] = 1
        today_history_log[str(chosen)] = info

        assigned_today.add(str(chosen))
        if chosen in candidates_s1:
            candidates_s1.remove(chosen)
        if chosen in candidates_s2:
            candidates_s2.remove(chosen)
        assigned_s1.append(chosen)

    # ================== 2 СМЕНА ==================
    for slot in slots_s2:
        slot_start = slot['time_info']['start_dt']
        chosen = choose_driver_for_slot(
            candidates_s2, history, slot_start, 2, assigned_today
        )

        if not chosen and ALLOW_WEEKEND_EXTRA_WORK:
            chosen = choose_driver_for_slot(
                weekend_pool, history, slot_start, 2, assigned_today
            )

        if not chosen:
            ws.cell(row=slot['excel_row'], column=COL_SHIFT_2_INSERT).value = "НЕТ_РЕЗЕРВА"
            continue

        ws.cell(row=slot['excel_row'], column=COL_SHIFT_2_INSERT).value = chosen
        info = slot['time_info'].copy()
        info['shift_code'] = 2
        today_history_log[str(chosen)] = info

        assigned_today.add(str(chosen))
        if chosen in candidates_s2:
            candidates_s2.remove(chosen)
        if chosen in candidates_s1:
            candidates_s1.remove(chosen)
        assigned_s2.append(chosen)

    print(f"[Итог] маршрут {route_number}: назначено 1см={len(assigned_s1)}, 2см={len(assigned_s2)}")

    wb.save(output_file)
    save_history(target_day, today_history_log)




if __name__ == "__main__":
    # при запуске скриптом: пробегаем по всем дням в config
    from structure_model.config import TOTAL_DAYS_IN_MONTH
    for day in range(1, TOTAL_DAYS_IN_MONTH + 1):
        prev_day = day - 1
        run_planner(day, prev_day)
