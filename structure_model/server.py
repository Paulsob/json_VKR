from flask import Flask, render_template, request, jsonify, send_from_directory
import os
import json
import pandas as pd
import datetime

from structure_model.config import (
    OUTPUT_DIR,
    HISTORY_JSON_DIR,
    TOTAL_DAYS_IN_MONTH,
    FILE_PATH,
    TAB_SHEET_NAME,
    COL_SHIFT_1_INSERT,
    COL_SHIFT_2_INSERT,
    SCHEDULE_SHEETS,
)
from structure_model.driver_scheduler import run_planner as run_planner_for_day

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

app = Flask(__name__)
app.config['SECRET_KEY'] = 'json-only-version'

ABSENCES_FILE = os.path.join(OUTPUT_DIR, "real_absences.json")

# =========================================================
# =============== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==================
# =========================================================

def load_absences():
    if not os.path.exists(ABSENCES_FILE):
        return []
    try:
        with open(ABSENCES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_absences(data):
    with open(ABSENCES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# =========================================================
# ======================= DASHBOARD =======================
# =========================================================

@app.route('/')
def dashboard():
    report_path = os.path.join(OUTPUT_DIR, "Отчет_Нагрузки_Дни_1_по_30.xlsx")
    base_count = 0

    if os.path.exists(report_path):
        df = pd.read_excel(report_path, index_col=0)
        base_count = len(df)

    real_absences = load_absences()
    unique_absent_drivers = len(set(item['tab_no'] for item in real_absences))
    statistical_absent = round(base_count * 0.217)

    return render_template(
        'dashboard.html',
        base_count=base_count,
        statistical_absent=statistical_absent,
        real_absent=unique_absent_drivers,
        total_drivers_real=base_count + unique_absent_drivers,
        total_drivers_stat=base_count + statistical_absent
    )


# =========================================================
# ===================== КАЛЕНДАРЬ =========================
# =========================================================

@app.route('/calendar-data')
def calendar_data():
    days_with_schedules = set()

    if not os.path.exists(OUTPUT_DIR):
        return jsonify([])

    for entry in os.listdir(OUTPUT_DIR):
        route_path = os.path.join(OUTPUT_DIR, entry)
        if not os.path.isdir(route_path):
            continue

        for day in range(1, TOTAL_DAYS_IN_MONTH + 1):
            fname = f"Расписание_Итог_{day}.xlsx"
            if os.path.exists(os.path.join(route_path, fname)):
                days_with_schedules.add(day)

    return jsonify(sorted(days_with_schedules))


# =========================================================
# ==================== ОТСУТСТВИЯ =========================
# =========================================================

@app.route('/submit-absence', methods=['POST'])
def submit_absence():
    data = request.json
    absences = load_absences()

    absences.append({
        "tab_no": str(data.get("tab_no")).strip(),
        "shift": int(data.get("shift")),
        "day": int(data.get("day")),
        "reason": str(data.get("reason", "")),
    })

    save_absences(absences)

    day = int(data.get("day"))
    prev_day = max(day - 1, 0)

    for route_number in sorted({k[0] for k in SCHEDULE_SHEETS.keys()}):
        run_planner_for_day(day, prev_day, route_number)

    return jsonify({"success": True})


@app.route('/delete-absence', methods=['POST'])
def delete_absence():
    data = request.json
    index = data.get("id")

    absences = load_absences()
    if index < 0 or index >= len(absences):
        return jsonify({"error": "Запись не найдена"}), 404

    day = absences[index]["day"]
    absences.pop(index)
    save_absences(absences)

    prev_day = max(day - 1, 0)
    for route_number in sorted({k[0] for k in SCHEDULE_SHEETS.keys()}):
        run_planner_for_day(day, prev_day, route_number)

    return jsonify({"success": True})


@app.route('/get-real-absences')
def get_real_absences():
    absences = load_absences()
    return [
        {**item, "id": idx}
        for idx, item in enumerate(absences)
    ]


# =========================================================
# ===================== ПЕРЕСЧЁТ ==========================
# =========================================================

@app.route('/api/recalculate/<int:day>', methods=['POST'])
def api_recalculate(day):
    if day < 1 or day > TOTAL_DAYS_IN_MONTH:
        return jsonify({'error': 'Некорректный день'}), 400

    data = request.get_json(silent=True) or {}
    route = data.get("route")

    prev_day = max(day - 1, 0)

    if route is None:
        run_planner_for_day(day, prev_day)
    else:
        run_planner_for_day(day, prev_day, int(route))

    return jsonify({"success": True})


# =========================================================
# ===================== МАРШРУТЫ ==========================
# =========================================================

@app.route('/api/routes')
def api_routes():
    routes = sorted({k[0] for k in SCHEDULE_SHEETS.keys()})
    return jsonify(routes)


# =========================================================
# ==================== РАСПИСАНИЕ =========================
# =========================================================

@app.route('/api/schedule/<int:day>/<int:route>')
def api_schedule(day, route):
    if day < 1 or day > TOTAL_DAYS_IN_MONTH:
        return jsonify({'error': 'Некорректный день'}), 400

    current_date = datetime.date.today().replace(day=day)
    is_weekend = current_date.weekday() >= 5

    sheet_name = SCHEDULE_SHEETS.get((route, is_weekend))
    if not sheet_name:
        return jsonify({'error': f'Маршрут {route} не найден'}), 400

    # ---------- читаем базовое расписание ----------
    if is_weekend:
        filename = f"Расписание_выходного_дня_{route}.xlsx"
    else:
        filename = f"Расписание_рабочего_дня_{route}.xlsx"

    filepath = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(filepath):
        return jsonify({'error': f'Файл {filename} не найден'}), 404

    df = pd.read_excel(filepath, sheet_name="Лист1", header=None)

    # ---------- читаем итог (водителей) ----------
    drivers_s1 = {}
    drivers_s2 = {}

    itog_path = os.path.join(
        OUTPUT_DIR,
        str(route),
        f"Расписание_Итог_{day}.xlsx"
    )

    if os.path.exists(itog_path):
        try:
            itog_df = pd.read_excel(itog_path, sheet_name="Расписание", header=None)

            col1 = COL_SHIFT_1_INSERT - 1
            col2 = COL_SHIFT_2_INSERT - 1

            for idx in range(len(itog_df)):
                if col1 < itog_df.shape[1]:
                    v1 = itog_df.iat[idx, col1]
                    if v1 and str(v1).strip() != "НЕТ_РЕЗЕРВА":
                        drivers_s1[idx] = str(v1).strip()

                if col2 < itog_df.shape[1]:
                    v2 = itog_df.iat[idx, col2]
                    if v2 and str(v2).strip() != "НЕТ_РЕЗЕРВА":
                        drivers_s2[idx] = str(v2).strip()

        except Exception as e:
            print(f"[WARN] Не удалось прочитать итог: {e}")

    # ---------- собираем строки ----------
    rows = []

    for i in range(3, len(df)):
        base_row = df.iloc[i].fillna('').tolist()

        if any(str(c).strip() for c in base_row):
            base_row.append(drivers_s1.get(i, ""))
            base_row.append(drivers_s2.get(i, ""))
            rows.append(base_row)

    return jsonify({
        "success": True,
        "day": day,
        "route": route,
        "is_weekend": is_weekend,
        "rows": rows,
        "columns": [
            "Номер маршрута",
            "Отправление 1 смена",
            "Прибытие 1 смена",
            "Отправление 2 смена",
            "Прибытие 2 смена",
            "Водитель 1 смена",
            "Водитель 2 смена",
        ]
    })


# =========================================================
# ======================== ОТЧЁТ =========================
# =========================================================

@app.route('/get-report')
def get_report():
    filename = "Отчет_Нагрузки_Дни_1_по_30.xlsx"
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


# =========================================================

if __name__ == "__main__":
    app.run(debug=True, port=5000)
