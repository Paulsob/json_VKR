from flask import Flask, render_template, request, jsonify, send_file
import os
import json
import pandas as pd
import datetime
from pathlib import Path

from structure_model.config import (
    TRANSPORT_CONFIGS,
    BASE_DIR,
    TOTAL_DAYS_IN_MONTH,
    FILE_PATH,
    COL_SHIFT_1_INSERT,
    COL_SHIFT_2_INSERT,
    ABSENCES_FILE
)
from structure_model.driver_scheduler import run_planner as run_planner_for_day, get_all_routes

app = Flask(__name__)
app.config['SECRET_KEY'] = 'json-only-version'

# =========================================================
# Вспомогательные функции
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
# DASHBOARD
# =========================================================
@app.route('/')
def dashboard():
    report_path = BASE_DIR / "output" / "Отчет_Нагрузки_Дни_1_по_30.xlsx"
    base_count = 0

    if report_path.exists():
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
# КАЛЕНДАРЬ
# =========================================================
@app.route('/calendar-data')
def calendar_data():
    days_with_schedules = set()

    # Обходим все output-каталоги по транспортам
    for cfg in TRANSPORT_CONFIGS.values():
        output_root: Path = Path(cfg["output_dir"])
        if not output_root.exists():
            continue

        for entry in os.listdir(output_root):
            route_path = output_root / entry
            if not route_path.is_dir():
                continue
            for day in range(1, TOTAL_DAYS_IN_MONTH + 1):
                fname = f"Расписание_Итог_{day}.xlsx"
                if (route_path / fname).exists():
                    days_with_schedules.add(day)

    return jsonify(sorted(days_with_schedules))


# =========================================================
# ОТСУТСТВИЯ
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

    # Пересчёт для всех маршрутов всех типов (можно оптимизировать)
    for transport, cfg in TRANSPORT_CONFIGS.items():
        sheets = cfg["sheets"]
        routes = sorted({k[0] for k in sheets.keys()})
        for route_number in routes:
            run_planner_for_day(day, prev_day, route_number, transport)

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
    for transport, cfg in TRANSPORT_CONFIGS.items():
        sheets = cfg["sheets"]
        routes = sorted({k[0] for k in sheets.keys()})
        for route_number in routes:
            run_planner_for_day(day, prev_day, route_number, transport)

    return jsonify({"success": True})


@app.route('/get-real-absences')
def get_real_absences():
    absences = load_absences()
    return [
        {**item, "id": idx}
        for idx, item in enumerate(absences)
    ]


# =========================================================
# ПЕРЕСЧЁТ
# =========================================================
@app.route('/api/recalculate/<int:day>', methods=['POST'])
def api_recalculate(day):
    if day < 1 or day > TOTAL_DAYS_IN_MONTH:
        return jsonify({'error': 'Некорректный день'}), 400

    data = request.get_json(silent=True) or {}
    route = data.get("route")
    transport = data.get("transport")  # optional: "bus" | "obus" | "tram"

    prev_day = max(day - 1, 0)

    if route is None:
        # пересчёт для всех маршрутов (возможно ограничить транспорт)
        if transport:
            sheets = TRANSPORT_CONFIGS.get(transport, {}).get("sheets", {})
            for route_number in sorted({k[0] for k in sheets.keys()}):
                run_planner_for_day(day, prev_day, route_number, transport)
        else:
            for t, cfg in TRANSPORT_CONFIGS.items():
                sheets = cfg["sheets"]
                for route_number in sorted({k[0] for k in sheets.keys()}):
                    run_planner_for_day(day, prev_day, route_number, t)
    else:
        # пересчёт конкретного маршрута (с учётом типа транспорта)
        route_number = int(route)
        if transport:
            run_planner_for_day(day, prev_day, route_number, transport)
        else:
            # пытаемся угадать транспорт по SCHEDULE_SHEETS
            found = False
            for t, cfg in TRANSPORT_CONFIGS.items():
                if (route_number, True) in cfg["sheets"] or (route_number, False) in cfg["sheets"]:
                    run_planner_for_day(day, prev_day, route_number, t)
                    found = True
            if not found:
                # fallback: запустить для всех (маловероятно)
                for t in TRANSPORT_CONFIGS.keys():
                    run_planner_for_day(day, prev_day, route_number, t)

    return jsonify({"success": True})


# =========================================================
# МАРШРУТЫ
# =========================================================
@app.route('/api/routes')
def api_routes():
    """
    Возвращает список маршрутов с указанием типа транспорта:
    [{ "route": 55, "transport": "bus" }, ...]
    """
    out = []
    for transport, cfg in TRANSPORT_CONFIGS.items():
        for (r, _) in cfg["sheets"].keys():
            out.append({"route": int(r), "transport": transport})
    # сортируем по номеру и по транспорту
    out = sorted(out, key=lambda x: (x["route"], x["transport"]))
    return jsonify(out)


# =========================================================
# РАСПИСАНИЕ
# =========================================================
@app.route('/api/schedule/<int:day>/<int:route>')
def api_schedule(day, route):
    if day < 1 or day > TOTAL_DAYS_IN_MONTH:
        return jsonify({'error': 'Некорректный день'}), 400

    # optional query param ?transport=obus
    transport = request.args.get("transport", None)

    # определяем транспорт (если не задан)
    selected_transport = None
    if transport:
        selected_transport = transport if transport in TRANSPORT_CONFIGS else None
    else:
        for t, cfg in TRANSPORT_CONFIGS.items():
            if (route, True) in cfg["sheets"] or (route, False) in cfg["sheets"]:
                selected_transport = t
                break

    if not selected_transport:
        return jsonify({'error': f'Маршрут {route} не найден ни для одного транспорта'}), 400

    cfg = TRANSPORT_CONFIGS[selected_transport]
    sheets = cfg["sheets"]
    output_root: Path = Path(cfg["output_dir"])

    # weekend detection
    current_date = datetime.date.today().replace(day=day)
    is_weekend = current_date.weekday() >= 5

    sheet_name = sheets.get((route, is_weekend))
    if not sheet_name:
        return jsonify({'error': f'Маршрут {route} не найден для транспорта {selected_transport}'}), 400

    # ---------- читаем базовое расписание ----------
    if is_weekend:
        filename = f"Расписание_выходного_дня_{route}.xlsx"
    else:
        filename = f"Расписание_рабочего_дня_{route}.xlsx"

    filepath = output_root / filename
    if not filepath.exists():
        return jsonify({'error': f'Файл {filename} не найден в {output_root}'}), 404

    df = pd.read_excel(filepath, sheet_name="Лист1", header=None)

    # ---------- читаем итог (водителей) ----------
    drivers_s1 = {}
    drivers_s2 = {}

    itog_path = output_root / str(route) / f"Расписание_Итог_{day}.xlsx"

    if itog_path.exists():
        try:
            itog_df = pd.read_excel(itog_path, sheet_name="Расписание", header=None)

            col1 = COL_SHIFT_1_INSERT - 1
            col2 = COL_SHIFT_2_INSERT - 1

            for idx in range(len(itog_df)):
                if col1 < itog_df.shape[1]:
                    v1 = itog_df.iat[idx, col1]
                    if v1 and str(v1).strip() != "НЕТ_РЕЗЕРВА" and str(v1).strip() != "НЕТ":
                        drivers_s1[idx] = str(v1).strip()

                if col2 < itog_df.shape[1]:
                    v2 = itog_df.iat[idx, col2]
                    if v2 and str(v2).strip() != "НЕТ_РЕЗЕРВА" and str(v2).strip() != "НЕТ":
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
        "transport": selected_transport,
        "is_weekend": is_weekend,
        "rows": rows,
        "columns": [
            "Номер маршрута",
            "Отправление 1 смена",
            "Прибытие 1 смена",
            "Водитель 1 смена",
            "Отправление 2 смена",
            "Прибытие 2 смена",
            "Водитель 2 смена",
        ]
    })


# =========================================================
# ОТЧЁТ
# =========================================================
@app.route('/get-report')
def get_report():
    # Берём корень проекта относительно этого файла
    path = BASE_DIR / "output" / "Отчет_Нагрузки_Дни_1_по_30.xlsx"

    print("Полный путь к файлу:", path)
    print("Файл существует?", path.exists())

    if not path.exists():
        return "Отчет не найден", 404

    return send_file(path, as_attachment=True)


# =========================================================
if __name__ == "__main__":
    app.run(debug=True, port=5000)
