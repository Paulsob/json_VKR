import pandas as pd
import json
from pathlib import Path

# ----------------------------
# Пути
# ----------------------------

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_FILE = BASE_DIR / "data" / "Журнал_закреплений_ТМ5.xlsx"
OUTPUT_DIR = BASE_DIR / "consolidation"

OUTPUT_DIR.mkdir(exist_ok=True)

# ----------------------------
# Загрузка Excel
# ----------------------------

xls = pd.ExcelFile(DATA_FILE, engine="openpyxl")

# Сюда собираем данные по маршрутам
routes_data = {}

# ----------------------------
# Парсинг листов (со 2-го)
# ----------------------------

for sheet_name in xls.sheet_names[1:]:
    df = pd.read_excel(
        xls,
        sheet_name=sheet_name,
        header=None,
        engine="openpyxl"
    )

    # Номер маршрута (H1)
    route_number_raw = df.iloc[0, 7]

    if pd.isna(route_number_raw):
        continue

    route_number = int(route_number_raw)

    # Поиск строки заголовков
    header_row = None
    for i in range(len(df)):
        if str(df.iloc[i, 0]).strip() == "Таб. номер":
            header_row = i
            break

    if header_row is None:
        continue

    # Инициализация маршрута
    routes_data.setdefault(route_number, [])

    # Чтение данных
    for i in range(header_row + 1, len(df)):
        first_cell = str(df.iloc[i, 0]).strip()

        # Стоп-слово
        if first_cell == "Итого":
            break

        tab_number = df.iloc[i, 0]

        if pd.isna(tab_number):
            continue

        # ФИО: столбцы 2–16
        fio_parts = df.iloc[i, 1:16].dropna().astype(str)
        fio = " ".join(fio_parts).strip()

        routes_data[route_number].append({
            "tab_number": int(tab_number),
            "employee": fio,
            "sheet": sheet_name
        })

# ----------------------------
# Сохранение в JSON
# ----------------------------

for route_number, employees in routes_data.items():
    route_dir = OUTPUT_DIR / str(route_number)
    route_dir.mkdir(parents=True, exist_ok=True)

    output_file = route_dir / "data.json"

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(
            {
                "route": route_number,
                "employees": employees
            },
            f,
            ensure_ascii=False,
            indent=4
        )

print("Готово. Данные сохранены в папке consolidation/")
