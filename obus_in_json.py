import pandas as pd
import os
import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
input_file = BASE_DIR / "new_VKR-main/data/Журнал_закреплений_ТБ2_обработанный.xlsx"
output_base = BASE_DIR / "new_VKR-main/consolidation/obus"



# создаём базовую папку, если нет
os.makedirs(output_base, exist_ok=True)

# читаем Excel полностью
xls = pd.ExcelFile(input_file)

for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=object)

    # пропускаем шапку (первая строка)
    df_data = df.iloc[1:, :3]  # берём только первые 3 столбца: Таб. №, ФИО, ТС
    df_data = df_data.dropna(subset=[0,1], how="all")  # удаляем полностью пустые строки

    employees = []
    for _, row in df_data.iterrows():
        tab_number = row[0]
        employee = row[1]
        transport = row[2] if pd.notna(row[2]) else None

        # если нет табельного номера или ФИО — пропускаем строку
        if pd.isna(tab_number) or pd.isna(employee):
            continue

        emp_dict = {
            "tab_number": int(tab_number) if isinstance(tab_number, (int,float)) else str(tab_number),
            "employee": str(employee),
            "transport": str(transport) if transport is not None else ""
        }
        employees.append(emp_dict)

    # формируем итоговый JSON
    route_json = {
        "route": sheet_name,
        "employees": employees
    }

    # создаём папку маршрута
    route_folder = os.path.join(output_base, str(sheet_name))
    os.makedirs(route_folder, exist_ok=True)

    # сохраняем JSON
    json_path = os.path.join(route_folder, "data.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(route_json, f, ensure_ascii=False, indent=4)

    print(f"Сохранили {json_path}")
