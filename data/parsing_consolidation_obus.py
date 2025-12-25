import pandas as pd

input_file = "Журнал_закреплений_ТБ2.xlsx"
output_file = "Журнал_закреплений_ТБ2_обработанный.xlsx"

# читаем всю книгу
xls = pd.ExcelFile(input_file)

result_sheets = {}

for sheet_idx, sheet_name in enumerate(xls.sheet_names):
    # Первый лист не трогаем — копируем как есть
    if sheet_idx == 0:
        df0 = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=object)
        result_sheets[sheet_name] = df0
        continue

    df = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=object)

    current_ts = None
    rows_to_drop = set()

    # проходим сверху вниз, отслеживаем блоки "Маршрут" / "ТС"
    for idx, row in df.iterrows():
        first = row.iloc[0] if len(row) > 0 else None

        # нормализуем строку для проверок
        if isinstance(first, str):
            s = first.strip()

            # удаляем строки заголовков таблицы и итоги
            if s.startswith("Таб. номер") or s.startswith("Итого"):
                rows_to_drop.add(idx)
                continue

            # метка начала блока маршрута
            if s.startswith("Маршрут"):
                rows_to_drop.add(idx)
                # обнуляем текущее ТС — ждём следующую строчку "ТС" для установки нового
                current_ts = None
                continue

            # строка с меткой ТС — читаем значение из 23-го столбца (index 22)
            if s.startswith("ТС"):
                rows_to_drop.add(idx)
                ts_val = None
                if len(row) > 22:
                    ts_val = row.iloc[22]
                # если есть значение — сохраняем как текущий ТС, иначе None
                if pd.notna(ts_val):
                    current_ts = ts_val
                else:
                    current_ts = None
                continue

        # удаляем полностью пустые строки
        if row.isna().all():
            rows_to_drop.add(idx)
            continue

        # если в строке есть ФИО (во 2-м столбце) — считаем это сотрудником
        second = row.iloc[1] if len(row) > 1 else None
        if pd.notna(second) and str(second).strip() != "":
            # если есть текущий ТС — записываем его в 3-й столбец (index 2)
            if current_ts is not None and pd.notna(current_ts):
                df.at[idx, 2] = current_ts
            # иначе оставляем как есть
            continue

        # прочие строки оставляем (на всякий случай) — не помечаем на удаление

    # удаляем помеченные строки, сбросим индексы
    if rows_to_drop:
        df = df.drop(index=sorted(rows_to_drop)).reset_index(drop=True)
    else:
        df = df.reset_index(drop=True)

    # оставляем только первые 4 столбца (0..3)
    df = df.iloc[:, :4]

    # удаляем полностью пустые строки после всех операций
    df = df.dropna(how="all").reset_index(drop=True)

    # добавляем шапку в начало листа
    header = pd.DataFrame([["Таб. №", "ФИО", "ТС", None]])
    df = pd.concat([header, df], ignore_index=True)

    result_sheets[sheet_name] = df

# сохраняем все листы в новый файл
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for name, sheet_df in result_sheets.items():
        sheet_df.to_excel(writer, sheet_name=name, index=False, header=False)

print("Готово — сохранено в:", output_file)
