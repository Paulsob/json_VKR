import pandas as pd
import os
from structure_model.config import OUTPUT_DIR

def write_summary_statistics():
    report_path = os.path.join(OUTPUT_DIR, "Отчет_Нагрузки_Дни_1_по_30.xlsx")
    if not os.path.exists(report_path):
        print(f"[Ошибка] Файл отчёта не найден: {report_path}")
        return

    # Загружаем отчёт (первый лист)
    df = pd.read_excel(report_path, index_col=0)

    # Общее число водителей = число строк
    total_drivers_with_schedule = len(df)

    # 21.7% от этого числа — это количество "отсутствующих" (по вашему условию)
    absent_drivers = round(total_drivers_with_schedule * 0.217)

    # Общее число водителей (включая "отсутствующих")
    total_drivers = total_drivers_with_schedule + absent_drivers

    # Формируем текст отчёта
    report_text = (
        f"Количество водителей по таблице 'Отчет_Нагрузки_Дни_1_по_30.xlsx': {total_drivers_with_schedule}\n"
        f"Количество водителей, отсутствующих по статистике (21,7%): {absent_drivers}\n"
        f"Общее количество водителей (включая отсутствующих): {total_drivers}\n"
    )

    # Сохраняем в файл
    output_file = os.path.join(OUTPUT_DIR, "summary_report.txt")
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(report_text)

    print(f"\n[Отчёт] Краткая статистика сохранена в: {output_file}")
    print(report_text)