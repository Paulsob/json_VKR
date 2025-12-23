import os
import pandas as pd
from structure_model.config import OUTPUT_DIR


def input_absent_drivers():
    print("\n=== ВВОД ДАННЫХ ОБ ОТСУТСТВУЮЩИХ ВОДИТЕЛЯХ ===")
    print("Формат ввода: <таб_номер>,<смена>,<день>,<причина>")
    print("Пример: 105,1,12,1")
    print("Причины: 0 — отпуск, 1 — больничный, 2 — не предупредил")
    print("Для завершения введите: finish\n")

    absences = []
    seen_drivers = set()  # для подсчёта уникальных

    while True:
        user_input = input(">>> ").strip()
        if user_input.lower() == "finish":
            break

        parts = user_input.split(',')
        if len(parts) != 4:
            print("❌ Неверный формат. Повторите ввод.")
            continue

        try:
            tab_no = str(parts[0]).strip()
            shift = int(parts[1])
            day = int(parts[2])
            reason_code = int(parts[3])

            if shift not in (1, 2):
                print("⚠️ Смена должна быть 1 или 2.")
                continue
            if not (1 <= day <= 30):
                print("⚠️ День должен быть от 1 до 30.")
                continue
            if reason_code not in (0, 1, 2):
                print("⚠️ Причина: 0, 1 или 2.")
                continue

            absences.append({
                'tab_no': tab_no,
                'shift': shift,
                'day': day,
                'reason_code': reason_code
            })
            seen_drivers.add(tab_no)

        except ValueError:
            print("❌ Некорректные данные. Используйте числа для смены, дня и причины.")
            continue

    # Получаем число водителей из отчёта
    report_path = os.path.join(OUTPUT_DIR, "Отчет_Нагрузки_Дни_1_по_30.xlsx")
    if not os.path.exists(report_path):
        print(f"❌ Файл отчёта не найден: {report_path}")
        base_count = 0
    else:
        df = pd.read_excel(report_path, index_col=0)
        base_count = len(df)

    additional_count = len(seen_drivers)
    total_count = base_count + additional_count

    # Вывод
    print("\n" + "=" * 50)
    print(f"Водителей в отчёте: {base_count}")
    print(f"Дополнительно введено отсутствующих (уникальных): {additional_count}")
    print(f"Общее число водителей: {total_count}")
    print("=" * 50)

    # Сохраняем введённые данные в файл (опционально)
    if absences:
        import json
        absence_file = os.path.join(OUTPUT_DIR, "absent_drivers.json")
        with open(absence_file, 'w', encoding='utf-8') as f:
            json.dump(absences, f, indent=2, ensure_ascii=False)
        print(f"\n📝 Данные об отсутствующих сохранены в: {absence_file}")

    return absences, total_count
