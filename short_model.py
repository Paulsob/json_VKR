import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

# --- 1. ВХОДНЫЕ ДАННЫЕ (Имитация загрузки файлов) ---

tabel_csv_content = """Таб.№,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15
25160,В,В,1,1,1,1,В,В,1,1,1,1,В,В,1
25180,В,В,2,2,2,2,В,В,2,2,2,2,В,В,2
22719,1,В,В,1,1,1,1,В,В,1,1,1,1,В,В,1
20549,2,В,В,2,2,2,2,В,В,2,2,2,2,В,В,2
25925,1,1,В,В,1,1,1,1,В,В,1,1,1,1,В,В
26761,2,2,В,В,2,2,2,2,В,В,2,2,2,2,В,В
25158,1,1,1,В,В,1,1,1,1,В,В,1,1,1,1,В
23277,2,2,2,В,В,2,2,2,2,В,В,2,2,2,2,В
25184,1,1,1,1,В,В,1,1,1,1,В,В,1,1,1,1
25243,2,2,2,2,В,В,2,2,2,2,В,В,2,2,2,2
"""
# Примечание: Я сократил данные табеля для примера,
# программа будет работать с полным файлом так же.

schedule_csv_content = """Наряд,Start1,End1,Start2,End2
1,04:39,11:23,16:10,01:18
2,04:46,12:13,12:13,21:22
3,04:48,11:29,15:53,00:12
4,04:50,13:02,16:37,00:49
5,04:54,13:10,16:28,00:41
6,04:57,11:37,16:56,01:07
7,04:58,12:32,12:32,21:40
8,05:02,12:28,15:43,00:02
9,05:05,11:46,16:09,00:28
10,05:07,12:48,12:48,21:56
"""


# --- 2. ВСПОМОГАТЕЛЬНЫЕ КЛАССЫ И ФУНКЦИИ ---

class TramScheduler:
    def __init__(self, start_date_str="2023-11-01"):
        self.base_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        # Словарь состояния водителей: {driver_id: datetime_of_last_shift_end}
        # Изначально ставим дату в прошлом, чтобы они были доступны сразу
        self.driver_status = {}
        self.final_schedule = []

    def parse_time(self, date_obj, time_str, is_next_day_check=False):
        """Парсит строку времени HH:MM в datetime объект.
        Учитывает переход через полночь."""
        if pd.isna(time_str) or time_str == '':
            return None

        try:
            h, m = map(int, time_str.split(':'))
            dt = date_obj.replace(hour=h, minute=m, second=0, microsecond=0)

            # Если это проверка окончания смены и время меньше начала (переход через сутки)
            # Например, смена 16:00 - 01:00. 01:00 < 16:00
            if is_next_day_check:
                dt += timedelta(days=1)
            return dt
        except:
            return None

    def load_data(self, tabel_csv, schedule_csv):
        # Загрузка табеля
        self.df_tabel = pd.read_csv(io.StringIO(tabel_csv), sep=',')
        # Очистка имен колонок (удаление пробелов)
        self.df_tabel.columns = [c.strip() for c in self.df_tabel.columns]

        # Загрузка расписания
        # Упрощаем структуру для парсинга (в реальном файле заголовки сложнее)
        self.df_schedule = pd.read_csv(io.StringIO(schedule_csv), sep=',')

        # Преобразуем расписание в список рейсов (Slots)
        self.daily_slots = []
        for _, row in self.df_schedule.iterrows():
            # Смена 1
            if pd.notna(row['Start1']):
                self.daily_slots.append({
                    'naryad': row['Наряд'],
                    'shift_type': '1',  # Тип смены как строка для совпадения с табелем
                    'start_time_str': row['Start1'],
                    'end_time_str': row['End1']
                })
            # Смена 2
            if pd.notna(row['Start2']):
                self.daily_slots.append({
                    'naryad': row['Наряд'],
                    'shift_type': '2',
                    'start_time_str': row['Start2'],
                    'end_time_str': row['End2']
                })

    def run_assignment(self, days_to_process=10):
        """
        Основной цикл распределения по дням.
        """
        # Сортируем слоты внутри дня по времени начала
        self.daily_slots.sort(key=lambda x: x['start_time_str'])

        for day_num in range(1, days_to_process + 1):
            current_date = self.base_date + timedelta(days=day_num - 1)
            day_col = str(day_num)

            if day_col not in self.df_tabel.columns:
                print(f"День {day_col} отсутствует в табеле.")
                continue

            print(f"\n--- Обработка даты: {current_date.date()} (День {day_num}) ---")

            # Создаем конкретные временные слоты для этого дня
            todays_tasks = []
            for slot in self.daily_slots:
                start_dt = self.parse_time(current_date, slot['start_time_str'])

                # Логика перехода через полночь для конца смены
                end_dt = self.parse_time(current_date, slot['end_time_str'])
                if end_dt and end_dt < start_dt:
                    end_dt += timedelta(days=1)
                elif end_dt is None:
                    # Если конца нет, считаем +8 часов (как заглушка)
                    end_dt = start_dt + timedelta(hours=8)

                todays_tasks.append({
                    'naryad': slot['naryad'],
                    'shift_type': slot['shift_type'],
                    'start': start_dt,
                    'end': end_dt
                })

            # Сортируем задачи по времени начала
            todays_tasks.sort(key=lambda x: x['start'])

            # Получаем водителей, работающих в этот день
            # Группируем их: '1': [ids...], '2': [ids...]
            available_drivers = {'1': [], '2': []}

            for _, driver in self.df_tabel.iterrows():
                shift_val = str(driver[day_col]).strip()
                driver_id = driver['Таб.№']

                # Если водитель еще не в базе статусов, добавляем его (отдыхал давно)
                if driver_id not in self.driver_status:
                    self.driver_status[driver_id] = self.base_date - timedelta(hours=48)

                if shift_val in ['1', '2']:
                    available_drivers[shift_val].append(driver_id)

            # --- ПРОЦЕСС НАЗНАЧЕНИЯ ---
            assigned_drivers_today = set()

            for task in todays_tasks:
                s_type = task['shift_type']
                needed_start = task['start']

                candidates = available_drivers.get(s_type, [])
                best_driver = None

                # Поиск подходящего кандидата
                # Критерий: (Start_Task - Last_Shift_End) >= 12 часов

                valid_candidates = []
                for d_id in candidates:
                    if d_id in assigned_drivers_today:
                        continue  # Уже назначен сегодня

                    last_end = self.driver_status[d_id]

                    # Проверка 12 часов
                    rest_time = needed_start - last_end
                    if rest_time >= timedelta(hours=12):
                        valid_candidates.append(d_id)

                if valid_candidates:
                    # Берем первого попавшегося (можно усложнить: того, кто отдыхал дольше всех)
                    best_driver = valid_candidates[0]

                    # НАЗНАЧЕНИЕ
                    self.final_schedule.append({
                        'Date': current_date.date(),
                        'DriverID': best_driver,
                        'Naryad': task['naryad'],
                        'ShiftType': s_type,
                        'Start': task['start'].strftime('%H:%M'),
                        'End': task['end'].strftime('%H:%M %d-%m')
                    })

                    # Обновляем статус водителя (когда он освободится)
                    self.driver_status[best_driver] = task['end']
                    assigned_drivers_today.add(best_driver)
                    print(f"Назначен: {best_driver} на Наряд {task['naryad']} ({task['start'].strftime('%H:%M')})")
                else:
                    print(
                        f"НЕ НАЙДЕН ВОДИТЕЛЬ: Наряд {task['naryad']} Смена {s_type} (Старт {task['start'].strftime('%H:%M')})")

    def get_schedule_df(self):
        return pd.DataFrame(self.final_schedule)


# --- 3. ЗАПУСК ---

# Инициализация
scheduler = TramScheduler()
scheduler.load_data(tabel_csv_content, schedule_csv_content)

# Запуск расчета на первые 5 дней
scheduler.run_assignment(days_to_process=5)

# Вывод результата
res = scheduler.get_schedule_df()
print("\n--- ИТОГОВОЕ РАСПРЕДЕЛЕНИЕ (Фрагмент) ---")
print(res.head(10).to_string())