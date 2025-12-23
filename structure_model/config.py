import os

HISTORY_JSON_DIR = "history_json"
OUTPUT_DIR = "output"
FILE_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data.xlsx")

os.makedirs(HISTORY_JSON_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ROW_START = 5
STEP = 2

COL_SHIFT_1_INSERT = 3
COL_SHIFT_2_INSERT = 8

COL_SHIFT_1_START = 5
COL_SHIFT_1_END = 6
COL_SHIFT_2_START = 10
COL_SHIFT_2_END = 11

TOTAL_DAYS_IN_MONTH = 30

ABSENCES_FILE = os.path.join(OUTPUT_DIR, "real_absences.json")

TAB_SHEET_NAME = "Весь_табель"
REST_HOURS = 12.0

# Номер маршрута, для которого запускается планировщик (9 или 55)
ROUTE_NUMBER = 55

# Код выходного дня в табеле
WEEKEND_CODE = "В"

# Разрешать ли вызывать водителей в их выходной день (за удвоенную плату)
ALLOW_WEEKEND_EXTRA_WORK = False

# Карта: (маршрут, выходной_ли_день) -> имя листа с расписанием в data.xlsx
SCHEDULE_SHEETS = {
    (9, True): "Расписание_выходного_дня_9",
    (9, False): "Расписание_рабочего_дня_9",
    (20, True): "Расписание_выходного_дня_20",
    (20, False): "Расписание_рабочего_дня_20",
    (21, True): "Расписание_выходного_дня_21",
    (21, False): "Расписание_рабочего_дня_21",
    (47, True): "Расписание_выходного_дня_47",
    (47, False): "Расписание_рабочего_дня_47",
    (48, True): "Расписание_выходного_дня_48",
    (48, False): "Расписание_рабочего_дня_48",
    (55, True): "Расписание_выходного_дня_55",
    (55, False): "Расписание_рабочего_дня_55",
    (61, True): "Расписание_выходного_дня_61",
    (61, False): "Расписание_рабочего_дня_61"
}
