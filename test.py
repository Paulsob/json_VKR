from structure_model.driver_scheduler import run_planner
from structure_model.config import TOTAL_DAYS_IN_MONTH, SCHEDULE_SHEETS

routes = sorted({route for route, _ in SCHEDULE_SHEETS.keys()})

print("Маршруты:", routes)

for day in range(1, TOTAL_DAYS_IN_MONTH + 1):
    prev_day = day - 1
    print(f"\n========== ДЕНЬ {day} ==========")

    for route in routes:
        run_planner(day, prev_day, route)
