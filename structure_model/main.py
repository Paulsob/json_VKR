from structure_model.config import TOTAL_DAYS_IN_MONTH, FILE_PATH, HISTORY_JSON_DIR
from structure_model.driver_scheduler import run_planner, run_planner_for_all_routes
from structure_model.report_generator import generate_work_summary
import os
from structure_model.summary_report import write_summary_statistics
from structure_model.absence_input import input_absent_drivers

def auto_run_simulation(total_days, file_path):
    print("--- Удаление старых логов ---")
    for d in range(1, total_days + 2):
        f = os.path.join(HISTORY_JSON_DIR, f"history_{d}.json")
        if os.path.exists(f):
            os.remove(f)

    for day in range(1, total_days + 1):
        print(f"ЗАПУСК ДНЯ {day:02d}")
        run_planner_for_all_routes(day, day - 1)

    print("\n##################### СИМУЛЯЦИЯ ЗАВЕРШЕНА #####################")
    generate_work_summary(total_days, file_path)


if __name__ == "__main__":
    auto_run_simulation(TOTAL_DAYS_IN_MONTH, FILE_PATH)
    write_summary_statistics()
    input_absent_drivers()