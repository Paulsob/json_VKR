import os
import json
from structure_model.config import HISTORY_JSON_DIR
from datetime import datetime

def load_history(day_num):
    filename = os.path.join(HISTORY_JSON_DIR, f"history_{day_num}.json")
    if not os.path.exists(filename):
        return {}
    with open(filename, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_history(day_num, data):
    filename = os.path.join(HISTORY_JSON_DIR, f"history_{day_num}.json")
    clean_data = {
        k: {key: val for key, val in v.items() if not isinstance(val, datetime)}
        for k, v in data.items()
    }
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(clean_data, f, indent=4, ensure_ascii=False)
    print(f"[Память] Данные о сменах сохранены в {filename}")