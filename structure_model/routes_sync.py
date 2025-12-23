# structure_model/routes_sync.py
import os
import re
from datetime import datetime

import pandas as pd

from structure_model import config
from .extensions import db
from structure_model.models import Route
from flask import current_app

WORK_RE = re.compile(r'Расписание_рабочего_дня[_\s]?(\d+)$', flags=re.IGNORECASE)
WEEK_RE = re.compile(r'Расписание_выходного_дня[_\s]?(\d+)$', flags=re.IGNORECASE)


def build_mapping_from_sheet_names(sheet_names):
    """
    На вход: список имён листов (строк).
    Возвращает dict:
    { route_number: {'sheet_name_workday':..., 'sheet_name_weekend':...} }
    """
    found = {}
    for name in sheet_names:
        n = name.strip()

        m = WORK_RE.search(n)
        if m:
            num = int(m.group(1))
            entry = found.setdefault(num, {})
            entry['sheet_name_workday'] = n
            continue

        m = WEEK_RE.search(n)
        if m:
            num = int(m.group(1))
            entry = found.setdefault(num, {})
            entry['sheet_name_weekend'] = n

    return found


def get_sheet_names_from_excel(file_path):
    """Читает workbook и возвращает список имён листов (без загрузки данных)."""
    try:
        xl = pd.ExcelFile(file_path)
        return xl.sheet_names
    except Exception:
        return []


def sync_routes(sheet_names=None, excel_file=None, output_dir=None, commit_missing_files=False):
    """
    Синхронизирует routes в БД и обновляет config.SCHEDULE_SHEETS.

    ВАЖНО:
    - НЕ вызывает ensure_db_created()
    - Предполагает, что app_context уже существует
    """
    output_dir = output_dir or config.OUTPUT_DIR

    if sheet_names is None:
        excel_file = excel_file or config.FILE_PATH
        sheet_names = get_sheet_names_from_excel(excel_file)

    mapping = build_mapping_from_sheet_names(sheet_names)

    results = {
        'added': [],
        'updated': [],
        'skipped': [],
        'mapping_count': len(mapping),
    }

    # app_context должен быть активен снаружи
    for num, info in sorted(mapping.items()):
        sheet_w = info.get('sheet_name_workday') or f'Расписание_рабочего_дня_{num}'
        sheet_q = info.get('sheet_name_weekend') or f'Расписание_выходного_дня_{num}'

        # обновляем глобальный config (legacy) — удобно для совместимости с ранним кодом
        config.SCHEDULE_SHEETS[(num, False)] = sheet_w
        config.SCHEDULE_SHEETS[(num, True)] = sheet_q

        expected_workfile = f"Расписание_рабочего_дня_{num}.xlsx"
        expected_weekfile = f"Расписание_выходного_дня_{num}.xlsx"

        work_exists = os.path.exists(os.path.join(output_dir, expected_workfile))
        week_exists = os.path.exists(os.path.join(output_dir, expected_weekfile))

        existing = Route.query.filter_by(route_number=num).first()

        if existing:
            changed = False

            if existing.sheet_name_workday != sheet_w:
                existing.sheet_name_workday = sheet_w
                changed = True

            if existing.sheet_name_weekend != sheet_q:
                existing.sheet_name_weekend = sheet_q
                changed = True

            if work_exists or commit_missing_files:
                if existing.file_workday != expected_workfile:
                    existing.file_workday = expected_workfile
                    changed = True

            if week_exists or commit_missing_files:
                if existing.file_weekend != expected_weekfile:
                    existing.file_weekend = expected_weekfile
                    changed = True

            if changed:
                existing.updated_at = datetime.utcnow()
                db.session.add(existing)
                results['updated'].append(num)
            else:
                results['skipped'].append(num)

        else:
            route = Route(
                route_number=num,
                name=f'Маршрут {num}',
                sheet_name_workday=sheet_w,
                sheet_name_weekend=sheet_q,
                file_workday=(expected_workfile if (work_exists or commit_missing_files) else None),
                file_weekend=(expected_weekfile if (week_exists or commit_missing_files) else None),
                is_active=True
            )
            db.session.add(route)
            results['added'].append(num)

    db.session.commit()

    results['final_schedule_sheets_keys'] = list(config.SCHEDULE_SHEETS.keys())
    return results
