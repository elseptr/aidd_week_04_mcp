#!/usr/bin/env python3
"""Формирует статистический Excel-отчёт формата 1.1 по задачам из todo_db.json.

Источник данных: apps/todo/todo_db.json (Flask-приложение, поле payload.tasks);
тот же JSON возвращает GET /api/tasks этого приложения.

Примеры:
    python make_report.py --input apps/todo/todo_db.json --output results/tasks_report.xlsx
    python make_report.py --input todo_db.json --output report.xlsx --today 2026-11-04

Результат — один .xlsx с листами «Задачи», «Сводка», «Определения»:
статические значения (без формул Excel), даты — как настоящие значения.
"""

import argparse
import json
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

# practice/agent/.qwen/skills/make_report -> practice/agent
_SKILL_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT = _SKILL_DIR.parents[3] / "results" / "tasks_report.xlsx"

PRIORITY_RU = {"low": "Низкий", "normal": "Средний", "high": "Высокий"}
STATUS_RU = {True: "Выполнена", False: "Не выполнена"}
PRIORITY_ORDER = ("low", "normal", "high")
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")


def load_tasks(path):
    """Загрузка: читает JSON-файл и возвращает список задач (list[dict])."""
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ValueError(f"не удалось прочитать {path}: {exc}") from exc
    except UnicodeDecodeError as exc:
        raise ValueError(f"неверная кодировка файла {path}: {exc}") from exc
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"повреждённый JSON в {path}: строка {exc.lineno}, колонка {exc.colno}: {exc.msg}"
        ) from exc
    if isinstance(data, dict):
        payload = data.get("payload")
        data = payload.get("tasks") if isinstance(payload, dict) else None
    if not isinstance(data, list):
        raise ValueError(
            "в JSON нет списка задач: ожидается {\"payload\": {\"tasks\": [...]}} "
            "либо массив задач на верхнем уровне"
        )
    return data


def parse_task(raw, index):
    """Валидация: превращает сырую запись в dict задачи или выбрасывает ValueError."""
    if not isinstance(raw, dict):
        raise ValueError(f"запись #{index} не является объектом JSON")

    task_id = raw.get("id")
    if isinstance(task_id, bool) or not isinstance(task_id, int):
        raise ValueError(f"у задачи #{index} поле 'id' отсутствует или имеет неверный тип")

    title = raw.get("title")
    if not isinstance(title, str):
        raise ValueError(f"у задачи #{task_id} поле 'title' отсутствует или имеет неверный тип")

    done = raw.get("done")
    if not isinstance(done, bool):
        raise ValueError(f"у задачи #{task_id} поле 'done' должно быть булевым (true/false)")

    priority = raw.get("priority")
    if priority not in PRIORITY_RU:
        raise ValueError(
            f"у задачи #{task_id} поле 'priority' = {priority!r}; допустимы low, normal, high"
        )

    due = raw.get("due")
    if due is None:
        due_date = None
    elif isinstance(due, str):
        try:
            due_date = date.fromisoformat(due)
        except ValueError:
            raise ValueError(
                f"у задачи #{task_id} поле 'due' имеет неверный формат: {due!r} "
                f"(ожидается YYYY-MM-DD)"
            ) from None
    else:
        raise ValueError(f"у задачи #{task_id} поле 'due' должно быть строкой YYYY-MM-DD или null")

    created = raw.get("created_at")
    if not isinstance(created, str):
        raise ValueError(
            f"у задачи #{task_id} поле 'created_at' отсутствует или имеет неверный тип"
        )
    try:
        created_dt = datetime.fromisoformat(created)
    except ValueError:
        raise ValueError(
            f"у задачи #{task_id} поле 'created_at' имеет неверный формат: {created!r} "
            f"(ожидается YYYY-MM-DDTHH:MM:SS)"
        ) from None

    return {
        "id": task_id,
        "title": title,
        "done": done,
        "priority": priority,
        "due": due_date,
        "created_at": created_dt,
    }


def is_overdue(task, today):
    return not task["done"] and task["due"] is not None and task["due"] < today


def days_to_due(task, today):
    """Расчёт статистики: (срок - сегодня) в днях; None для выполненных и без срока."""
    if task["done"] or task["due"] is None:
        return None
    return (task["due"] - today).days


def _ratio(numerator, denominator):
    return None if denominator == 0 else numerator / denominator


def _average(values):
    return None if not values else sum(values) / len(values)


_METRIC_DEFINITIONS = [
    ("Всего задач", "все строки списка"),
    ("Выполнено", "задачи со статусом «Выполнена»"),
    ("Не выполнено", "всего минус выполнено"),
    ("Со сроком", "строки, где срок не пустой"),
    ("Без срока", "всего минус со сроком"),
    ("Просрочено", "срок < сегодня И статус «Не выполнена»"),
    ("% просрочки", "просрочено / со сроком"),
    ("Соблюдённые сроки", "со сроком минус просрочено"),
    ("Срочные сейчас", "приоритет «Высокий» И статус «Не выполнена»"),
    ("Выполнено, низкий приоритет", "выполненные со статусом «Выполнена» и приоритетом «Низкий»"),
    ("Выполнено, средний приоритет", "выполненные со статусом «Выполнена» и приоритетом «Средний»"),
    ("Выполнено, высокий приоритет", "выполненные со статусом «Выполнена» и приоритетом «Высокий»"),
    ("Всего, низкий приоритет", "все задачи с приоритетом «Низкий»"),
    ("Всего, средний приоритет", "все задачи с приоритетом «Средний»"),
    ("Всего, высокий приоритет", "все задачи с приоритетом «Высокий»"),
    ("% выполнения по приоритету",
     "выполнено / всего по каждому приоритету: B17 — низкий, C17 — средний, D17 — высокий"),
    ("Создано сегодня", "дата создания = сегодня (сравнение по дате из created_at)"),
    ("Средняя просрочка, дн.", "среднее |Дней до срока| по просроченным (Дней до срока < 0)"),
    ("Среднее время до срока, дн.",
     "среднее Дней до срока по незавершённым со сроком (Дней до срока > 0)"),
]


def build_workbook(tasks, today):
    """Формирование XLSX: собирает Workbook с листами «Задачи», «Сводка», «Определения»."""
    wb = Workbook()

    ws_tasks = wb.active
    ws_tasks.title = "Задачи"
    ws_tasks.append(["ID", "Название", "Статус", "Приоритет", "Срок", "Создана",
                     "Просрочена", "Дней до срока"])
    for task in sorted(tasks, key=lambda t: t["id"]):
        ws_tasks.append([
            task["id"],
            task["title"],
            STATUS_RU[task["done"]],
            PRIORITY_RU[task["priority"]],
            task["due"],
            task["created_at"],
            "да" if is_overdue(task, today) else "нет",
            days_to_due(task, today),
        ])

    ws_summary = wb.create_sheet("Сводка")
    _write_summary(ws_summary, tasks, today)
    ws_definitions = wb.create_sheet("Определения")
    _write_definitions(ws_definitions, today)
    _style_workbook(wb)
    return wb


def _write_summary(ws, tasks, today):
    ws.append(["Показатель", "Значение"])

    total = len(tasks)
    done_count = sum(1 for t in tasks if t["done"])
    with_due = [t for t in tasks if t["due"] is not None]
    overdue = [t for t in with_due if not t["done"] and t["due"] < today]
    h_values = [days_to_due(t, today) for t in tasks]
    overdue_days = [abs(h) for h in h_values if h is not None and h < 0]
    future_days = [h for h in h_values if h is not None and h > 0]
    created_today = sum(1 for t in tasks if t["created_at"].date() == today)

    by_priority = {priority: {"total": 0, "done": 0} for priority in PRIORITY_ORDER}
    for t in tasks:
        by_priority[t["priority"]]["total"] += 1
        if t["done"]:
            by_priority[t["priority"]]["done"] += 1

    rows = [
        ("Всего задач", total),
        ("Выполнено", done_count),
        ("Не выполнено", total - done_count),
        ("Со сроком", len(with_due)),
        ("Без срока", total - len(with_due)),
        ("Просрочено", len(overdue)),
        ("% просрочки", _ratio(len(overdue), len(with_due))),
        ("Соблюдённые сроки", len(with_due) - len(overdue)),
        ("Срочные сейчас", by_priority["high"]["total"] - by_priority["high"]["done"]),
        ("Выполнено, низкий приоритет", by_priority["low"]["done"]),
        ("Выполнено, средний приоритет", by_priority["normal"]["done"]),
        ("Выполнено, высокий приоритет", by_priority["high"]["done"]),
        ("Всего, низкий приоритет", by_priority["low"]["total"]),
        ("Всего, средний приоритет", by_priority["normal"]["total"]),
        ("Всего, высокий приоритет", by_priority["high"]["total"]),
        ("% выполнения по приоритету", [
            _ratio(by_priority["low"]["done"], by_priority["low"]["total"]),
            _ratio(by_priority["normal"]["done"], by_priority["normal"]["total"]),
            _ratio(by_priority["high"]["done"], by_priority["high"]["total"]),
        ]),
        ("Создано сегодня", created_today),
        ("Средняя просрочка, дн.", _average(overdue_days)),
        ("Среднее время до срока, дн.", _average(future_days)),
    ]

    for offset, (label, value) in enumerate(rows):
        row = 2 + offset
        ws.cell(row=row, column=1, value=label)
        if isinstance(value, list):
            for column, item in enumerate(value, start=2):
                ws.cell(row=row, column=column, value=item)
        else:
            ws.cell(row=row, column=2, value=value)


def _write_definitions(ws, today):
    ws.append(["Определения"])
    ws.append(["Словарь статусов:"])
    ws.append(["Выполнена", "done = true"])
    ws.append(["Не выполнена", "done = false"])
    ws.append(["Словарь приоритетов:"])
    ws.append(["Низкий", "priority = low"])
    ws.append(["Средний", "priority = normal"])
    ws.append(["Высокий", "priority = high"])
    ws.append(["Метрики листа «Сводка»:"])
    for label, definition in _METRIC_DEFINITIONS:
        ws.append([label, definition])
    ws.append(["Дата экспорта", today.isoformat()])


def _style_workbook(wb):
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws = wb["Задачи"]
    last_row = ws.max_row
    for cell in ws[1]:
        cell.font = bold
        cell.fill = HEADER_FILL
        cell.alignment = center
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        for cell in row:
            cell.number_format = "yyyy-mm-dd"
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
        for cell in row:
            cell.number_format = "yyyy-mm-dd hh:mm"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{last_row}"
    for column, width in {"A": 6, "B": 40, "C": 14, "D": 12, "E": 12,
                          "F": 19, "G": 11, "H": 15}.items():
        ws.column_dimensions[column].width = width

    ws = wb["Сводка"]
    for cell in ws[1]:
        cell.font = bold
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 14
    ws["B8"].number_format = "0.00%"
    for column in ("B", "C", "D"):
        ws[f"{column}17"].number_format = "0.00%"

    ws = wb["Определения"]
    for row in (1, 2, 5, 9):
        ws.cell(row=row, column=1).font = bold
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 58


def _parse_today(value):
    if value is None:
        return date.today()
    try:
        return date.fromisoformat(value)
    except ValueError:
        raise ValueError(f"--today={value!r}: ожидается дата в формате YYYY-MM-DD") from None


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="make_report.py",
        description="Формирует Excel-отчёт формата 1.1 по задачам из todo_db.json",
    )
    parser.add_argument(
        "--input", required=True,
        help="путь к todo_db.json (Flask-приложение apps/todo)",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT),
        help=f"путь сохранения .xlsx (по умолчанию: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--today",
        default=None,
        help="дата расчёта YYYY-MM-DD; по умолчанию — текущая дата",
    )
    args = parser.parse_args(argv)

    try:
        today = _parse_today(args.today)
        raw_tasks = load_tasks(Path(args.input))
    except (ValueError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    tasks = []
    for index, raw in enumerate(raw_tasks, start=1):
        try:
            tasks.append(parse_task(raw, index))
        except ValueError as exc:
            print(f"error: {exc}", file=sys.stderr)
            return 1

    output_path = Path(args.output)
    tmp_path = output_path.with_name(output_path.name + ".tmp")
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook = build_workbook(tasks, today)
        workbook.save(tmp_path)
        tmp_path.replace(output_path)
    except OSError as exc:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)
        print(f"error: не удалось записать отчёт {output_path}: {exc}", file=sys.stderr)
        return 1

    print(f"Отчёт сохранён: {output_path} ({len(tasks)} задач)")
    return 0


if __name__ == "__main__":
    sys.exit(main())