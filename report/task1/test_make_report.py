"""Тесты для make_report.py: расчёты, формат XLSX, CLI и обработка ошибок.

Запуск: python -m pytest -v practice/agent/.qwen/skills/make_report/tests/
"""

import json
import subprocess
import sys
from datetime import date, datetime, time
from pathlib import Path

import openpyxl
import pytest

_SKILL_DIR = Path(__file__).resolve().parent.parent  # каталог skill (make_report)
if str(_SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(_SKILL_DIR))

import make_report

TODAY = date(2026, 11, 4)

# Тестовый набор из 8 задач: выполненные и не выполненные, все приоритеты,
# просроченные задачи, срок сегодня, задачи без срока.
# Ожидаемые результаты посчитаны вручную от даты 2026-11-04.
FIXTURE_TASKS = [
    {"id": 1, "title": "Выполненная, срок прошёл", "done": True, "priority": "high",
     "due": "2026-10-28", "created_at": "2026-11-01T09:00:00"},
    {"id": 2, "title": "Выполненная, без срока", "done": True, "priority": "normal",
     "due": None, "created_at": "2026-11-02T10:00:00"},
    {"id": 3, "title": "Не выполнена, срок сегодня", "done": False, "priority": "low",
     "due": "2026-11-04", "created_at": "2026-11-04T08:00:00"},
    {"id": 4, "title": "Не выполнена, просрочена", "done": False, "priority": "high",
     "due": "2026-10-25", "created_at": "2026-11-01T14:30:00"},
    {"id": 5, "title": "Не выполнена, срок впереди", "done": False, "priority": "normal",
     "due": "2026-11-18", "created_at": "2026-11-03T11:20:00"},
    {"id": 6, "title": "Выполненная, низкий приоритет", "done": True, "priority": "low",
     "due": "2026-11-20", "created_at": "2026-10-30T16:00:00"},
    {"id": 7, "title": "Не выполнена, средний приоритет", "done": False, "priority": "normal",
     "due": "2026-12-01", "created_at": "2026-11-04T09:15:00"},
    {"id": 8, "title": "Выполненная, срок сегодня", "done": True, "priority": "high",
     "due": "2026-11-04", "created_at": "2026-11-04T10:00:00"},
]

# Ожидаемые метрики листа «Сводка» (дата расчёта 2026-11-04)
FIXTURE_METRICS = {
    "Всего задач": 8,
    "Выполнено": 4,
    "Не выполнено": 4,
    "Со сроком": 7,
    "Без срока": 1,
    "Просрочено": 1,
    "% просрочки": 1 / 7,
    "Соблюдённые сроки": 6,
    "Срочные сейчас": 1,
    "Выполнено, низкий приоритет": 1,
    "Выполнено, средний приоритет": 1,
    "Выполнено, высокий приоритет": 2,
    "Всего, низкий приоритет": 2,
    "Всего, средний приоритет": 3,
    "Всего, высокий приоритет": 3,
    "% выполнения по приоритету": [0.5, 1 / 3, 2 / 3],
    "Создано сегодня": 3,
    "Средняя просрочка, дн.": 10.0,
    "Среднее время до срока, дн.": 20.5,
}

METRIC_LABELS = [
    "Всего задач", "Выполнено", "Не выполнено", "Со сроком", "Без срока",
    "Просрочено", "% просрочки", "Соблюдённые сроки", "Срочные сейчас",
    "Выполнено, низкий приоритет", "Выполнено, средний приоритет",
    "Выполнено, высокий приоритет", "Всего, низкий приоритет",
    "Всего, средний приоритет", "Всего, высокий приоритет",
    "% выполнения по приоритету", "Создано сегодня",
    "Средняя просрочка, дн.", "Среднее время до срока, дн.",
]

HEADERS = ["ID", "Название", "Статус", "Приоритет", "Срок", "Создана",
           "Просрочена", "Дней до срока"]
SHEET_NAMES = ["Задачи", "Сводка", "Определения"]

# (название, статус, приоритет, срок, создана, просрочена, дней до срока)
EXPECTED_ROWS = {
    1: ("Выполненная, срок прошёл", "Выполнена", "Высокий", date(2026, 10, 28),
        datetime(2026, 11, 1, 9, 0), "нет", None),
    2: ("Выполненная, без срока", "Выполнена", "Средний", None,
        datetime(2026, 11, 2, 10, 0), "нет", None),
    3: ("Не выполнена, срок сегодня", "Не выполнена", "Низкий", date(2026, 11, 4),
        datetime(2026, 11, 4, 8, 0), "нет", 0),
    4: ("Не выполнена, просрочена", "Не выполнена", "Высокий", date(2026, 10, 25),
        datetime(2026, 11, 1, 14, 30), "да", -10),
    5: ("Не выполнена, срок впереди", "Не выполнена", "Средний", date(2026, 11, 18),
        datetime(2026, 11, 3, 11, 20), "нет", 14),
    6: ("Выполненная, низкий приоритет", "Выполнена", "Низкий", date(2026, 11, 20),
        datetime(2026, 10, 30, 16, 0), "нет", None),
    7: ("Не выполнена, средний приоритет", "Не выполнена", "Средний", date(2026, 12, 1),
        datetime(2026, 11, 4, 9, 15), "нет", 27),
    8: ("Выполненная, срок сегодня", "Выполнена", "Высокий", date(2026, 11, 4),
        datetime(2026, 11, 4, 10, 0), "нет", None),
}


def write_fixture(path, tasks):
    path.write_text(json.dumps(tasks, ensure_ascii=False), encoding="utf-8")


def run_cli(*args):
    return subprocess.run(
        [sys.executable, str(_SKILL_DIR / "make_report.py"), *args],
        capture_output=True, text=True, timeout=60,
    )


def read_report(path):
    return openpyxl.load_workbook(path)


def test_all_metrics(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    out = tmp_path / "report.xlsx"
    rc = run_cli("--input", str(src), "--output", str(out), "--today", "2026-11-04")
    assert rc.returncode == 0, rc.stderr
    wb = read_report(out)
    ws = wb["Сводка"]
    assert ws["A1"].value == "Показатель" and ws["B1"].value == "Значение"
    for offset, label in enumerate(METRIC_LABELS):
        assert ws.cell(row=2 + offset, column=1).value == label, f"row {2 + offset}"
    for label, expected in FIXTURE_METRICS.items():
        row = METRIC_LABELS.index(label) + 2
        if isinstance(expected, list):
            for col, exp in enumerate(expected, start=2):
                got = ws.cell(row=row, column=col).value
                assert got == pytest.approx(exp), f"{label} [{col}] got {got!r}"
        else:
            got = ws.cell(row=row, column=2).value
            assert got == pytest.approx(expected), f"{label} got {got!r}"


def test_task_rows_sorted_by_id(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, list(reversed(FIXTURE_TASKS)))
    out = tmp_path / "report.xlsx"
    assert run_cli("--input", str(src), "--output", str(out),
                   "--today", "2026-11-04").returncode == 0
    ws = read_report(out)["Задачи"]
    assert [ws.cell(row=r, column=1).value for r in range(2, 10)] == list(range(1, 9))
    for task_id, (title, status, priority, due, created, overdue, days) in EXPECTED_ROWS.items():
        row = task_id + 1
        due_value = datetime.combine(due, time.min) if due is not None else None
        expected = [title, status, priority, due_value, created, overdue, days]
        got = [ws.cell(row=row, column=c).value for c in range(2, 9)]
        assert got == expected, f"task {task_id}"


def test_overdue_and_days_rules(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, [
        {"id": 1, "title": "Готово, срок прошёл", "done": True, "priority": "high",
         "due": "2026-10-01", "created_at": "2026-11-01T09:00:00"},
        {"id": 2, "title": "Срок сегодня", "done": False, "priority": "normal",
         "due": "2026-11-04", "created_at": "2026-11-04T09:00:00"},
        {"id": 3, "title": "Нет срока", "done": False, "priority": "low",
         "due": None, "created_at": "2026-11-04T09:00:00"},
    ])
    out = tmp_path / "report.xlsx"
    assert run_cli("--input", str(src), "--output", str(out),
                   "--today", "2026-11-04").returncode == 0
    ws = read_report(out)["Задачи"]
    assert ws["G2"].value == "нет" and ws["H2"].value is None   # выполненная с прошедшим сроком
    assert ws["G3"].value == "нет" and ws["H3"].value == 0      # срок сегодня: не просрочена, 0 дн.
    assert ws["G4"].value == "нет" and ws["H4"].value is None   # без срока
    sw = read_report(out)["Сводка"]
    assert sw["B5"].value == 2 and sw["B6"].value == 1          # со сроком / без срока
    assert sw["B7"].value == 0                                  # просрочено
    assert sw["B19"].value is None and sw["B20"].value is None  # нет ни отрицательных, ни положительных H


def test_sheets_and_headers(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    out = tmp_path / "report.xlsx"
    assert run_cli("--input", str(src), "--output", str(out),
                   "--today", "2026-11-04").returncode == 0
    wb = read_report(out)
    assert wb.sheetnames == SHEET_NAMES
    ws = wb["Задачи"]
    assert [ws.cell(row=1, column=c).value for c in range(1, 9)] == HEADERS
    assert ws.max_row == len(FIXTURE_TASKS) + 1


def test_date_cells_are_real_values(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    out = tmp_path / "report.xlsx"
    assert run_cli("--input", str(src), "--output", str(out),
                   "--today", "2026-11-04").returncode == 0
    ws = read_report(out)["Задачи"]
    e4 = ws["E5"].value      # задача 4: срок 2026-10-25
    f4 = ws["F5"].value      # created_at 2026-11-01T14:30:00
    assert isinstance(e4, date) and not isinstance(e4, str)
    assert isinstance(f4, datetime)
    assert ws["E3"].value is None   # задача 2 без срока
    assert ws["H2"].value is None   # выполненная (id 1)
    assert ws["H3"].value is None   # без срока (id 2)


def test_empty_cells_are_truly_empty(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    out = tmp_path / "report.xlsx"
    assert run_cli("--input", str(src), "--output", str(out),
                   "--today", "2026-11-04").returncode == 0
    ws = read_report(out)["Задачи"]
    placeholders = ("None", "null", "-")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            assert cell.value not in placeholders, f"заглушка в {cell.coordinate}"


def test_percent_format(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    out = tmp_path / "report.xlsx"
    assert run_cli("--input", str(src), "--output", str(out),
                   "--today", "2026-11-04").returncode == 0
    wb = read_report(out)
    sw = wb["Сводка"]
    assert sw["B8"].number_format == "0.00%"
    for column in ("B", "C", "D"):
        assert sw[f"{column}17"].number_format == "0.00%"
    assert sw["B7"].value == 1 and sw["B5"].value == 7
    assert sw["B8"].value == pytest.approx(1 / 7)


def test_division_by_zero_gives_empty(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, [{
        "id": 1, "title": "Без срока", "done": False, "priority": "low",
        "due": None, "created_at": "2026-11-04T09:00:00",
    }])
    out = tmp_path / "report.xlsx"
    assert run_cli("--input", str(src), "--output", str(out),
                   "--today", "2026-11-04").returncode == 0
    sw = read_report(out)["Сводка"]
    assert sw["B5"].value == 0        # со сроком: 0
    assert sw["B8"].value is None     # % просрочки: 0/0 -> пустая ячейка
    assert sw["B19"].value is None    # средняя просрочка: нет отрицательных H
    assert sw["B20"].value is None    # среднее время до срока: нет положительных H
    assert sw["B17"].value == 0.0     # 0/1 = 0 — не заглушка


def test_formatting(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    out = tmp_path / "report.xlsx"
    assert run_cli("--input", str(src), "--output", str(out),
                   "--today", "2026-11-04").returncode == 0
    wb = read_report(out)
    ws = wb["Задачи"]
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref == f"A1:H{len(FIXTURE_TASKS) + 1}"
    assert ws["A1"].font.b is True
    assert ws["E2"].number_format == "yyyy-mm-dd"
    assert ws["F2"].number_format == "yyyy-mm-dd hh:mm"
    sw = wb["Сводка"]
    assert sw["A1"].font.b is True
    ow = wb["Определения"]
    assert ow["A1"].value == "Определения"
    definitions = {ow.cell(row=r, column=1).value: ow.cell(row=r, column=2).value
                   for r in range(1, ow.max_row + 1)}
    assert definitions["Выполнена"] == "done = true"
    assert definitions["Высокий"] == "priority = high"
    assert ("Дата экспорта", "2026-11-04") in definitions.items()


def test_today_flag_affects_metrics(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, [{
        "id": 1, "title": "Срок завтра", "done": False, "priority": "normal",
        "due": "2026-11-05", "created_at": "2026-11-01T09:00:00",
    }])
    def run(today):
        out = tmp_path / f"report_{today}.xlsx"
        rc = run_cli("--input", str(src), "--output", str(out), "--today", today)
        return rc, read_report(out)
    rc_earlier, wb_earlier = run("2026-11-04")
    assert rc_earlier.returncode == 0
    assert wb_earlier["Задачи"]["H2"].value == 1
    assert wb_earlier["Сводка"]["B7"].value == 0
    rc_later, wb_later = run("2026-11-06")
    assert rc_later.returncode == 0
    assert wb_later["Задачи"]["H2"].value == -1
    assert wb_later["Сводка"]["B7"].value == 1


def test_cli_success_and_help(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    rc = run_cli("--input", str(src), "--output", str(tmp_path / "report.xlsx"),
                 "--today", "2026-11-04")
    assert rc.returncode == 0
    assert "Отчёт сохранён" in rc.stdout
    rc_help = run_cli("--help")
    assert rc_help.returncode == 0
    assert "usage" in rc_help.stdout.lower()


def test_missing_input_file(tmp_path):
    out = tmp_path / "report.xlsx"
    rc = run_cli("--input", str(tmp_path / "no-such-file.json"),
                 "--output", str(out), "--today", "2026-11-04")
    assert rc.returncode != 0
    assert not out.exists()
    assert rc.stderr


def test_empty_json(tmp_path):
    src = tmp_path / "empty.json"
    src.write_text("", encoding="utf-8")
    out = tmp_path / "report.xlsx"
    rc = run_cli("--input", str(src), "--output", str(out), "--today", "2026-11-04")
    assert rc.returncode != 0
    assert not out.exists()


def test_corrupted_json(tmp_path):
    src = tmp_path / "broken.json"
    src.write_text("{broken json", encoding="utf-8")
    out = tmp_path / "report.xlsx"
    rc = run_cli("--input", str(src), "--output", str(out), "--today", "2026-11-04")
    assert rc.returncode != 0
    assert not out.exists()


def test_invalid_task_data(tmp_path):
    cases = [
        [{"id": 1, "title": "Ок", "done": False, "priority": "ultramax",
          "due": None, "created_at": "2026-11-04T09:00:00"}],        # неизвестный приоритет
        [{"id": 1, "title": "Ок", "done": "да", "priority": "low",
          "due": None, "created_at": "2026-11-04T09:00:00"}],        # неверный тип done
        [{"id": 1, "done": False, "priority": "low", "due": None,
          "created_at": "2026-11-04T09:00:00"}],                    # отсутствует title
        [{"title": "Без id", "done": False, "priority": "low", "due": None,
          "created_at": "2026-11-04T09:00:00"}],                    # отсутствует id
        [{"id": 1, "title": "Ок", "done": False, "priority": "low",
          "due": "2026/11/04", "created_at": "2026-11-04T09:00:00"}],  # неверный формат даты
    ]
    for idx, tasks in enumerate(cases):
        src = tmp_path / f"case_{idx}.json"
        write_fixture(src, tasks)
        out = tmp_path / f"report_{idx}.xlsx"
        rc = run_cli("--input", str(src), "--output", str(out), "--today", "2026-11-04")
        assert rc.returncode != 0, f"case {idx}"
        assert not out.exists(), f"case {idx}: отчёт создан при ошибке"


def test_created_at_without_time_accepted(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, [{"id": 1, "title": "Без времени", "done": False, "priority": "low",
                         "due": None, "created_at": "2026-11-04"}])
    out = tmp_path / "report.xlsx"
    rc = run_cli("--input", str(src), "--output", str(out), "--today", "2026-11-04")
    assert rc.returncode == 0, rc.stderr
    f1 = read_report(out)["Задачи"]["F2"].value
    assert isinstance(f1, datetime)
    assert f1.hour == 0


def test_invalid_today_flag(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    out = tmp_path / "report.xlsx"
    rc = run_cli("--input", str(src), "--output", str(out), "--today", "04.11.2026")
    assert rc.returncode != 0
    assert not out.exists()


def test_bad_arguments(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    out = tmp_path / "report.xlsx"
    rc = run_cli("--input", str(src), "--output", str(out), "--nope")
    assert rc.returncode != 0
    assert not out.exists()


def test_auto_create_output_directory(tmp_path):
    src = tmp_path / "todo_db.json"
    write_fixture(src, FIXTURE_TASKS)
    out = tmp_path / "sub" / "deep" / "report.xlsx"
    assert run_cli("--input", str(src), "--output", str(out),
                   "--today", "2026-11-04").returncode == 0
    assert out.exists()
    assert read_report(out).sheetnames == SHEET_NAMES


def test_no_partial_report_on_error(tmp_path):
    out = tmp_path / "report.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("старые данные, не xlsx", encoding="utf-8")
    rc = run_cli("--input", str(tmp_path / "no-such.json"), "--output", str(out),
                 "--today", "2026-11-04")
    assert rc.returncode != 0
    assert out.read_text(encoding="utf-8") == "старые данные, не xlsx"