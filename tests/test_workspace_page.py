# -*- coding: utf-8 -*-
"""Виджет «Рабочее место»: отметки, основная кнопка, сигналы."""

import os
import subprocess
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from pages.workspace_page import (COL_CHECK, COL_STEP0, WorkspacePage,
                                  aggregate_status, parse_point_numbers)


def _row(point, *steps, state="", kind="info"):
    return {"point": point, "folder": f"C:/ПНОС/п.{point}",
            "steps": {i + 1: s for i, s in enumerate(steps)},
            "state": state, "state_kind": kind}


ROWS = [
    _row("2758", *["done"] * 6, "idle"),
    _row("3186", "done", "done", "wait", *["idle"] * 4, state="Ждёт формы", kind="warn"),
    _row("3189", "error", *["idle"] * 6, state="Не корректный", kind="error"),
]


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


@pytest.fixture
def page(app):
    p = WorkspacePage()
    p.set_points(ROWS)
    return p


def _check(page, *points):
    for i, r in enumerate(ROWS):
        if r["point"] in points:
            page.table.item(i, COL_CHECK).setCheckState(Qt.Checked)


def test_parse_point_numbers_keeps_order_without_duplicates():
    assert parse_point_numbers("п.3186, 2758;3186\n2759 п. 2758") == ["3186", "2758", "2759"]
    assert parse_point_numbers("нет номеров") == []


def test_aggregate_status():
    assert aggregate_status([]) == "idle"
    assert aggregate_status(["done", "done"]) == "done"
    assert aggregate_status(["done", "idle"]) == "idle"
    assert aggregate_status(["done", "wait", "error"]) == "error"
    assert aggregate_status(["error", "run"]) == "run"


def test_no_selection_explains_itself(page):
    assert page.selected_points() == []
    assert not page.btn_run.isEnabled()
    assert "Отметьте" in page.btn_run.text()


def test_button_shows_earliest_unfinished_step(page):
    _check(page, "2758")
    assert page.btn_run.text().startswith("Шаг 7 ")
    _check(page, "3186")
    assert page.btn_run.text().startswith("Шаг 3 ")
    assert page.btn_run.isEnabled()


def test_checks_survive_refresh_by_number(page):
    _check(page, "3186")
    page.set_points(list(reversed(ROWS)))
    assert page.selected_points() == ["3186"]
    page.set_points(ROWS[:1])
    assert page.selected_points() == []


def test_check_all_toggles(page):
    page.chk_all.click()
    assert page.selected_points() == ["2758", "3186", "3189"]
    page.chk_all.click()
    assert page.selected_points() == []


def test_run_step_signal(page):
    got = []
    page.run_step_requested.connect(lambda step, pts: got.append((step, pts)))
    _check(page, "2758", "3189")
    page.btn_run.click()
    assert got == [(1, ["2758", "3189"])]


def test_busy_blocks_launch_and_enables_stop(page):
    _check(page, "2758")
    page.set_busy(True)
    assert not page.btn_run.isEnabled()
    assert not page.btn_to_end.isEnabled()
    assert page.btn_stop.isEnabled()
    page.set_busy(False)
    assert page.btn_run.isEnabled()
    assert not page.btn_stop.isEnabled()


def test_run_status_enables_stop(page):
    page.set_run_status(5, "2759", 40, "01:02")
    assert page.btn_stop.isEnabled()
    assert page.lbl_run.fullText() == "Шаг 5 · п.2759"
    page.set_run_status(None, None, 0, "")
    assert not page.btn_stop.isEnabled()


def test_add_points_signal(page):
    got = []
    page.add_points_requested.connect(got.append)
    page.input_points.setText("3186, 3186 2760")
    page.btn_add.click()
    assert got == [["3186", "2760"]]
    assert page.input_points.text() == ""


def test_form_cell_opens_form(page):
    got = []
    page.open_form_requested.connect(got.append)
    page.table.cellClicked.emit(1, COL_STEP0 + 2)
    assert got == ["3186"]


def test_log_line_format(page):
    page.add_log(5, "слияние", "warn")
    page.add_log(0, "общее", "nonsense")
    assert [page.log_list.item(i).text() for i in range(2)] == ["Шаг 5 · слияние", "общее"]
    assert page.lbl_log_last.fullText() == "общее"


def test_empty_table_shows_hint(page):
    page.set_points([])
    assert page.table_stack.currentIndex() == 1


def test_import_is_light():
    """Страница грузится при старте — тяжёлые модули ей не нужны."""
    code = ("import sys; import pages.workspace_page; "
            "heavy = [m for m in ('pandas', 'aiohttp', 'win32com', 'docx', 'openpyxl') "
            "if m in sys.modules]; print(heavy)")
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = subprocess.run([sys.executable, "-c", code], cwd=root,
                         capture_output=True, text=True, check=True).stdout
    assert out.strip() == "[]"
