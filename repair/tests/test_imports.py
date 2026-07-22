# -*- coding: utf-8 -*-
"""
Тесты слоя листа «Импорт»: чтение, подбор строк, диф записи.
"""

import openpyxl
import pytest

from repair.imports import (
    read_import_sections, rows_for_numbers, build_write_ops, apply_ops_com,
    WriteOp, COL_THICK, COL_DIAM, COL_T1, COL_TMIN, SHEET_IMPORT,
)
from repair.model import Section
from repair.nominals import apply_nominals


# ─── Чтение ───────────────────────────────────────────────────────────

def _make_import_xlsx(path, rows):
    """rows: списки значений A…L."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_IMPORT
    ws.append(["Номер", "Сечение", "Элемент", "Тисп", "Диам",
               "Т1", "Т2", "Т3", "Т4", "Т5", "Т6", "Тмин"])
    for r in rows:
        ws.append(r)
    wb.save(path)


def test_read_sections(tmp_path):
    path = str(tmp_path / "т.xlsx")
    _make_import_xlsx(path, [
        [1, 1, "Отвод", 12, 219, 11.2, 11.2, 11.4, 11.3, 11.2, 11.3, 11.2],
        [2, 2, "Труба", 10, 219, 10.1, 10.1, 10.3, 10.1, None, None, 10.1],
        [3, 6, "Задвижка", None, None, 17.9, None, None, None, None, None, 17.9],
        [None, None, None, None, None, None, None, None, None, None, None, None],
    ])
    secs = read_import_sections(path)
    assert len(secs) == 3
    a, b, z = secs
    assert (a.number, a.row, a.element) == (1, 2, "Отвод")
    assert a.measures == ["11.2", "11.2", "11.4", "11.3", "11.2", "11.3"]
    assert a.thickness_nom == "12"
    assert b.measures == ["10.1", "10.1", "10.3", "10.1"]   # хвостовые пустые срезаны
    assert z.element == "Задвижка"
    assert z.measures == ["17.9"]


def test_read_real_2701_workbook():
    """Чтение реального ПНОС_ТТП.xlsm (до ремонта) — только чтение."""
    path = r"C:\Users\volod\Desktop\п.2701\ПНОС_ТТП.xlsm"
    import os
    if not os.path.exists(path):
        pytest.skip("реального файла нет на этой машине")
    secs = read_import_sections(path)
    assert len(secs) > 100
    by_num = {}
    for s in secs:
        by_num.setdefault(s.number, []).append(s)
    # известная коллизия: задвижка №6 и отвод №6
    six = by_num.get(6, [])
    assert len(six) == 2
    assert {s.element for s in six} == {"Задвижка", "Отвод"}


# ─── Подбор строк ─────────────────────────────────────────────────────

def _secs():
    return [
        Section(6, "Задвижка", "", "", ["17.9"], "17.9", row=7),
        Section(6, "Отвод", "10", "219", ["9.1", "9.1"], "9.1", row=8),
        Section(7, "Труба", "10", "219", ["7.7"], "7.7", row=9),
        Section(115, "Бобышка", "", "32", ["5.8"], "5.8", row=10),
    ]


def test_rows_collision_valve_skipped():
    rows, warns = rows_for_numbers(_secs(), {6: 10.0})
    assert rows == {8: 10.0}                    # только отвод, задвижка мимо
    assert any("Задвижка" in w for w in warns)


def test_rows_small_element_skipped():
    rows, warns = rows_for_numbers(_secs(), {115: 6.0})
    assert rows == {}                           # бобышка по умолчанию не трогается
    assert any("115" in w for w in warns)


def test_rows_missing_number_warned():
    rows, warns = rows_for_numbers(_secs(), {999: 5.0})
    assert rows == {}
    assert any("999" in w for w in warns)


def test_rows_plain():
    rows, warns = rows_for_numbers(_secs(), {7: 8.0})
    assert rows == {9: 8.0}
    assert warns == []


# ─── Диф записи ───────────────────────────────────────────────────────

def test_build_write_ops_only_changed():
    old = [Section(7, "Труба", "10", "219", ["9.1", "9.2", "", "9.1"], "9.1", row=9)]
    new, _ = apply_nominals(old, {7: 8.0})
    ops = build_write_ops(old, new)

    got = {(o.row, o.col): o.value for o in ops}
    assert got[(9, COL_THICK)] == 8.0           # Тисп
    assert got[(9, COL_T1)] == 8.0              # т.1
    assert got[(9, COL_T1 + 1)] == 8.0          # т.2
    assert (9, COL_T1 + 2) not in got           # пустая точка не дорисована
    assert got[(9, COL_T1 + 3)] == 8.0          # т.4
    assert got[(9, COL_TMIN)] == 8.0
    assert (9, COL_DIAM) not in got             # диаметр не менялся


def test_build_write_ops_perehod_text():
    old = [Section(11, "Переход", "11 / 12", "273 / 219",
                   ["9.9", "10.1"], "9.9", row=12)]
    new, _ = apply_nominals(old, {11: 12.0}, {11: ("12 / 10", "273 / 219")})
    ops = build_write_ops(old, new)
    got = {(o.row, o.col): o.value for o in ops}
    assert got[(12, COL_THICK)] == "12 / 10"    # текст остаётся текстом
    assert (12, COL_DIAM) not in got


def test_build_write_ops_untouched_section_empty():
    old = [Section(7, "Труба", "10", "219", ["9.1"], "9.1", row=9)]
    new, _ = apply_nominals(old, {})
    assert build_write_ops(old, new) == []


# ─── Применение к COM ─────────────────────────────────────────────────

class _FakeWS:
    def __init__(self):
        self.written = {}
    def Cells(self, r, c):
        ws = self
        class _C:
            def __setattr__(self, name, value):
                if name == "Value":
                    ws.written[(r, c)] = value
                object.__setattr__(self, name, value)
        return _C()


def test_apply_ops_com():
    ws = _FakeWS()
    n = apply_ops_com(ws, [WriteOp(9, COL_THICK, 8.0), WriteOp(9, COL_T1, 8.0)])
    assert n == 2
    assert ws.written == {(9, COL_THICK): 8.0, (9, COL_T1): 8.0}


# ─── Живой путь: применение по строкам ────────────────────────────────

def test_apply_by_row_ignores_same_number_other_row():
    """Два сечения №6: номинал применяется только к выбранной строке."""
    secs = _secs()
    rows, _ = rows_for_numbers(secs, {6: 10.0})     # → {8: 10.0}, задвижка мимо
    new, changes = apply_nominals(secs, rows, by_row=True)

    by_row = {s.row: s for s in new}
    assert by_row[8].measures == ["10.0", "10.0"]   # отвод №6 обновлён
    assert by_row[7].measures == ["17.9"]           # задвижка №6 не тронута
    assert len(changes) == 1
