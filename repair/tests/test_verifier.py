# -*- coding: utf-8 -*-
"""Тесты верификатора (Э2): R2 строка, R3 зеркальность, R4 #REF!, R5 задвижки."""

import docx
import openpyxl

from repair.model import ActRef, RepairRecord, Section, ID_PRESENT
from repair.verifier import (
    check_svedenia, check_mirror, check_ref_errors, check_valves_in_table3,
)

BASE = ("Исполнительная документация в наличии и приложена "
        "к техническому паспорту на трубопровод")


# ─── R2 ───────────────────────────────────────────────────────────────

def _rec():
    return [RepairRecord(id_status=ID_PRESENT, date_done="10.04.2026",
                         acts=[ActRef("222-1-5", "01.02.2026")])]


def test_svedenia_match():
    want = (BASE + ": 10.04.2026 г. проведена замена участков трубопровода: "
            "01.02.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-5")
    (f,) = check_svedenia(_rec(), want)
    assert f.level == "ok"


def test_svedenia_match_tolerates_spacing():
    """Ручная пунктуация («;» без пробела, точка в конце) — не ошибка."""
    manual = (BASE + ":  10.04.2026 г. проведена замена участков трубопровода: "
              "01.02.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-5.")
    (f,) = check_svedenia(_rec(), manual)
    assert f.level == "ok"


def test_svedenia_mismatch():
    (f,) = check_svedenia(_rec(), BASE)   # в отчёте базовая фраза, а ремонт есть
    assert f.level == "error"
    assert "не соответствует" in f.message


def test_svedenia_absent():
    (f,) = check_svedenia(_rec(), None)
    assert f.level == "warn"


# ─── R3 ───────────────────────────────────────────────────────────────

def _sec(row, number=1, tmin="9.1"):
    return Section(number=number, element="Труба", thickness_nom="10",
                   diameter="219", measures=["9.1"], tmin=tmin, row=row)


def test_mirror_ok():
    (f,) = check_mirror([_sec(2), _sec(3, 2)], [_sec(2), _sec(3, 2)])
    assert f.level == "ok"


def test_mirror_value_diff():
    (f,) = check_mirror([_sec(2)], [_sec(2, tmin="8.0")])
    assert f.level == "error"
    assert "строка 2" in f.message


def test_mirror_row_missing():
    (f,) = check_mirror([_sec(2), _sec(3, 2)], [_sec(2)])
    assert f.level == "error"
    assert "нет в «Импорт (ПР)»" in f.message


# ─── R4 ───────────────────────────────────────────────────────────────

def test_ref_errors(tmp_path):
    p = str(tmp_path / "т.xlsx")
    wb = openpyxl.Workbook()
    wb.active.title = "Табл"
    wb.active["A1"] = "#REF!"
    wb.active["B2"] = "нормально"
    ws2 = wb.create_sheet("чисто")
    ws2["A1"] = 5
    wb.save(p)
    (f,) = check_ref_errors(p)
    assert f.level == "warn"
    assert "«Табл»: 1" in f.message


def test_ref_errors_clean(tmp_path):
    p = str(tmp_path / "т.xlsx")
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(p)
    (f,) = check_ref_errors(p)
    assert f.level == "ok"


# ─── R5 ───────────────────────────────────────────────────────────────

def _report_with_t3(path, rows):
    d = docx.Document()
    t = d.add_table(rows=1 + len(rows), cols=3)
    t.rows[0].cells[0].text = ("Наименование элемента трубопровода")
    for i, name in enumerate(rows, start=1):
        t.rows[i].cells[0].text = name
    d.save(path)


def test_valves_found(tmp_path):
    p = str(tmp_path / "о.docx")
    _report_with_t3(p, ["Труба", "Задвижка", "Отвод"])
    (f,) = check_valves_in_table3(p, ["Задвижка"])
    assert f.level == "warn"
    assert "Задвижка" in f.message


def test_valves_clean(tmp_path):
    p = str(tmp_path / "о.docx")
    _report_with_t3(p, ["Труба", "Отвод", "Бобышка"])
    (f,) = check_valves_in_table3(p, ["Задвижка"])
    assert f.level == "ok"


def test_no_table3(tmp_path):
    p = str(tmp_path / "о.docx")
    docx.Document().save(p)
    (f,) = check_valves_in_table3(p, ["Задвижка"])
    assert f.level == "warn"
    assert "не найдена" in f.message
