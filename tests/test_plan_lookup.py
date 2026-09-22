# -*- coding: utf-8 -*-
"""
Номер пункта → строка графика для этапа 4 (excel_service).

Папка пункта названа по столбцу B графика, макрос Restore ищет строку по
столбцу A. Одно и то же число в A и в B — разные пункты, поэтому ошибка
перевода не падает, а тихо собирает отчёт по чужому пункту: п.3224 (АВТ-5)
вышел с установкой 517 ХВП и актами ГИ от неё.
"""

import os

import pytest

from excel_service import (
    PlanLookupError, build_plan_index, plan_key, resolve_plan_number,
)


def _row(a, b, place="", kind="ТТП", diag="ЭПБ", deleted=None, mark=None):
    """Строка листа «ПЛАН», столбцы A…H."""
    return (a, b, "ПППН", place, kind, diag, deleted, mark)


# ─── Главный случай ─────────────────────────────────────────────────────

def test_point_3224_goes_to_its_own_row_not_to_column_a():
    """B=3224 — АВТ-5 (A=3543); A=3224 — совсем другой пункт, 517 ХВП."""
    index = build_plan_index([
        _row(3224, 1086, "517 ХВП Химводоподготовка", diag="ТО"),
        _row(3543, 3224, "АВТ–5"),
    ])
    match = resolve_plan_number(index, "3224")
    assert match["a"] == 3543
    assert match["place"] == "АВТ–5"
    assert match["row"] == 2


def test_missing_point_is_an_error_not_a_fallback():
    """
    Раньше ненайденный номер уходил в F9 как есть — то есть в поиск по
    столбцу A. Так п.3224 из устаревшей копии графика собрался по 517 ХВП.
    """
    index = build_plan_index([_row(3224, 1086, "517 ХВП")])
    with pytest.raises(PlanLookupError, match="нет в столбце B"):
        resolve_plan_number(index, "3224")


# ─── Нумерация в B своя у каждого вида оборудования ─────────────────────

def test_same_b_in_other_equipment_types_is_ignored():
    """
    Старый код брал последнюю строку с таким B на всём листе — у 15 пунктов
    ТТП / ЭПБ это была строка печи или РВС.
    """
    index = build_plan_index([
        _row(37, 7, "36-30"),
        _row(1589, 7, "36-30", kind="Печь"),
        _row(3542, 7, "Склад № 7", kind="ПС"),
    ])
    assert resolve_plan_number(index, 7)["a"] == 37


def test_only_other_types_means_not_found():
    index = build_plan_index([_row(1589, 7, kind="Печь"), _row(2393, 7, kind="РВС")])
    with pytest.raises(PlanLookupError):
        resolve_plan_number(index, 7)


def test_technical_survey_rows_are_not_epb_points():
    """Этап 1 берёт только ЭПБ; строка ТТП / ТО с тем же B — чужая."""
    index = build_plan_index([
        _row(100, 212, "36-20", diag="ТО"),
        _row(2202, 212, "36-20"),
    ])
    assert resolve_plan_number(index, 212)["a"] == 2202


# ─── Повторы внутри ТТП / ЭПБ ───────────────────────────────────────────

def test_duplicate_resolved_by_marked_rows():
    """Повторы у ТТП — строки, уже помеченные в H или G; этап 1 их не берёт."""
    index = build_plan_index([
        _row(2169, 210, "36-20", mark="ДО/ТТП"),
        _row(2170, 210, "36-20", deleted=1),
        _row(2171, 210, "36-20"),
    ])
    assert resolve_plan_number(index, 210)["a"] == 2171


def test_two_live_rows_is_an_error():
    index = build_plan_index([_row(2272, 359, "24-6"), _row(2273, 359, "24-6")])
    with pytest.raises(PlanLookupError, match=r"несколько раз \(строки 1, 2\)"):
        resolve_plan_number(index, 359)


def test_all_duplicates_marked_is_still_ambiguous():
    index = build_plan_index([
        _row(2272, 359, mark="ДО/ТТП"),
        _row(2273, 359, mark="ДО/ТТП"),
    ])
    with pytest.raises(PlanLookupError, match="несколько раз"):
        resolve_plan_number(index, 359)


# ─── Ячейки и ключи ─────────────────────────────────────────────────────

@pytest.mark.parametrize("cell", [3224, 3224.0, "3224", " 3224 ", "3224,0"])
def test_number_forms_give_one_key(cell):
    assert plan_key(cell) == 3224


def test_float_from_com_is_written_as_integer():
    """COM отдаёт числа как float: в F9 должно уйти 3543, а не 3543.0."""
    index = build_plan_index([_row(3543.0, 3224.0, "АВТ–5")])
    match = resolve_plan_number(index, "3224")
    assert match["a"] == 3543 and isinstance(match["a"], int)


def test_empty_column_a_is_an_error():
    index = build_plan_index([_row(None, 3224, "АВТ–5")])
    with pytest.raises(PlanLookupError, match="пуст столбец A"):
        resolve_plan_number(index, 3224)


def test_non_numeric_folder_is_not_found():
    """Раньше странное имя папки молча превращалось в п.2719."""
    index = build_plan_index([_row(2719, 2719)])
    with pytest.raises(PlanLookupError):
        resolve_plan_number(index, "3224а")


def test_header_and_short_rows_are_skipped():
    index = build_plan_index([
        ("Номер п/п", "№ п/п по графикам", "Подразделение", "Объект",
         "Вид оборудования", " Вид диагностирования", "Удаленные", "V"),
        (None,),
        (),
        _row(3543, 3224, "АВТ–5"),
    ])
    assert list(index) == [3224]
    assert index[3224][0]["row"] == 4


# ─── Настоящий график ──────────────────────────────────────────────────

_MASTER = r"C:\Users\volod\Desktop\ПНОС сводный график Перечни ЭПБ на 2026г.xlsm"


@pytest.mark.skipif(not os.path.isfile(_MASTER), reason="нет копии графика")
def test_real_schedule():
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.load_workbook(_MASTER, read_only=True, data_only=True)
    try:
        rows = wb["ПЛАН"].iter_rows(min_col=1, max_col=8, values_only=True)
        index = build_plan_index(rows)
    finally:
        wb.close()

    assert resolve_plan_number(index, 3224)["a"] == 3543   # АВТ–5, а не 517 ХВП
    assert resolve_plan_number(index, 3081)["a"] == 3010   # УНСО тит.80-130, а не 80–132


# ─── Проверка уже собранных пунктов (tools/check_plan_numbers.py) ──────

def _check_tool():
    import importlib.util
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "tools", "check_plan_numbers.py")
    spec = importlib.util.spec_from_file_location("check_plan_numbers", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _point(tmp_path, name, f9, report=None):
    openpyxl = pytest.importorskip("openpyxl")
    folder = tmp_path / name
    folder.mkdir()
    wb = openpyxl.Workbook()
    wb.active.title = "паспорт"
    wb.active["F9"] = f9
    wb.save(folder / "ПНОС_ТТП.xlsm")
    if report:
        (folder / report).write_bytes(b"")
    return str(folder)


_ROWS = [
    _row(3224, 1086, "517 ХВП Химводоподготовка", diag="ТО"),
    _row(3543, 3224, "АВТ–5"),
]


def _by_a(rows):
    return {plan_key(r[0]): (n, r[3], r[4]) for n, r in enumerate(rows, start=1)}


def test_check_finds_point_built_by_column_a(tmp_path):
    tool = _check_tool()
    folder = _point(tmp_path, "п.3224(отпр Шар)", 3224, report="Отчет_3224.docx")
    status, text = tool.check_folder(folder, build_plan_index(_ROWS), _by_a(_ROWS))
    assert status == "wrong"
    assert "517 ХВП" in text and "должно быть 3543" in text and "Отчет_3224.docx" in text


def test_check_accepts_point_built_by_its_row(tmp_path):
    tool = _check_tool()
    folder = _point(tmp_path, "п.3224", 3543.0)
    status, _ = tool.check_folder(folder, build_plan_index(_ROWS), _by_a(_ROWS))
    assert status == "ok"


def test_check_reports_point_not_built_yet(tmp_path):
    tool = _check_tool()
    folder = _point(tmp_path, "п.3224", None)
    status, _ = tool.check_folder(folder, build_plan_index(_ROWS), _by_a(_ROWS))
    assert status == "not_run"
