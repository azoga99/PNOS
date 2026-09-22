# -*- coding: utf-8 -*-
"""
Пункты, собранные по чужой строке графика (этап 4, workers.stage3_worker).

До 2.2.4 п.3224 (АВТ-5) получил в паспорт!F9 номер 3224 вместо 3543 и
собрался по строке 517 ХВП. Новый этап 4 такие пункты находит по F9 и
собирает заново — но только их: готовые пункты с верным F9 и пункты, которые
проверить нечем, не трогаются.
"""

import json
import os

import pytest

openpyxl = pytest.importorskip("openpyxl")

from status_manager import get_status_file_path, get_status_value, reset_stage_status  # noqa: E402
from workers.stage3_worker import (  # noqa: E402
    FOREIGN_REPORTS_DIR, REBUILD_KEYS, built_by_foreign_row, set_aside_reports,
)
from workers.stage4_worker import find_report_file  # noqa: E402

PLAN = {"row": 3573, "a": 3543, "place": "АВТ–5"}


def _book(folder, f9, name="ПНОС_ТТП.xlsm"):
    wb = openpyxl.Workbook()
    wb.active.title = "паспорт"
    wb.active["F9"] = f9
    wb.save(folder / name)


def _touch(folder, *names):
    for n in names:
        (folder / n).write_bytes(b"")


# ─── Распознавание ─────────────────────────────────────────────────────

def test_point_built_by_column_a_is_found(tmp_path):
    _book(tmp_path, 3224)
    assert built_by_foreign_row(str(tmp_path), PLAN) == 3224


def test_point_built_by_its_row_is_left_alone(tmp_path):
    _book(tmp_path, 3543.0)   # COM пишет числа как float
    assert built_by_foreign_row(str(tmp_path), PLAN) is None


def test_cleansource_is_not_taken_for_the_point_workbook(tmp_path):
    """
    CleanSource.xlsx от этапа 5 идёт по алфавиту раньше ПНОС_ТТП.xlsm.
    Прочитай проверка F9 из неё — пункт посчитался бы собранным верно.
    """
    _book(tmp_path, 3224)
    _book(tmp_path, 3543, name="CleanSource.xlsx")
    assert built_by_foreign_row(str(tmp_path), PLAN) == 3224


@pytest.mark.parametrize("setup", ["empty_f9", "no_book", "broken_book"])
def test_nothing_to_check_means_no_rebuild(tmp_path, setup):
    """Готовый пункт по догадке не пересобирается."""
    if setup == "empty_f9":
        _book(tmp_path, None)
    elif setup == "broken_book":
        (tmp_path / "ПНОС_ТТП.xlsm").write_bytes(b"not a zip")
    assert built_by_foreign_row(str(tmp_path), PLAN) is None


# ─── Перенос отчётов ───────────────────────────────────────────────────

def test_reports_are_moved_aside_not_deleted(tmp_path):
    wrong = "ЗЭПБ № ТТП-3224.26 - ттп № 1086.docx"
    _touch(tmp_path, wrong, "ЗЭПБ № ТТП-3224.26 - ттп № 1086 ПЕЧАТЬ.docx",
           "Отчет_3224.docx", "ПНОС_ТТП.docx", "~$ЗЭПБ.docx", "Акт замеров.docx")

    moved = set_aside_reports(str(tmp_path))

    assert sorted(moved) == sorted([
        wrong, "ЗЭПБ № ТТП-3224.26 - ттп № 1086 ПЕЧАТЬ.docx", "Отчет_3224.docx"])
    aside = tmp_path / FOREIGN_REPORTS_DIR
    assert sorted(os.listdir(aside)) == sorted(moved)
    # шаблон, временный файл Word и документы пункта остаются на месте
    left = set(os.listdir(tmp_path)) - {FOREIGN_REPORTS_DIR}
    assert left == {"ПНОС_ТТП.docx", "~$ЗЭПБ.docx", "Акт замеров.docx"}


def test_stale_report_no_longer_wins_over_rebuilt_one(tmp_path):
    """
    Без переноса после пересборки в папке два «ЗЭПБ…», и неверный «ТТП-3224»
    по алфавиту раньше верного «ТТП-3543» — этапы 6–7 вставили бы фото и
    акты в него.
    """
    _touch(tmp_path, "ЗЭПБ № ТТП-3224.26 - ттп № 1086.docx")
    set_aside_reports(str(tmp_path))
    assert find_report_file(str(tmp_path)) is None

    _touch(tmp_path, "ЗЭПБ № ТТП-3543.26 - ттп № 3224.docx")
    assert os.path.basename(find_report_file(str(tmp_path))) == \
        "ЗЭПБ № ТТП-3543.26 - ттп № 3224.docx"


def test_second_rebuild_does_not_overwrite_earlier_aside_copy(tmp_path):
    (tmp_path / FOREIGN_REPORTS_DIR).mkdir()
    (tmp_path / FOREIGN_REPORTS_DIR / "Отчет_3224.docx").write_bytes(b"first")
    (tmp_path / "Отчет_3224.docx").write_bytes(b"second")

    set_aside_reports(str(tmp_path))

    aside = tmp_path / FOREIGN_REPORTS_DIR
    assert (aside / "Отчет_3224.docx").read_bytes() == b"first"
    assert (aside / "Отчет_3224 (2).docx").read_bytes() == b"second"


# ─── Флаги ─────────────────────────────────────────────────────────────

def test_rebuild_keys_keep_manual_form_and_download(tmp_path):
    """
    Restore пишет только паспорт!AL2:AL51, поля формы шага 3 лежат в
    столбце B — форму заново заполнять не нужно, её флаги остаются.
    """
    status = {"stage1": True, "stage2": True, "manual3": True,
              "manual3_draft": {"B22": "Ибрагимов Р.А."},
              "stage3": True, "stage4": True, "stage5": True, "stage6": True,
              "gi_number": "39-40"}
    with open(get_status_file_path(str(tmp_path)), "w", encoding="utf-8") as f:
        json.dump(status, f, ensure_ascii=False)

    for key in REBUILD_KEYS:
        reset_stage_status(str(tmp_path), key)

    for key in ("stage3", "stage4", "stage5", "stage6", "gi_number"):
        assert get_status_value(str(tmp_path), key) is None, key
    for key in ("stage1", "stage2", "manual3", "manual3_draft"):
        assert get_status_value(str(tmp_path), key) == status[key], key
