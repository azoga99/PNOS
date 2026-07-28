# -*- coding: utf-8 -*-
"""
Выбор листов коррозионной карты (repair.map_sheets).

Наборы имён взяты с реальных сданных пунктов: п.2704 и п.2783 — два листа,
п.2808 — три, п.2758 — один лист с именем, не совпадающим с номером пункта.
"""

import os

from repair.map_sheets import (
    is_working_map, sheet_number, working_map_files, likely_sheet,
    original_map_file,
)


# ─── Рабочая карта против исходника ОТН ───────────────────────────────

def test_original_has_extra_dot():
    """Исходник ОТН опознаётся по точке внутри имени и к правке не идёт."""
    assert not is_working_map("2763..png")
    assert not is_working_map("2704._1.png")
    assert not is_working_map("2808._3.png")


def test_working_map_has_no_dot_in_stem():
    assert is_working_map("2763.png")
    assert is_working_map("2704_1.png")
    assert is_working_map("2808_3.png")


def test_non_png_is_not_a_map():
    assert not is_working_map("2763.pdf")
    assert not is_working_map("Акт замеров толщины стенки.docx")


def test_full_path_accepted():
    assert is_working_map(r"C:\п.2704\Первичка\2704_1.png")
    assert not is_working_map(r"C:\п.2704\Первичка\2704._1.png")


# ─── Номер листа ──────────────────────────────────────────────────────

def test_sheet_number_from_suffix():
    assert sheet_number("2704_1.png") == 1
    assert sheet_number("2704_2.png") == 2
    assert sheet_number("2808_3.png") == 3


def test_single_sheet_is_first():
    assert sheet_number("2763.png") == 1
    assert sheet_number("3024.png") == 1


# ─── Отбор и порядок листов ───────────────────────────────────────────

def test_two_sheet_point():
    """п.2704: исходники отброшены, рабочие листы по порядку."""
    имена = ["2704._1.png", "2704._2.png", "2704_1.png", "2704_2.png"]
    assert working_map_files(имена) == ["2704_1.png", "2704_2.png"]


def test_three_sheet_point():
    имена = ["2808._1.png", "2808._2.png", "2808._3.png",
             "2808_1.png", "2808_2.png", "2808_3.png"]
    assert working_map_files(имена) == ["2808_1.png", "2808_2.png", "2808_3.png"]


def test_single_sheet_point():
    assert working_map_files(["2763..png", "2763.png"]) == ["2763.png"]


def test_map_name_may_differ_from_point_number():
    """п.2758: карта называется «3024» — опознаём по форме имени, не по номеру."""
    assert working_map_files(["3024..png", "3024.png"]) == ["3024.png"]


def test_sorted_by_number_not_by_string():
    """«_10» идёт после «_9», а не между «_1» и «_2»."""
    имена = ["m_10.png", "m_9.png", "m_1.png"]
    assert working_map_files(имена) == ["m_1.png", "m_9.png", "m_10.png"]


def test_empty_input():
    assert working_map_files([]) == []
    assert working_map_files(None) == []


# ─── Подсказка листа по соседям ───────────────────────────────────────

def test_neighbours_point_to_the_right_sheet():
    """
    Реальный спорный случай п.2783: сечение 1 нашлось на обоих листах.
    Рядом с настоящей подписью на листе 1 стоит сечение 2 — туда и подсказка.
    """
    by_sheet = {0: {1, 2, 5, 8, 9}, 1: {1, 7, 27, 110}}
    assert likely_sheet(1, by_sheet) == 0


def test_no_hint_on_tie():
    """Соседей поровну — подсказки нет, решает сотрудник."""
    by_sheet = {0: {50, 52}, 1: {49, 53}}
    assert likely_sheet(51, by_sheet) is None


def test_no_hint_without_neighbours():
    by_sheet = {0: {200, 201}, 1: {300, 301}}
    assert likely_sheet(5, by_sheet) is None


def test_hint_ignores_the_number_itself():
    """Сам номер соседом не считается — иначе подсказка была бы всегда."""
    by_sheet = {0: {7}, 1: {7}}
    assert likely_sheet(7, by_sheet) is None


# ─── Имя исходника ОТН ────────────────────────────────────────────────

def test_original_name_single_sheet():
    """Точка ставится перед расширением: «2763.png» → «2763..png»."""
    assert original_map_file("2763.png") == "2763..png"


def test_original_name_multi_sheet():
    """У листа точка встаёт перед суффиксом, а не в конец имени."""
    assert original_map_file("2704_1.png") == "2704._1.png"
    assert original_map_file("2808_3.png") == "2808._3.png"


def test_original_name_keeps_folder():
    из = os.path.join("C:\\", "п.2704", "Первичка", "2704_2.png")
    assert original_map_file(из).endswith(os.path.join("Первичка", "2704._2.png"))


def test_original_is_not_a_working_map():
    """Круг замкнулся: имя исходника рабочей картой не считается."""
    for имя in ("2763.png", "2704_1.png", "2808_3.png"):
        assert not is_working_map(original_map_file(имя))
