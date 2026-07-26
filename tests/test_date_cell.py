# -*- coding: utf-8 -*-
"""Значение ячейки с маской даты (widgets.date_cell_value)."""

from widgets import date_cell_value


def test_empty_mask_is_empty_string():
    """Нетронутое поле — это пустая ячейка, а не подчёркивания."""
    assert date_cell_value("__.__.____") == ""
    assert date_cell_value("  .  .    ") == ""
    assert date_cell_value("") == ""


def test_full_date_kept():
    assert date_cell_value("16.04.2026") == "16.04.2026"


def test_spaces_trimmed():
    assert date_cell_value("  01.04.2026 ") == "01.04.2026"


def test_partial_input_kept_as_is():
    """Недобитую дату не выбрасываем — сотрудник увидит её и допишет."""
    assert date_cell_value("16.04.____") == "16.04.____"
