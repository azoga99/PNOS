# -*- coding: utf-8 -*-
"""
Чтение паспорта без Word (passport_data).

Год ввода в эксплуатацию попадает в отчёт через данные!C12 = YEAR(паспорт!B19),
поэтому здесь проверяется главное: год разбирается из любого вида даты, а при
правке года день и месяц не теряются.
"""

import datetime
import glob
import os

import pytest

import passport_data as pd


# Реальные пункты на этой машине — как в repair/tests: если их нет, тесты
# со сканами пропускаются, а чистая логика проверяется всегда.
_CANDIDATES = [
    r"C:\Users\volod\Desktop\Сданы на ПНОС — копия",
    r"C:\Users\volod\Desktop\ПНОС1",
]


def _real_points(limit=6):
    out = []
    for root in _CANDIDATES:
        for folder in sorted(glob.glob(os.path.join(root, "п.*"))):
            if os.path.isdir(folder):
                out.append(folder)
            if len(out) >= limit:
                return out
    return out


# ── год из значения любого вида ──────────────────────────────────────

@pytest.mark.parametrize("value, expected", [
    ("01.05.1975", "1975"),
    ("1.5.1975", "1975"),
    (datetime.datetime(2023, 3, 11), "2023"),
    (datetime.date(1976, 6, 30), "1976"),
    ("1975", "1975"),
    (1975, "1975"),
    ("", ""),
    (None, ""),
    ("   ", ""),
    ("не указан", ""),
])
def test_year_of(value, expected):
    assert pd.year_of(value) == expected


@pytest.mark.parametrize("value, expected", [
    ("01.05.1975", ("01", "05")),
    ("1.5.1975", ("01", "05")),
    (datetime.datetime(2023, 3, 11), ("11", "03")),
    ("1975", None),
    ("", None),
    (None, None),
])
def test_day_month_of(value, expected):
    assert pd.day_month_of(value) == expected


# ── сборка значения для B19 ──────────────────────────────────────────

def test_commission_keeps_day_and_month():
    """Правка года не должна терять день и месяц исходной даты."""
    assert pd.commission_value("1976", "01.05.1975") == "01.05.1976"


def test_commission_prefers_first_parsable_source():
    """Источники перебираются по порядку: ячейка, затем паспорт."""
    assert pd.commission_value("1980", None, "22.10.1975") == "22.10.1980"
    assert pd.commission_value("1980", "", "22.10.1975") == "22.10.1980"


def test_commission_falls_back_to_first_of_january():
    """Если даты нет нигде — 01.01, лишь бы YEAR() вернул нужный год."""
    assert pd.commission_value("1980") == "01.01.1980"
    assert pd.commission_value("1980", None, "") == "01.01.1980"


def test_commission_from_datetime_cell():
    assert pd.commission_value("2024", datetime.datetime(2023, 3, 11)) == "11.03.2024"


@pytest.mark.parametrize("year", ["", None, "19", "19755", "abcd", "  "])
def test_commission_rejects_non_year(year):
    """Мусор не пишется в книгу: пустая строка — сигнал «не трогать ячейку»."""
    assert pd.commission_value(year, "01.05.1975") == ""


# ── реальные паспорта ────────────────────────────────────────────────

def test_reads_real_passports():
    points = _real_points()
    if not points:
        pytest.skip("нет локальных папок пунктов")
    found = 0
    for folder in points:
        doc = pd.find_passport_docx(folder)
        if not doc:
            continue
        assert doc.lower().endswith((".docx", ".docm"))
        variables = pd.read_doc_vars(doc)
        assert variables, f"у {os.path.basename(folder)} не прочитались DocVariables"
        # F051 есть не у всех паспортов — ради этого поле и заводилось
        date = pd.commission_date(folder)
        if date:
            assert pd.year_of(date).isdigit()
        found += 1
    if not found:
        pytest.skip("в локальных папках нет паспортов .docx")


def test_missing_folder_is_not_an_error():
    assert pd.find_passport_docx(r"C:\нет\такой\папки") is None
    assert pd.commission_date(r"C:\нет\такой\папки") == ""
    assert pd.read_doc_vars(r"C:\нет\такого\файла.docx") == {}
