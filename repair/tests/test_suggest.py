# -*- coding: utf-8 -*-
"""
Тесты подбора сечений по таблице замены.

Данные — реальные сечения пунктов 2775 и 2701 из фикстур. Главное, что тут
проверяется: подбор НЕ выдаёт один ответ там, где сечений с одинаковым
элементом и диаметром несколько.
"""

from repair.suggest import parse_replacement, suggest_sections, describe
from repair.tests.conftest import load_fixture, sections_from


def _sections(point: str):
    return sections_from(load_fixture(point)["sections_old"])


# ─── Разбор строки таблицы замены ──────────────────────────────────────

def test_parse_perehod():
    s = parse_replacement("Переход 108х6-89х6")
    assert s.элемент == "Переход"
    assert s.размеры == [("108", "6"), ("89", "6")]
    assert s.диаметр == "108 / 89"
    assert s.номинал == "6"          # обе стенки одинаковые


def test_parse_truba_with_length():
    s = parse_replacement("Труба 219х10 L=2940мм")
    assert s.элемент == "Труба"
    assert s.размеры == [("219", "10")]
    assert s.номинал == "10"


def test_parse_otvod_with_angle():
    """«П90» — угол отвода, а не типоразмер."""
    s = parse_replacement("Отвод П90 219х12 4шт.")
    assert s.элемент == "Отвод"
    assert s.размеры == [("219", "12")]


def test_parse_ocr_noise_element():
    """OCR глотает буквы: «Пе еход» должен опознаваться."""
    assert parse_replacement("Пе еход 108х6-89х6").элемент == "Переход"


def test_parse_garbage_gives_nothing():
    s = parse_replacement("Заглушка [08х6")
    assert s.размеры == [] or s.размеры[0][0] != "108"


# ─── Подбор сечений ────────────────────────────────────────────────────

def test_many_candidates_for_common_pipe():
    """
    В п.2775 «Труба 108» встречается 26 раз — подбор обязан вернуть их все,
    а не выбрать одно наугад.
    """
    s = suggest_sections("Труба 108х6", _sections("2775"))[0]
    assert len(s.кандидаты) > 5
    assert not s.однозначно
    assert all(c.element == "Труба" for c in s.кандидаты)


def test_perehod_matches_both_diameters():
    """Переход подбирается по ОБОИМ диаметрам: «159 / 108», а не только 159."""
    s = suggest_sections("Переход 159х6-108х6", _sections("2775"))[0]
    assert s.кандидаты, "переход 159/108 в пункте есть"
    for c in s.кандидаты:
        assert set(c.diameter.replace(" ", "").split("/")) == {"159", "108"}


def test_wrong_diameter_gives_no_candidates():
    s = suggest_sections("Труба 999х8", _sections("2775"))[0]
    assert s.кандидаты == []


def test_valves_are_not_offered():
    """Задвижки по правилам не обновляются — в кандидаты не попадают."""
    for s in suggest_sections(["Задвижка 159", "Труба 159х7"],
                              _sections("2775")):
        assert all("адвижк" not in c.element for c in s.кандидаты)


def test_describe_mentions_map_when_ambiguous():
    s = suggest_sections("Труба 108х6", _sections("2775"))[0]
    текст = describe(s)
    assert "возможные сечения" in текст and "карте" in текст


def test_describe_single_candidate():
    s = suggest_sections("Труба 999х8", _sections("2775"))[0]
    assert "не найдено" in describe(s)


def test_suggestion_keeps_source_line():
    s = suggest_sections("Отвод П90 219х12 4шт.", _sections("2701"))[0]
    assert s.строка == "Отвод П90 219х12 4шт."
