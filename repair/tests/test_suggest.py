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


def test_wrong_diameter_falls_back_to_wear():
    """
    Диаметра 999 в пункте нет — но пустой ответ бесполезен: при замене диаметр
    иногда меняется. Показываем утонённые трубы и помечаем догадку как догадку.
    """
    from repair.suggest import ИСТОЧНИК_ИЗНОС
    s = suggest_sections("Труба 999х8", _sections("2775"))[0]
    assert s.источник == ИСТОЧНИК_ИЗНОС
    assert not s.однозначно


def test_valves_are_not_offered():
    """Задвижки по правилам не обновляются — в кандидаты не попадают."""
    for s in suggest_sections(["Задвижка 159", "Труба 159х7"],
                              _sections("2775")):
        assert all("адвижк" not in c.element for c in s.кандидаты)


def test_describe_mentions_map_when_ambiguous():
    s = suggest_sections("Труба 108х6", _sections("2775"))[0]
    текст = describe(s)
    assert "возможные сечения" in текст and "карте" in текст


def test_describe_explains_the_fallback():
    """Из подписи должно быть ясно: это не совпадение, а список утонённых."""
    s = suggest_sections("Труба 999х8", _sections("2775"))[0]
    текст = describe(s)
    assert "диаметр мог измениться" in текст and "утонённые" in текст


def test_suggestion_keeps_source_line():
    s = suggest_sections("Отвод П90 219х12 4шт.", _sections("2701"))[0]
    assert s.строка == "Отвод П90 219х12 4шт."


# ─── Запасной ход по износу ───────────────────────────────────────────

def _секции():
    """Отводы п.2687: Ø57 нетронуты, Ø45 сильно утонён."""
    from repair.model import Section
    return [
        Section(3, "Отвод", "3.5", "57", ["2.4"], "2.4", row=4),
        Section(29, "Отвод", "3", "45", ["2.1"], "2.1", row=30),
        Section(25, "Отвод", "3", "45", ["3.5"], "3.5", row=26),
        Section(30, "Труба", "3", "45", ["3.6"], "3.6", row=31),
    ]


def test_wear_ratio():
    from repair.suggest import wear
    секции = {s.number: s for s in _секции()}
    assert round(wear(секции[29]), 2) == 0.70      # 2.1 при номинале 3
    assert wear(секции[25]) > 1.0                  # замер выше номинала


def test_worn_sections_sorted_by_wear():
    """Самое утонённое — первым; неутонённые не показываются вовсе."""
    from repair.suggest import worn_sections
    из = worn_sections(_секции(), "Отвод")
    # 3: 2.4 при 3.5 → 0.69; 29: 2.1 при 3 → 0.70. Порядок по доле остатка.
    assert [s.number for s in из] == [3, 29]


def test_worn_sections_respect_element_type():
    from repair.suggest import worn_sections
    assert [s.number for s in worn_sections(_секции(), "Труба")] == []


def test_fallback_when_diameter_matches_nothing():
    """
    Диаметр из акта не встречается в книге — при замене он мог измениться.
    Вместо пустого ответа показываем утонённые сечения того же типа.
    """
    from repair.suggest import suggest_sections, ИСТОЧНИК_ИЗНОС
    (s,) = suggest_sections(["Отвод 90°-89х5"], _секции())
    assert s.источник == ИСТОЧНИК_ИЗНОС
    assert [c.number for c in s.кандидаты] == [3, 29]


def test_fallback_never_auto_applies():
    """Догадка по износу не даёт права подставить номинал автоматически."""
    from repair.suggest import suggest_sections
    from repair.model import Section
    одно = [Section(29, "Отвод", "3", "45", ["2.1"], "2.1", row=30)]
    (s,) = suggest_sections(["Отвод 90°-89х5"], одно)
    assert len(s.кандидаты) == 1
    assert not s.однозначно


def test_exact_match_still_auto_applies():
    """Точное совпадение по типоразмеру право на подстановку сохраняет."""
    from repair.suggest import suggest_sections, ИСТОЧНИК_РАЗМЕР
    from repair.model import Section
    одно = [Section(7, "Труба", "10", "219", ["9.1"], "9.1", row=9)]
    (s,) = suggest_sections(["Труба 219х10"], одно)
    assert s.источник == ИСТОЧНИК_РАЗМЕР
    assert s.однозначно


def test_fallback_can_be_switched_off():
    from repair.suggest import suggest_sections
    (s,) = suggest_sections(["Отвод 90°-89х5"], _секции(), по_износу=False)
    assert s.кандидаты == []


# ─── Название элемента не прочиталось ──────────────────────────────────

def test_unreadable_element_matched_by_diameter():
    """
    OCR не прочитал название — подбираем по одному диаметру, среди всех типов.

    На акте п.2688 строка приходит как «Т ба 45х4 Бесшовная ст. 20», а на
    актах 222-1-121 и 222-1-143 названия нет вовсе: «1 108х6 2,5 п. м.
    Бесшовная ст. 20». Раньше такие строки не давали ни одного кандидата.
    """
    from repair.suggest import ИСТОЧНИК_БЕЗ_ЭЛЕМЕНТА

    секции = _sections("2775")
    s = suggest_sections(["1 108х6 2,5 п. м. Бесшовная ст. 20"], секции)[0]
    assert s.элемент == ""
    assert s.размеры == [("108", "6")]
    assert s.кандидаты, "по диаметру 108 в п.2775 сечения есть"
    assert s.источник == ИСТОЧНИК_БЕЗ_ЭЛЕМЕНТА
    # Тип элемента неизвестен — подставлять автоматически нельзя даже одному.
    assert not s.однозначно


def test_unreadable_element_falls_back_to_wear():
    """Диаметра из акта в книге нет — показываем самые утонённые любого типа."""
    from repair.suggest import ИСТОЧНИК_ИЗНОС

    секции = _sections("2775")
    s = suggest_sections(["Т ба 999х4 Бесшовная ст. 20"], секции)[0]
    assert s.элемент == ""
    assert s.источник == ИСТОЧНИК_ИЗНОС
    assert not s.однозначно
    assert "любого типа" in describe(s)


def test_unreadable_element_is_named_in_description():
    """Сотруднику честно сообщаем, что название не прочиталось."""
    секции = _sections("2775")
    s = suggest_sections(["1 108х6 2,5 п. м. Бесшовная ст. 20"], секции)[0]
    текст = describe(s)
    assert "элемент не прочитан" in текст
    assert "сверьтесь с таблицей замены" in текст


def test_known_element_still_restricts_type():
    """Прочитанное название по-прежнему сужает поиск — поведение не изменилось."""
    from repair.suggest import ИСТОЧНИК_РАЗМЕР

    секции = _sections("2775")
    s = suggest_sections(["1 Труба 108х6 2 п. м. Бесшовная ст. 20"], секции)[0]
    assert s.элемент == "Труба"
    assert s.источник == ИСТОЧНИК_РАЗМЕР
    assert all("труб" in c.element.lower() for c in s.кандидаты)
