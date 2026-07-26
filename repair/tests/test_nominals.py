# -*- coding: utf-8 -*-
"""
Регрессия ядра пересчёта на эталонных пунктах.

Берём таблицу толщинометрии ДО ремонта + выверенный список заменённых сечений,
применяем apply_nominals и сверяем с таблицей ПОСЛЕ ремонта из готового отчёта.
Ожидание: полное совпадение (п.2701 — 105/105, п.2775 — 95/95).
"""

import pytest

from repair.nominals import apply_nominals
from conftest import load_fixture, sections_from, replaced_from, perehod_from


def _norm(v: str) -> str:
    """Сравнение чисел без оглядки на запятую/точку («10,0» == «10.0»)."""
    return (v or "").replace(",", ".").strip()


@pytest.mark.parametrize("point,expected_total", [("2701", 105), ("2775", 95)])
def test_nominals_match_reference(point, expected_total):
    fx = load_fixture(point)
    old = sections_from(fx["sections_old"])
    expected = {s.number: s for s in sections_from(fx["sections_new"])}

    result, changes = apply_nominals(old, replaced_from(fx), perehod_from(fx))
    got = {s.number: s for s in result}

    assert len(result) == expected_total
    assert set(got) == set(expected), "набор сечений разошёлся с эталоном"

    mismatches = []
    for num, exp in expected.items():
        g = got[num]
        fields = [
            ("тип", g.element, exp.element),
            ("Sисп", _norm(g.thickness_nom), _norm(exp.thickness_nom)),
            ("диаметр", _norm(g.diameter), _norm(exp.diameter)),
            ("мин", _norm(g.tmin), _norm(exp.tmin)),
        ]
        for i, (gm, em) in enumerate(zip(g.measures, exp.measures)):
            fields.append((f"т.{i+1}", _norm(gm), _norm(em)))
        for name, gv, ev in fields:
            if gv != ev:
                mismatches.append(f"сеч.{num}/{name}: получено {gv!r}, эталон {ev!r}")

    assert not mismatches, "расхождения:\n" + "\n".join(mismatches)


@pytest.mark.parametrize("point,n_replaced", [("2701", 43), ("2775", 5)])
def test_changes_count(point, n_replaced):
    """На каждое заменённое сечение — ровно одно изменение для предпросмотра."""
    fx = load_fixture(point)
    old = sections_from(fx["sections_old"])
    _, changes = apply_nominals(old, replaced_from(fx), perehod_from(fx))
    assert len(changes) == n_replaced
    assert all(c.target == "Импорт" and c.enabled for c in changes)


def test_measures_become_nominal():
    """Все точки замеров заменённого сечения = номинал."""
    from repair.model import Section
    s = Section(number=7, element="Труба", thickness_nom="10",
                diameter="219", measures=["9.1", "9.2", "9.3", "9.1"], tmin="9.1")
    (res,), _ = apply_nominals([s], {7: 8.0})
    assert res.measures == ["8.0", "8.0", "8.0", "8.0"]
    assert res.tmin == "8.0"
    assert res.thickness_nom == "8"
    assert res.diameter == "219"   # диаметр сохраняется


def test_perehod_two_sided():
    """Переход получает двусторонние толщину и диаметр."""
    from repair.model import Section
    s = Section(number=11, element="Переход", thickness_nom="11 / 12",
                diameter="273 / 219", measures=["9.9", "10.1", "10.1", "10.1"], tmin="9.9")
    (res,), _ = apply_nominals([s], {11: 12.0}, {11: ("12 / 10", "273 / 219")})
    assert res.thickness_nom == "12 / 10"
    assert res.diameter == "273 / 219"
    assert res.measures == ["12.0", "12.0", "12.0", "12.0"]


def test_untouched_section_is_same_object():
    """Незаменённое сечение возвращается без изменений."""
    from repair.model import Section
    s = Section(number=99, element="Труба", thickness_nom="6",
                diameter="108", measures=["5.9"], tmin="5.9")
    res, changes = apply_nominals([s], {7: 8.0})
    assert res[0] is s
    assert changes == []


# ─── Раздельные поля правки (Тисп / диаметр / замер) ──────────────────

def _perehod_114():
    """Реальное сечение 114 п.2763: переход с двусторонней Тисп."""
    from repair.model import Section
    return Section(number=114, element="Переход", thickness_nom="6 / 6",
                   diameter="108 / 89",
                   measures=["2.6", "2.6", "2.8", "2.8", "5.0", "5.2"],
                   tmin="2.6", row=115)


def test_edit_measure_only_keeps_two_sided_thickness():
    """
    Замер задан, Тисп не трогаем — двусторонняя «6 / 6» остаётся как была.

    Ровно этого не хватало в одном поле «новый номинал»: число 6.0 затирало
    «6 / 6» половиной, и переход выглядел как односторонний.
    """
    from repair.nominals import SectionEdit, apply_edits
    s = _perehod_114()
    (res,), (ch,) = apply_edits([s], {114: SectionEdit(measure="6.0")})
    assert res.thickness_nom == "6 / 6"
    assert res.diameter == "108 / 89"
    assert res.measures == ["6.0"] * 6
    assert res.tmin == "6.0"
    assert "Тисп 6 / 6 · 108 / 89 · замеры 6.0/6.0" in ch.after


def test_edit_fields_are_independent():
    """Каждое поле правится отдельно; пустое — не меняется."""
    from repair.nominals import SectionEdit, apply_edits
    s = _perehod_114()
    edit = SectionEdit(thickness="6 / 5", diameter="108 / 76", measure="5.0")
    (res,), _ = apply_edits([s], {114: edit})
    assert (res.thickness_nom, res.diameter) == ("6 / 5", "108 / 76")
    assert res.measures == ["5.0"] * 6


def test_edit_measure_derived_from_thickness():
    """Замер не введён — выводится из новой Тисп (труба новая, мерить нечего)."""
    from repair.nominals import SectionEdit, apply_edits, measure_from_thickness
    s = _perehod_114()
    (res,), _ = apply_edits([s], {114: SectionEdit(thickness="8")})
    assert res.measures == ["8.0"] * 6 and res.tmin == "8.0"
    assert measure_from_thickness("12 / 10") == "12.0"
    assert measure_from_thickness("") == ""


def test_edit_diameter_only_keeps_measures():
    """Правка одного диаметра замеры не трогает."""
    from repair.nominals import SectionEdit, apply_edits
    s = _perehod_114()
    (res,), _ = apply_edits([s], {114: SectionEdit(diameter="108 / 76")})
    assert res.measures == s.measures and res.tmin == s.tmin
    assert res.thickness_nom == "6 / 6"


def test_edit_by_row_addresses_sheet_rows():
    """Живой путь адресуется строкой листа: номера сечений не уникальны."""
    from repair.nominals import SectionEdit, apply_edits
    s = _perehod_114()
    res, changes = apply_edits([s], {115: SectionEdit(measure="6.0")}, by_row=True)
    assert res[0].tmin == "6.0" and len(changes) == 1


def test_empty_edit_changes_nothing():
    from repair.nominals import SectionEdit, apply_edits
    s = _perehod_114()
    res, changes = apply_edits([s], {114: SectionEdit()})
    assert res[0] is s and changes == []
