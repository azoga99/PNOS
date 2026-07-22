# -*- coding: utf-8 -*-
"""
Тесты строки «Сведения о ремонте».

Два уровня:
  1) сверка с эталонными отчётами по СОДЕРЖИМОМУ (даты и номера актов в нужном
     порядке) — пунктуация в реальных отчётах набивалась вручную и
     непоследовательна, поэтому сверяем поля, а не байты;
  2) точная сверка вывода шаблона на синтетических записях (С0…С3), чтобы
     сам шаблон был зафиксирован.
"""

import re

from repair.model import ActRef, RepairRecord, ID_PRESENT, ID_EXPECTED
from repair.svedenia import build_svedenia
from conftest import load_fixture

DATE = re.compile(r"\d{2}\.\d{2}\.\d{4}")
ACTNUM = re.compile(r"№\s*(\d+-\d+-\d+)")


def _fields(s: str):
    """Упорядоченные даты и номера актов — «поля» строки для сверки."""
    return DATE.findall(s), ACTNUM.findall(s)


# ─── Уровень 1: сверка с эталонами ────────────────────────────────────

def test_2701_fields_match_reference():
    """С3: два комплекта ИД, ловушка даты акта (24.02 из грифа, не 24.03)."""
    records = [
        RepairRecord(id_status=ID_PRESENT, date_done="13.04.2026",
                     acts=[ActRef("222-1-84", "24.02.2026")]),
        RepairRecord(id_status=ID_PRESENT, date_done="09.04.2026",
                     acts=[ActRef("222-1-309", "07.07.2023")]),
    ]
    got = build_svedenia(records)
    ref = load_fixture("2701")["svedenia_new"]

    assert _fields(got) == _fields(ref)
    # все шесть полей на месте
    assert _fields(got)[0] == ["13.04.2026", "24.02.2026", "09.04.2026", "07.07.2023"]
    assert _fields(got)[1] == ["222-1-84", "222-1-309"]
    # признак С3 — нумерация записей
    assert "1)" in got and "2)" in got


def test_2775_fields_match_reference():
    """С1: есть акты, ИД ещё нет — без даты выполнения и без нумерации."""
    rec = RepairRecord(
        id_status=ID_EXPECTED,
        acts=[ActRef("222-1-121", "24.03.2026"), ActRef("222-1-143", "30.03.2026")],
    )
    got = build_svedenia([rec])
    ref = load_fixture("2775")["svedenia_new"]

    assert _fields(got) == _fields(ref)
    assert _fields(got)[1] == ["222-1-121", "222-1-143"]
    assert "проведена замена" not in got   # ИД нет → нет фразы о выполнении
    assert "1)" not in got                 # одна запись → без нумерации


# ─── Уровень 2: фиксация шаблона (С0…С3) ──────────────────────────────

BASE = ("Исполнительная документация в наличии и приложена "
        "к техническому паспорту на трубопровод")


def test_c0_no_repair():
    assert build_svedenia([]) == BASE
    # запись без валидных актов не создаёт ремонта
    assert build_svedenia([RepairRecord(acts=[ActRef("")])]) == BASE


def test_c1_acts_only():
    rec = RepairRecord(id_status=ID_EXPECTED, acts=[ActRef("222-1-5", "01.02.2026")])
    got = build_svedenia([rec])
    assert got == BASE + ": 01.02.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-5"


def test_c2_one_id():
    rec = RepairRecord(id_status=ID_PRESENT, date_done="10.04.2026",
                       acts=[ActRef("222-1-5", "01.02.2026")])
    got = build_svedenia([rec])
    assert got == (BASE + ": 10.04.2026 г. проведена замена участков трубопровода: "
                   "01.02.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-5")


def test_c2_two_acts_one_id():
    """Один комплект ИД на два акта — акты через разделитель, без нумерации."""
    rec = RepairRecord(id_status=ID_PRESENT, date_done="14.04.2026",
                       acts=[ActRef("222-1-121", "24.03.2026"),
                             ActRef("222-1-143", "30.03.2026")])
    got = build_svedenia([rec])
    assert "1)" not in got
    assert got.count("Акт ревизии") == 2
    assert got.startswith(BASE + ": 14.04.2026 г. проведена замена")


def test_c3_numbered():
    recs = [
        RepairRecord(id_status=ID_PRESENT, date_done="13.04.2026",
                     acts=[ActRef("222-1-84", "24.02.2026")]),
        RepairRecord(id_status=ID_PRESENT, date_done="09.04.2026",
                     acts=[ActRef("222-1-309", "07.07.2023")]),
    ]
    got = build_svedenia(recs)
    assert "1) 13.04.2026 г. проведена замена" in got
    assert "2) 09.04.2026 г. проведена замена" in got
