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
    assert got == BASE + ":\n01.02.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-5"


def test_act_without_date_omits_date_part():
    """Дата акта ещё не внесена (черновик) — « г.» с пустой датой не пишется."""
    rec = RepairRecord(id_status=ID_EXPECTED, acts=[ActRef("222-1-5")])
    got = build_svedenia([rec])
    assert got == BASE + ":\nАкт ревизии и отбраковки трубопровода № 222-1-5"
    assert " г. Акт" not in got


def test_c2_one_id():
    rec = RepairRecord(id_status=ID_PRESENT, date_done="10.04.2026",
                       acts=[ActRef("222-1-5", "01.02.2026")])
    got = build_svedenia([rec])
    assert got == (BASE + ":\n10.04.2026 г. проведена замена участков трубопровода:\n"
                   "01.02.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-5")


def test_single_replaced_element_uses_singular():
    """Заменён один элемент — «замена участка», как в готовом отчёте п.2763."""
    rec = RepairRecord(id_status=ID_PRESENT, date_done="16.04.2026",
                       acts=[ActRef("222-1-179", "01.04.2026")],
                       replacements="Переход 108х6-89х6")
    assert "замена участка трубопровода" in build_svedenia([rec])


def test_several_replaced_elements_use_plural():
    rec = RepairRecord(id_status=ID_PRESENT, date_done="09.04.2026",
                       acts=[ActRef("222-1-309", "07.07.2023")],
                       replacements="Труба 89х5; Труба 219х10; Отвод 219х12")
    assert "замена участков трубопровода" in build_svedenia([rec])


def test_c2_two_acts_one_id():
    """Один комплект ИД на два акта — акты через разделитель, без нумерации."""
    rec = RepairRecord(id_status=ID_PRESENT, date_done="14.04.2026",
                       acts=[ActRef("222-1-121", "24.03.2026"),
                             ActRef("222-1-143", "30.03.2026")])
    got = build_svedenia([rec])
    assert "1)" not in got
    assert got.count("Акт ревизии") == 2
    assert got.startswith(BASE + ":\n14.04.2026 г. проведена замена")


# ─── Точная сверка с готовыми отчётами ─────────────────────────────────
# Строки сняты из отчётов ЗЭПБ готовых пунктов (папка «Готовые — копия»).
# Разделитель в ячейке Word — перенос строки, поэтому и в шаблоне «\n».

def test_matches_finished_report_2763():
    """С2: один комплект ИД, один заменённый элемент → «участка»."""
    rec = RepairRecord(id_status=ID_PRESENT, date_done="16.04.2026",
                       acts=[ActRef("222-1-179", "01.04.2026")],
                       replacements="Переход 108х6-89х6")
    assert build_svedenia([rec]) == (
        "Исполнительная документация в наличии и приложена к техническому "
        "паспорту на трубопровод:\n"
        "16.04.2026 г. проведена замена участка трубопровода:\n"
        "01.04.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-179")


def test_matches_finished_report_2701():
    """С3: два комплекта ИД, нумерация записей и перенос строки."""
    recs = [
        RepairRecord(id_status=ID_PRESENT, date_done="13.04.2026",
                     acts=[ActRef("222-1-84", "24.02.2026")],
                     replacements="Труба 219х10; Отвод 219х12"),
        RepairRecord(id_status=ID_PRESENT, date_done="09.04.2026",
                     acts=[ActRef("222-1-309", "07.07.2023")],
                     replacements="Труба 89х5; Труба 219х10; Отвод 219х12"),
    ]
    assert build_svedenia(recs) == (
        "Исполнительная документация в наличии и приложена к техническому "
        "паспорту на трубопровод:\n"
        "1) 13.04.2026 г. проведена замена участков трубопровода:\n"
        "24.02.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-84;\n"
        "2) 09.04.2026 г. проведена замена участков трубопровода:\n"
        "07.07.2023 г. Акт ревизии и отбраковки трубопровода № 222-1-309")


def test_matches_finished_report_2775():
    """
    С1: акты есть, ИД ещё нет.

    В готовом отчёте строка заканчивается точкой, в двух других — нет.
    Точка не ставится: пунктуация в отчётах набивалась вручную и
    непоследовательна, а поправить строку сотрудник может прямо в
    предпросмотре.
    """
    rec = RepairRecord(id_status=ID_EXPECTED,
                       acts=[ActRef("222-1-121", "24.03.2026"),
                             ActRef("222-1-143", "30.03.2026")])
    assert build_svedenia([rec]) == (
        "Исполнительная документация в наличии и приложена к техническому "
        "паспорту на трубопровод:\n"
        "24.03.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-121;\n"
        "30.03.2026 г. Акт ревизии и отбраковки трубопровода № 222-1-143")


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
