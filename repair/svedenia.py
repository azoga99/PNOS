# -*- coding: utf-8 -*-
"""
Генератор строки «Сведения о ремонте» (ячейка таблицы 1 отчёта).

Единица списка — комплект ИД, не акт. Состояние выводится из данных реестра:

    С0  ремонта нет                  → только базовая фраза
    С1  есть акты, ИД нет            → базовая фраза + акты без даты выполнения
    С2  один комплект ИД             → + «дата сдачи г. проведена замена ...»
    С3  несколько комплектов ИД      → то же с нумерацией «1) 2)»

Пунктуация в реальных отчётах набивалась вручную и слегка непоследовательна;
здесь она приведена к единому виду и вынесена в шаблоны (config.SV). Строка
всегда проходит через предпросмотр, где сотрудник её видит и правит.

Проверено на эталонах: даты и номера актов совпадают с готовыми отчётами
п.2701 (С3), п.2775 (С1) и подтверждённой реконструкцией п.2775 после ИД (С2).
"""

from __future__ import annotations

from repair.config import REPAIR_CONFIG
from repair.model import RepairRecord


def _act_text(act, sv) -> str:
    return sv["ACT"].format(date=act.date, num=act.number)


def _acts_joined(record: RepairRecord, sv) -> str:
    parts = [_act_text(a, sv) for a in record.acts if a.is_valid()]
    return sv["ACT_SEP"].join(parts)


def _record_body(record: RepairRecord, sv) -> str:
    """
    Тело одной записи без нумерации:
      с ИД  → «{дата сдачи} г. проведена замена участков трубопровода: {акты}»
      без ИД → «{акты}»
    """
    acts = _acts_joined(record, sv)
    if record.has_id:
        return sv["REPLACE_PREFIX"].format(date=record.date_done) + acts
    return acts


def build_svedenia(records: list[RepairRecord], config: dict | None = None) -> str:
    """
    Собрать строку «Сведения о ремонте» из записей реестра.

    records — список RepairRecord (пустой = ремонта не было).
    """
    cfg = config or REPAIR_CONFIG
    sv = cfg["SV"]
    base = sv["BASE"]

    # С0 — ремонта нет.
    valid = [r for r in records if any(a.is_valid() for a in r.acts)]
    if not valid:
        return base

    with_id = [r for r in valid if r.has_id]

    if len(with_id) >= 2:
        # С3 — несколько комплектов ИД: нумеруем записи.
        items = [
            sv["NUMBERED_ITEM"].format(n=i, body=_record_body(r, sv))
            for i, r in enumerate(valid, start=1)
        ]
        body = sv["RECORD_SEP"].join(items)
    else:
        # С1 (акты без ИД) или С2 (один комплект ИД) — без нумерации.
        body = sv["RECORD_SEP"].join(_record_body(r, sv) for r in valid)

    return base + sv["BASE_SEP"] + body
