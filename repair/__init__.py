# -*- coding: utf-8 -*-
"""
Модуль «Ремонт» (тестовая ветка repair-module).

Автоматизация учёта ремонтов трубопровода в заключениях ЭПБ: перенос сведений
из актов ревизии и отбраковки и исполнительной документации во все три
артефакта пункта — Excel-расчёт, отчёт Word и коррозионную карту.

Пакет строится снизу вверх: сначала изолированное ядро логики (pure-Python,
без GUI и без COM), проверяемое ретро-тестами на эталонных пунктах п.2701 и
п.2775, затем — интеграция в конвейер ПНОС.

Полная спецификация: см. память проекта pnos-repair-module-rules.
Ядро подтверждено: п.2701 — 105/105 сечений, п.2775 — 95/95, строка — 6/6 полей.
"""

from repair.model import ActRef, RepairRecord, Section, Change
from repair.svedenia import build_svedenia
from repair.nominals import apply_nominals, apply_edits, SectionEdit

__all__ = [
    "ActRef", "RepairRecord", "Section", "Change",
    "build_svedenia", "apply_nominals", "apply_edits", "SectionEdit",
]
