# -*- coding: utf-8 -*-
"""
Этапы распознавания: названия в коде и в наборах этапов не должны разойтись.

`_set_phase` находит сегмент полосы по названию. Опечатка в названии не упала
бы с ошибкой — просто перестал бы двигаться прогресс, и заметить это можно было
бы только глазами на живом комплекте ИД.
"""

import os
import re

from workers.repair_worker import RepairOcrWorker

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ВСЕ_ЭТАПЫ = set(RepairOcrWorker.PHASES_ACT) | set(RepairOcrWorker.PHASES_ID)


def _used_names(path: str, pattern: str) -> set:
    with open(path, encoding="utf-8") as f:
        return set(re.findall(pattern, f.read()))


def test_phase_sets_are_distinct_and_named():
    assert RepairOcrWorker.PHASES_ACT and RepairOcrWorker.PHASES_ID
    for набор in (RepairOcrWorker.PHASES_ACT, RepairOcrWorker.PHASES_ID):
        assert len(set(набор)) == len(набор), "названия этапов повторяются"


def test_worker_phase_calls_are_known():
    """Каждое `_set_phase("…")` в воркере попадает в один из наборов."""
    названия = _used_names(os.path.join(ROOT, "workers", "repair_worker.py"),
                           r'_set_phase\("([^"]+)"\)')
    assert названия, "вызовы _set_phase не найдены — тест устарел"
    assert названия <= ВСЕ_ЭТАПЫ, f"неизвестные этапы: {названия - ВСЕ_ЭТАПЫ}"


def test_id_map_phase_calls_are_known():
    """`build_page_map` отмечает этапы теми же названиями."""
    названия = _used_names(os.path.join(ROOT, "repair", "id_map.py"),
                           r'phase\("([^"]+)"\)')
    assert названия, "вызовы phase() в id_map не найдены — тест устарел"
    assert названия <= set(RepairOcrWorker.PHASES_ID)


def test_id_phases_cover_the_long_parts():
    """Разбор реестра и блок подписей — обязательные вехи комплекта ИД."""
    assert "Реестр ИД" in RepairOcrWorker.PHASES_ID
    assert RepairOcrWorker.PHASES_ID[-1] == "Блок подписей"
