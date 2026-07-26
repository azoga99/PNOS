# -*- coding: utf-8 -*-
"""Тесты превью PDF (Э4 слой 1). Реальные файлы — со skip, если их нет."""

import os

import pytest

from repair.pdf_preview import page_count, render_page, close_all

_CANDIDATES = [
    r"C:\Users\volod\Desktop\Сданы на ПНОС — копия\п.2775 (Всем отправлено)"
    r"\Акт отбраковки\222-1-143 № 122 39-40 30.03.2026.pdf",
    r"C:\Users\volod\Desktop\Готовые — копия\п.2775(ШАР)\Акт отбраковки"
    r"\222-1-143 № 122 39-40 30.03.2026.pdf",
    r"C:\Users\volod\Desktop\Готовые\п.2775(ШАР)\Акт отбраковки"
    r"\222-1-143 № 122 39-40 30.03.2026.pdf",
    r"C:\Users\volod\Desktop\п.2775\Акт отбраковки"
    r"\222-1-143 № 122 39-40 30.03.2026.pdf",
]
REAL_ACT = next((p for p in _CANDIDATES if os.path.isfile(p)), _CANDIDATES[0])


@pytest.fixture(autouse=True)
def _cleanup():
    yield
    close_all()


def _need_real():
    if not os.path.isfile(REAL_ACT):
        pytest.skip("реального PDF нет на этой машине")


def test_page_count_real():
    _need_real()
    assert page_count(REAL_ACT) == 2      # акт: лист 1 текст + лист 2 схема


def test_render_page_real():
    _need_real()
    arr = render_page(REAL_ACT, 0, scale=1.0)
    h, w = arr.shape[:2]
    assert arr.shape[2] == 3
    assert 500 < w < 700 and 700 < h < 950   # А4 при 72 dpi
    # скан документа: фон в основном светлый
    assert (arr > 200).mean() > 0.5


def test_render_second_page_uses_cache():
    _need_real()
    a1 = render_page(REAL_ACT, 0, scale=0.5)
    a2 = render_page(REAL_ACT, 1, scale=0.5)
    assert a1.shape == a2.shape
