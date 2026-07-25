# -*- coding: utf-8 -*-
"""
Тесты распознавания: автоповорот страниц и OCR зонами.

Опираются на обучающую выборку («для обучени»), где отобраны реальные
документы с повёрнутыми страницами. Если выборки или Windows OCR на машине
нет — тесты пропускаются.
"""

import os

import pytest

from repair.ocr import (available, ocr_page, ocr_zone, ocr_zone_multi,
                        best_match, plain_text, score_lines)

DATASET = r"C:\Users\volod\Desktop\для обучени"

# Страницы, на которых поворот критичен (замеры из разметки выборки).
ROTATED_90 = os.path.join(
    DATASET, "36-30", "п. 2701", "Исполнительная", "36-30, №6, 222-1-309.pdf")
ROTATED_180 = os.path.join(
    DATASET, "39-40", "п. 2775", "Исполнительная",
    "ИД по Акту ОТН №222-1-143,121. Линия №122.pdf_v00.100001268600097.pdf")
ACT = os.path.join(
    DATASET, "39-40", "п. 2763", "Акт отбраковки",
    "222-1-179 №108 39-40 01.04.2026.pdf")


def _need(path):
    if not available():
        pytest.skip("Windows OCR на этой машине недоступен")
    if not os.path.isfile(path):
        pytest.skip("обучающей выборки нет на этой машине")


def test_landscape_page_is_rotated_and_read():
    """стр. 7 п.2701 повёрнута на 270°: без автоповорота OCR возвращал пусто."""
    _need(ROTATED_90)
    res = ocr_page(ROTATED_90, 7)
    assert res["rotation"] == 270
    assert res["lines"], "после поворота страница обязана читаться"
    low = plain_text(res).lower()
    assert "качеств" in low and "36-30" in low


def test_upside_down_page_is_rotated():
    """стр. 9 п.2775 перевёрнута на 180°: без поворота OCR даёт мусор."""
    _need(ROTATED_180)
    res = ocr_page(ROTATED_180, 9)
    assert res["rotation"] == 180
    assert res["score"] > res["scores"][0]


def test_upright_page_is_not_rotated():
    _need(ACT)
    res = ocr_page(ACT, 1)
    assert res["rotation"] == 0
    assert "222-1-179" in plain_text(res)


def test_words_have_boxes():
    """Рамки слов нужны для разбора таблиц: реестр OCR отдаёт по ячейкам."""
    _need(ACT)
    res = ocr_page(ACT, 1)
    word = res["lines"][0]["words"][0]
    assert set(word) == {"t", "x", "y", "w", "h"}
    assert word["w"] > 0 and word["h"] > 0


def test_zone_reads_replacement_table():
    """Зона «нижние 70–100 %» последнего листа акта — это таблица замены."""
    _need(ACT)
    zone = ocr_zone(ACT, 2, 0.70, 1.0)
    joined = " ".join(ln["text"] for ln in zone["lines"]).lower()
    assert "замене" in joined            # заголовок таблицы
    assert "переход" in joined or "пе еход" in joined   # OCR глотает «р»


def test_multi_scale_consensus_reads_size_cell():
    """
    OCR нестабилен по масштабу рендера: типоразмер «108х6-89х6» читается
    верно при 1.5/2.2/3.0/3.5 и ломается при 2.0/2.5. Согласование нескольких
    проходов даёт устойчивый результат — один проход не даёт.
    """
    _need(ACT)
    multi = ocr_zone_multi(ACT, 2, 0.70, 1.0)
    size = best_match(multi["варианты"],
                      r"\d{2,3}х\d(?:[.,]\d)?(?:-\d{2,3}х\d(?:[.,]\d)?)?")
    assert size == "108х6-89х6"


def test_score_empty_is_zero():
    assert score_lines([]) == 0
    assert score_lines([{"text": "   ", "words": []}]) == 0
