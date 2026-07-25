# -*- coding: utf-8 -*-
"""
Тесты карты комплекта ИД (правило Н1).

Разбор таблицы реестра проверяется на СОХРАНЁННОМ OCR из обучающей выборки
(`для обучени/_ocr/<id>/pNNN.json`) — такие тесты детерминированы и не требуют
Windows OCR. Построение карты целиком проверяется отдельно и пропускается, если
выборки или распознавания на машине нет.
"""

import json
import os

import pytest

from repair.id_map import (
    parse_registry, registry_score, quality, build_page_map,
    quality_cert_section, find_section, section_pages, SEC_ACT,
)

DATASET = r"C:\Users\volod\Desktop\для обучени"
OCR = os.path.join(DATASET, "_ocr")


def _ocr(doc_id: str, page: int) -> dict:
    path = os.path.join(OCR, doc_id, f"p{page:03d}.json")
    if not os.path.isfile(path):
        pytest.skip("обучающей выборки нет на этой машине")
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ─── Разбор таблицы реестра на сохранённом OCR ────────────────────────

def test_registry_parsed_from_saved_ocr():
    """п.2763, стр. 3 — копия реестра: 19 позиций, диапазоны подряд."""
    reg = parse_registry(_ocr("2763-id", 3), pages_total=44)
    rows = [r for r in reg["строки"] if r.get("стр")]
    assert len(rows) >= 17
    assert reg["качество"]["оценка"] == "высокое"

    # Разделы идут без разрывов: конец предыдущего + 1 = начало следующего.
    for a, b in zip(rows, rows[1:]):
        assert b["стр"][0] == a["стр"][1] + 1

    первый = rows[0]
    assert "акт" in первый["название"].lower()
    assert первый["стр"] == [1, 2]        # акт ОТН на двух листах


def test_registry_recovers_lost_range_bounds():
    """OCR теряет то левое, то правое число диапазона — они достраиваются
    по «Кол-ву листов» и непрерывности разделов."""
    reg = parse_registry(_ocr("2763-id", 3), pages_total=44)
    восстановленные = [r for r in reg["строки"] if r.get("стр_восстановлены")]
    assert восстановленные, "хотя бы один диапазон должен быть достроен"
    for r in reg["строки"]:
        if r.get("стр") and r.get("листов"):
            assert r["стр"][1] - r["стр"][0] + 1 == r["листов"]


def test_title_page_is_not_registry():
    """Титул «ИСПОЛНИТЕЛЬНАЯ ДОКУМЕНТАЦИЯ» похож на реестр по словам, но
    реестром считаться не должен: иначе смещение уезжает."""
    титул = parse_registry(_ocr("2763-id", 1), pages_total=44)
    реестр = parse_registry(_ocr("2763-id", 3), pages_total=44)
    assert registry_score(титул) < registry_score(реестр)
    assert титул["похожесть_заголовка"] < реестр["похожесть_заголовка"]


def test_registry_offset_candidates():
    """Смещение по формуле Н1: реестр на стр. 3, сам себя в списке не числит."""
    reg = parse_registry(_ocr("2763-id", 3), pages_total=44)
    assert reg["смещение"] == 3
    assert reg["реестр_в_списке"] is False


def test_header_reads_act_number_and_date_with_ocr_noise():
    """
    п.2701 («ТП Инжиниринг»): в шапке реестра печатно стоит
    «№222-1-ЗО9 от 07.07.2023г.» — цифры подменены кириллицей, а дата
    склеена с «г.». Это единственный автоматический источник даты акта
    для таких комплектов.
    """
    reg = parse_registry(_ocr("2701-id", 2), pages_total=62)
    assert reg["шапка"]["акты"] == ["222-1-309"]
    assert reg["шапка"]["дата_акта"] == "07.07.2023"


def test_header_line_number_lookalike():
    """«Трубопровод №б» — это линия 6, а не буква."""
    reg = parse_registry(_ocr("2701-id", 2), pages_total=62)
    assert reg["шапка"]["линия"] == "6"


def test_quality_flags_gaps():
    q = quality([{"стр": [1, 2]}, {"стр": [5, 6]}])
    assert q["оценка"] != "высокое"
    assert any("разрыв" in c for c in q["причины"])


def test_quality_empty():
    assert quality([])["оценка"] == "низкое"


# ─── Построение карты целиком (нужен Windows OCR) ──────────────────────

ID_2763 = os.path.join(
    DATASET, "39-40", "п. 2763", "Исполнительная",
    "ИД по Акту ОТН №222-1-179. Линия №108.pdf_v00.100001285300097.pdf")
ID_2775 = os.path.join(
    DATASET, "39-40", "п. 2775", "Исполнительная",
    "ИД по Акту ОТН №222-1-143,121. Линия №122.pdf_v00.100001268600097.pdf")


def _need_real(path):
    from repair.ocr import available
    if not available():
        pytest.skip("Windows OCR на этой машине недоступен")
    if not os.path.isfile(path):
        pytest.skip("обучающей выборки нет на этой машине")


def test_page_map_2763():
    """Проверено рукой: удостоверение о качестве — PDF 6–8, смещение 3."""
    _need_real(ID_2763)
    m = build_page_map(ID_2763)
    assert m["смещение"] == 3
    assert m["страница_реестра"] == 3
    assert m["смещение_источник"] == "копия акта"
    assert quality_cert_section(m)["pdf"] == [6, 8]
    assert find_section(m, SEC_ACT)["pdf"] == [4, 5]


def test_page_map_2775_uses_last_registry_copy():
    """Реестр продублирован (стр. 2 и 3). Смещение считается от последней
    копии: иначе карта уезжает на страницу. Проверено рукой: 8–10."""
    _need_real(ID_2775)
    m = build_page_map(ID_2775)
    assert m["смещение"] == 3
    assert quality_cert_section(m)["pdf"] == [8, 10]


def test_section_pages_limits_range():
    _need_real(ID_2763)
    m = build_page_map(ID_2763)
    pages = section_pages(m, SEC_ACT)
    assert pages == [4, 5]


def test_replacements_table_on_cert_sheet():
    """У ПСК «Стройурал» таблица замены — первый лист удостоверения."""
    _need_real(ID_2763)
    m = build_page_map(ID_2763)
    assert m["замена_страница"] == 6
    assert m["удостоверение_страница"] == 6


# ─── Раскладка «ТП Инжиниринг» ─────────────────────────────────────────

ID_2701 = os.path.join(
    DATASET, "36-30", "п. 2701", "Исполнительная", "36-30, №6, 222-1-309.pdf")
ID_2714 = os.path.join(
    DATASET, "36-30", "п. 2714", "Исполнительная", "36-30, №53, 222-1-17.pdf")


def test_page_map_2701_offset_from_section_order():
    """
    У «ТП Инжиниринг» реестр входит в собственную нумерацию (стр. 1), поэтому
    копия акта — стр. 2, а смещение 1. Полагаться на строку акта из таблицы
    нельзя: номер документа «№1-6» попадает в колонку страниц соседней
    позиции, та получает «стр. 1», и смещение уезжает на единицу.
    """
    _need_real(ID_2701)
    m = build_page_map(ID_2701)
    assert m["смещение"] == 1
    assert m["проверка_смещения"]["страница_акта"] == 3


def test_replacements_table_in_appendix():
    """
    У «ТП Инжиниринг» таблица замены вынесена из удостоверения в приложение
    («См. Приложение №1»), поэтому лист удостоверения и лист с таблицей —
    разные. Проверено глазами: таблица на стр. 8.
    """
    _need_real(ID_2701)
    m = build_page_map(ID_2701)
    assert m["замена_страница"] == 8
    assert m["удостоверение_страница"] != m["замена_страница"]


def test_set_without_quality_certificate():
    """
    В п.2714 удостоверения о качестве в реестре нет вовсе — значит и таблицы
    замены нет. Правильный ответ «не найдено», а не случайная страница.
    """
    _need_real(ID_2714)
    m = build_page_map(ID_2714)
    assert m.get("замена_страница") is None
    assert m.get("удостоверение_страница") is None
