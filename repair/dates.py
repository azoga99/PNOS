# -*- coding: utf-8 -*-
"""
Правила извлечения дат ремонта (Д1 и Д2 из обучающей выборки).

Обоснования и контрпримеры — в `для обучени/РАЗМЕТКА.md`, §5.

**Д1 — дата выполнения работ** (`RepairRecord.date_done`). Берётся самая
поздняя дата в блоке подписей реестра ИД, независимо от того, в какой она
строке. Строка «Сдал» на практике часто пустая (п.2763), а порядок строк
у подрядчиков разный: у ПСК «Стройурал» это Сдал / Проверил / Принял, у
«ТП Инжиниринг» — Сдал / Принял / Проверил. Привязка к строке поэтому
невозможна.

**Д2 — дата акта** (`ActRef.date`): гриф УТВЕРЖДАЮ → дата из имени файла →
пусто. Печатная строка «по установке … произведена ЭПБ трубопровода» НЕ
используется: это дата проведения ЭПБ, и в акте 222-1-121 она расходится с
грифом на шесть дней (18.03.2026 против 24.03.2026).

Важно: даты в блоке подписей почти всегда рукописные или штампованные, и
Windows OCR их не читает вовсе. Поэтому Д1 применима лишь там, где дата
печатная; в остальных случаях функция возвращает None, и программа обязана
показать сотруднику кроп блока подписей (принцип «ручной ввод равноправен»).
"""

from __future__ import annotations

import datetime
import re

_MONTHS = {
    "янв": 1, "фев": 2, "мар": 3, "апр": 4, "ма": 5, "июн": 6, "июл": 7,
    "авг": 8, "сен": 9, "окт": 10, "ноя": 11, "дек": 12,
}

_NUM_RE = re.compile(r"\b(\d{1,2})[.,/](\d{1,2})[.,/](\d{4})\b")
# Штамп «1 5 АПР 2026»: OCR разбивает день пробелом, месяц — тремя буквами.
_STAMP_RE = re.compile(
    r"\b(\d)\s?(\d)?\s+([А-Яа-яЁё]{3})[А-Яа-яЁё]*\.?\s+(\d{4})\b")
_WORD_RE = re.compile(
    r"\b(\d{1,2})\s+(янва\w*|февра\w*|март\w*|апрел\w*|ма[яй]\w*|июн\w*|"
    r"июл\w*|август\w*|сентябр\w*|октябр\w*|ноябр\w*|декабр\w*)\s+(\d{4})",
    re.IGNORECASE)

DATE_FMT = "%d.%m.%Y"


def _month(word: str) -> int | None:
    low = word.lower()
    for key, num in _MONTHS.items():
        if low.startswith(key):
            return num
    return None


def _make(day: int, month: int, year: int) -> datetime.date | None:
    try:
        return datetime.date(year, month, day)
    except ValueError:
        return None


def find_dates(text: str) -> list[datetime.date]:
    """Все даты в тексте: «16.04.2026», «16 апреля 2026», штамп «1 5 АПР 2026»."""
    out: list[datetime.date] = []
    for m in _NUM_RE.finditer(text):
        d = _make(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if d:
            out.append(d)
    for m in _WORD_RE.finditer(text):
        month = _month(m.group(2))
        if month:
            d = _make(int(m.group(1)), month, int(m.group(3)))
            if d:
                out.append(d)
    for m in _STAMP_RE.finditer(text):
        day = int(m.group(1) + (m.group(2) or ""))
        month = _month(m.group(3))
        if month:
            d = _make(day, month, int(m.group(4)))
            if d:
                out.append(d)
    return out


def latest_date(texts) -> str | None:
    """
    Правило Д1: самая поздняя дата в переданных строках (блок подписей
    реестра ИД) в виде «ДД.ММ.ГГГГ». None — дат нет (обычная ситуация:
    они рукописные).
    """
    if isinstance(texts, str):
        texts = [texts]
    все: list[datetime.date] = []
    for t in texts:
        все.extend(find_dates(t))
    return max(все).strftime(DATE_FMT) if все else None


def act_date(from_grif: str = "", filename: str = "",
             from_registry: str = "") -> str:
    """
    Правило Д2: дата акта. Приоритет — гриф УТВЕРЖДАЮ, затем дата из имени
    файла акта, затем печатная шапка реестра ИД («№222-1-309 от 07.07.2023г.»).
    Неполная дата грифа («.03.2026», день от руки не распознан) уступает
    полной дате из любого следующего источника.
    """
    grif = (from_grif or "").strip()
    полная = r"\d{2}\.\d{2}\.\d{4}"
    if re.fullmatch(полная, grif):
        return grif

    if filename:
        from repair.detect import parse_doc_filename
        from_name = parse_doc_filename(filename).date
        if from_name:
            return from_name

    реестр = (from_registry or "").strip()
    if re.fullmatch(полная, реестр):
        return реестр
    return grif          # либо неполная дата грифа, либо пусто
