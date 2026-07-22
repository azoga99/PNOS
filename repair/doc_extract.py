# -*- coding: utf-8 -*-
"""
Извлечение полей из распознанного текста документов ремонта (Э4, слой 2).

Вход — строки OCR (repair.ocr) со страниц акта/ИД. Выход — черновик полей:
номера и даты актов, установка, линия, строки замены, материал. Всё это
только ПРЕДЗАПОЛНЕНИЕ: сотрудник видит и правит значения в форме реестра.

Замечания по качеству источника:
- «Акт ревизии и отбраковки трубопровода № 222-1-121» — печатный, надёжно;
- дата грифа «„24" февраля 2026 г.» — день часто ВПИСАН ОТ РУКИ и может не
  распознаться: тогда берём «месяц + год», день остаётся на сотруднике;
- таблица «подлежащих замене» (лист 2) печатная: «Труба | 108х6 | 2 п.м...».
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

_MONTHS = {
    "янв": "01", "фев": "02", "мар": "03", "апр": "04", "ма": "05",
    "июн": "06", "июл": "07", "авг": "08", "сен": "09", "окт": "10",
    "ноя": "11", "дек": "12",
}

# OCR читает «№» как «N9»/«No»/«Ng» — привязка к знаку номера ненадёжна,
# ищем сам шаблон номера акта (как в detect.py)
_ACT_RE = re.compile(r"\b(\d{3}-\d(?:-\d+)+)\b")
# хвост «2023г.» пишут слитно — поэтому lookahead на нецифру, а не \b
_DATE_NUM_RE = re.compile(r"\b(\d{2}\.\d{2}\.\d{4})(?!\d)")
# «„24" февраля 2026», «24 » февраля 2026 — кавычки любые, день опционален
_DATE_WORD_RE = re.compile(
    r"[«\"„]?\s*(\d{1,2})?\s*[»\"“]?\s*"
    r"(янва\w*|февра\w*|март\w*|апрел\w*|ма[яй]\w*|июн\w*|июл\w*|август\w*|"
    r"сентябр\w*|октябр\w*|ноябр\w*|декабр\w*)\s+(\d{4})", re.IGNORECASE)
_UNIT_RE = re.compile(r"установк\w*\s+(\d{2}-\d{2})", re.IGNORECASE)
_LINE_RE = re.compile(r"(?:трубопровод\w*|лини\w*)\s*№?\s*(\d{1,4})\b",
                      re.IGNORECASE)
# строка таблицы замены: элемент + типоразмер
_ELEMENTS = ("труба", "отвод", "переход", "фланец", "опора", "бобышка",
             "заглушка")
_SIZE_RE = re.compile(r"\d{2,3}\s*[хx×]\s*\d{1,2}(?:[,.]\d)?", re.IGNORECASE)
_MATERIAL_RE = re.compile(r"(Ст\.?\s?20|12Х18Н10Т|10Х2М|09Г2С)", re.IGNORECASE)


@dataclass
class DocInfo:
    act_numbers: list[str] = field(default_factory=list)
    act_date: str = ""            # ДД.ММ.ГГГГ или .ММ.ГГГГ (день не распознан)
    unit: str = ""
    line: str = ""
    replacements: list[str] = field(default_factory=list)
    material: str = ""


def _month_num(word: str) -> str | None:
    w = word.lower()
    for k, v in _MONTHS.items():
        if w.startswith(k):
            return v
    return None


def _find_word_date(text: str) -> str:
    m = _DATE_WORD_RE.search(text)
    if not m:
        return ""
    day, month_w, year = m.group(1), m.group(2), m.group(3)
    month = _month_num(month_w)
    if not month:
        return ""
    day_s = day.zfill(2) if day else ""
    return f"{day_s}.{month}.{year}" if day_s else f".{month}.{year}"


def _normalize_for_acts(text: str) -> str:
    """Артефакты OCR в номерах: дефис читается как кавычка («222-1„84»),
    «№» прилипает к цифрам («N9222-1-84»)."""
    t = re.sub(r"[„“”]", "-", text)
    t = re.sub(r"\bN[9gоoО0]\s*(?=\d{3}[-])", "№ ", t)
    return t


def extract_doc_info(lines: list[str]) -> DocInfo:
    """Строки OCR (можно с нескольких страниц) → черновик полей."""
    info = DocInfo()
    text = "\n".join(lines)

    seen = set()
    for m in _ACT_RE.finditer(_normalize_for_acts(text)):
        if m.group(1) not in seen:
            seen.add(m.group(1))
            info.act_numbers.append(m.group(1))

    # Дата акта — только из надёжных мест:
    #   1) числовая дата в одной строке с номером акта («…№ 222-1-309 от 07.07.2023»);
    #   2) дата в 1–3 строках после «УТВЕРЖДАЮ» (гриф), исключая строки
    #      с «произведена…» — там дата РЕВИЗИИ, её брать нельзя.
    # Нечитаемый рукописный день грифа → пусто или «.ММ.ГГГГ»: лучше пусто,
    # чем чужая дата — впишет сотрудник по превью.
    for ln in lines:
        if _ACT_RE.search(ln) and "акт" in ln.lower():
            d = _DATE_NUM_RE.search(ln)
            if d:
                info.act_date = d.group(1)
                break
    if not info.act_date:
        for i, ln in enumerate(lines):
            if "утвержда" not in ln.lower():
                continue
            for cand in lines[i + 1:i + 4]:
                low = cand.lower()
                if "произведен" in low or "ревизи" in low:
                    continue
                d = _DATE_NUM_RE.search(cand)
                if d:
                    info.act_date = d.group(1)
                    break
                wd = _find_word_date(cand)
                if wd:
                    info.act_date = wd
                    break
            break

    um = _UNIT_RE.search(text)
    if um:
        info.unit = um.group(1)
    lm = _LINE_RE.search(text)
    if lm:
        info.line = lm.group(1)

    for ln in lines:
        low = ln.lower()
        if any(low.lstrip().startswith(e) or f" {e}" in low[:24]
               for e in _ELEMENTS):
            if _SIZE_RE.search(ln) or "дн" in low:
                info.replacements.append(" ".join(ln.split()))

    mm = _MATERIAL_RE.search(text)
    if mm:
        info.material = mm.group(1)

    return info


def merge_into_records(records, info: DocInfo) -> int:
    """
    Влить распознанное в записи реестра: заполняются ТОЛЬКО ПУСТЫЕ поля
    записей, чьи акты пересекаются с распознанными номерами (ничего не
    перезаписывается — правки сотрудника священны). Возвращает число
    заполненных полей.
    """
    filled = 0
    matched = set(info.act_numbers)
    for rec in records:
        hit = [a for a in rec.acts if a.number in matched]
        if not hit:
            continue
        if info.act_date:
            for a in hit:
                if not a.date:
                    a.date = info.act_date
                    filled += 1
        if info.material and not rec.material:
            rec.material = info.material
            filled += 1
        if info.replacements and not rec.replacements:
            rec.replacements = "; ".join(info.replacements)
            filled += 1
    return filled
