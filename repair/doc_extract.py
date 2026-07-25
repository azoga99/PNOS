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
# Терпимый шаблон для копий актов внутри комплекта ИД: там скан хуже, и
# «222-1-179» читается как «222-l.179» — единица становится буквой, дефис
# точкой. Разделители и «l/I/|» приводятся к норме после захвата.
_ACT_LOOSE_RE = re.compile(
    r"(?<![\d-])(\d{3})\s?[-.,]\s?([\dlLIi|])\s?[-.,]\s?(\d{1,4})(?![\d-])")
# хвост «2023г.» пишут слитно — поэтому lookahead на нецифру, а не \b
_DATE_NUM_RE = re.compile(r"\b(\d{2}\.\d{2}\.\d{4})(?!\d)")
# «„24" февраля 2026», «24 » февраля 2026 — кавычки любые, день опционален
_DATE_WORD_RE = re.compile(
    r"[«\"„]?\s*(\d{1,2})?\s*[»\"“]?\s*"
    r"(янва\w*|февра\w*|март\w*|апрел\w*|ма[яй]\w*|июн\w*|июл\w*|август\w*|"
    r"сентябр\w*|октябр\w*|ноябр\w*|декабр\w*)\s+(\d{4})", re.IGNORECASE)
# Установка: «по установке 39-40», но в ИД пишут «ПКМ уст.39-40 тит. 39-40» —
# точка вместо пробела, и старый шаблон установку не находил.
_UNIT_RE = re.compile(r"(?:установк\w*|уст)\.?\s*(\d{2}-\d{2})", re.IGNORECASE)
# Линия/трубопровод. Суффикс из букв обязателен: бывают линии «39-МЦК».
# `(?!\d)` закрывает откат регулярки: без него «222-1-149» давало линию «22»
# (движок отбрасывал цифру, чтобы обойти запрет ниже).
_LINE_TAIL = r"(\d{1,4}(?:-[А-Яа-яЁёA-Za-z]+)?)(?!\d)"
# Тот же отрицательный просмотр, что и ниже: «Акт ревизии и отбраковки
# трубопровода № 222-1-149» — это номер акта, а не линия «222».
_LINE_RE = re.compile(r"(?:трубопровод\w*|лини\w*)\s*№?\s*" + _LINE_TAIL
                      + r"(?!\s*-\s*\d)", re.IGNORECASE)
# Одиночное «№108» — так номер линии подписан в таблице дефектов акта.
# Отрицательный просмотр отсекает начало номера акта: «№222-1-149» — это акт,
# а не линия «222» (нераспознанные акты остаются в тексте и путали разбор).
_LINE_BARE_RE = re.compile(r"№\s*" + _LINE_TAIL + r"(?!\s*-\s*\d)")
# строка таблицы замены: элемент + типоразмер
_ELEMENTS = ("труба", "отвод", "переход", "фланец", "опора", "бобышка",
             "заглушка", "тройник")
# Слова из описания дефекта. Строка «(Коррозионный износ, минимальная толщина
# для перехода Дн108х89 составляет 2,6мм, отбраковочная — 2,0мм)» содержит и
# название элемента, и типоразмер — и раньше попадала в колонку «Замена»
# вместо «Переход 108х6-89х6».
_NOT_REPLACEMENT = ("толщин", "износ", "составля", "отбраков", "дефект",
                    "коррози", "выявлен", "устранен",
                    # строки журнала сварочных работ: те же трубы, но с датой
                    # сварки, клеймом и отметкой годности
                    "годен", "сертиф", "клейм")
# Начало строки таблицы: возможен порядковый номер позиции («1 Труба 273х10»).
_ROW_NUM_RE = re.compile(r"^\s*\d{1,2}\s*[.)]?\s+")
_SIZE_RE = re.compile(r"\d{2,3}\s*[хx×]\s*\d{1,2}(?:[,.]\d)?", re.IGNORECASE)
# Размер без стенки: «Дн45», «Ду80» — так записывают бобышки и фланцы.
_DN_RE = re.compile(r"[ДдDd][нуyНУY]\s?\d{2,3}", re.IGNORECASE)
# Марка стали. В акте пишут «Ст. 20», в удостоверении о качестве — «Сталь 20»:
# нужны оба варианта, иначе материал из удостоверения не подхватывался.
_MATERIAL_RE = re.compile(
    r"(Стал[ья]\s?20|Ст\.?\s?20|12Х18Н10Т|10Х2М|09Г2С)", re.IGNORECASE)


@dataclass
class DocInfo:
    act_numbers: list[str] = field(default_factory=list)
    act_date: str = ""            # ДД.ММ.ГГГГ или .ММ.ГГГГ (день не распознан)
    unit: str = ""
    line: str = ""
    replacements: list[str] = field(default_factory=list)
    material: str = ""
    # Дата выполнения работ по правилу Д1 (самая поздняя в блоке подписей
    # реестра ИД). Обычно пусто: в блоке подписей даты рукописные, и OCR их
    # не читает — тогда сотрудник вписывает её по кропу.
    date_done: str = ""


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


# Знак «№» Windows OCR читает десятком способов — замеры по обучающей
# выборке («для обучени»): N2, Ne, N9, Nel, М•, Х!, Хд, У2, Уб и т.п.
NUMERO_VARIANTS = ("N2", "Ne", "N9", "Nel", "N!", "N°", "No", "Nº",
                   "Х!", "Хд", "Х2", "У2", "Уб", "Ул", "М•", "Ng", "Њ")


# Перечислить все варианты списком не выходит: «№119» встретилось и как
# «*2119». Общая форма — одиночный знак (буква или символ), иногда с прилипшей
# цифрой, прямо перед номером. Знак не должен быть частью слова, иначе
# пострадают обычные тексты.
# Знак обязан стоять отдельно (перед ним пробел, кавычка, начало строки).
# Иначе правило срабатывало внутри типоразмера «273х10» и марки «12Х18Н10Т»,
# превращая их в «273№10» и «12№18Н10Т».
_NUMERO_RE = re.compile(r"(?<!\w)[NnХхУуYyМмMm*#][°º!129]?(?=\s?\d{2,4}(?!\d))")


def normalize_numero(text: str) -> str:
    """Привести все варианты распознавания «№» к самому знаку."""
    out = text
    for v in NUMERO_VARIANTS:
        out = out.replace(v, "№")
    return _NUMERO_RE.sub("№", out)


def _normalize_for_acts(text: str) -> str:
    """Артефакты OCR в номерах: дефис читается как кавычка («222-1„84»),
    «№» прилипает к цифрам («N9222-1-84»)."""
    t = re.sub(r"[„“”]", "-", normalize_numero(text))
    t = re.sub(r"\bN[9gоoО0]\s*(?=\d{3}[-])", "№ ", t)
    # «№» перед номером акта OCR читает как «буква + цифра» и склеивает с ним:
    # «Акт ОТН Y2222-1-179». Из-за лишней ведущей цифры номер акта не
    # опознавался вовсе — отделяем его.
    t = re.sub(r"(?<![\d-])[A-Za-zА-Яа-яЁё][0-9](?=\d{3}-\d)", "№ ", t)
    return t


def is_replacement_line(line: str) -> bool:
    """
    Похожа ли строка на позицию таблицы «подлежащих замене».

    Требования, выведенные из реальных документов:
      * строка НАЧИНАЕТСЯ с названия элемента (допускается номер позиции
        впереди) — «Переход 108х6-89х6», «1 Труба 273х10 26 п. м.»;
      * название сверяется нечётко: OCR глотает буквы («Пе еход»);
      * есть типоразмер вида «108х6»;
      * нет слов из описания дефекта — иначе в реестр попадает строка
        «…минимальная толщина для перехода Дн108х89 составляет 2,6мм…».
    """
    from rapidfuzz import fuzz

    текст = " ".join((line or "").split())
    if not текст:
        return False
    low = текст.lower()
    if any(w in low for w in _NOT_REPLACEMENT):
        return False
    # Типоразмер «108х6» либо диаметр «Дн45» (так пишут бобышки и фланцы).
    if not (_SIZE_RE.search(текст) or _DN_RE.search(текст)):
        return False
    голова = _ROW_NUM_RE.sub("", low)[:14]
    return any(fuzz.partial_ratio(e, голова) >= 80 for e in _ELEMENTS)


def _replacement_key(line: str) -> tuple:
    """Суть позиции: тип элемента и набор типоразмеров — для отсева дублей."""
    from rapidfuzz import fuzz

    low = " ".join(line.split()).lower()
    голова = _ROW_NUM_RE.sub("", low)[:14]
    элемент = max(_ELEMENTS, key=lambda e: fuzz.partial_ratio(e, голова))
    размеры = tuple(sorted(
        m.group(0).replace(" ", "").lower()
        for rx in (_SIZE_RE, _DN_RE) for m in rx.finditer(line)))
    return элемент, размеры


def extract_doc_info(lines: list[str]) -> DocInfo:
    """Строки OCR (можно с нескольких страниц) → черновик полей."""
    info = DocInfo()
    # Нормализация «№» нужна всему разбору, а не только номерам актов: OCR
    # пишет его как N2 / Ne / Х! / У2 / М• и т.п.
    lines = [normalize_numero(ln) for ln in lines]
    text = "\n".join(lines)
    norm = _normalize_for_acts(text)

    seen = set()
    for m in _ACT_RE.finditer(norm):
        if m.group(1) not in seen:
            seen.add(m.group(1))
            info.act_numbers.append(m.group(1))
    # Второй проход терпимым шаблоном: в копиях актов внутри ИД скан хуже.
    for m in _ACT_LOOSE_RE.finditer(norm):
        mid = re.sub(r"[lLIi|]", "1", m.group(2))
        num = f"{m.group(1)}-{mid}-{m.group(3)}"
        if num not in seen:
            seen.add(num)
            info.act_numbers.append(num)

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

    # Линия: сперва «трубопровод №N» / «Линия №N», иначе одиночное «№N» в
    # тексте без номеров актов и без установки (так подписана первая ячейка
    # таблицы дефектов в акте: «№108 «№108 "фильтрат…"»»).
    lm = _LINE_RE.search(text)
    if lm:
        info.line = lm.group(1)
    else:
        rest = norm
        for num in info.act_numbers:
            rest = rest.replace(num, " ")
        if info.unit:
            rest = rest.replace(info.unit, " ")
        lb = _LINE_BARE_RE.search(rest)
        if lb:
            info.line = lb.group(1)

    # Одна и та же позиция приходит несколько раз: зона читается в нескольких
    # масштабах, и OCR всякий раз коверкает её по-своему («Отвод П90 219х12»,
    # «отвед 7190 219х12»). Поэтому дубли ищутся не по тексту, а по сути:
    # тип элемента плюс набор типоразмеров.
    видели = set()
    for ln in lines:
        if not is_replacement_line(ln):
            continue
        текст = " ".join(ln.split())
        ключ = _replacement_key(текст)
        if ключ in видели:
            continue
        видели.add(ключ)
        info.replacements.append(текст)

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
        if info.date_done and not rec.date_done:
            rec.date_done = info.date_done
            filled += 1
        if info.material and not rec.material:
            rec.material = info.material
            filled += 1
        if info.replacements and not rec.replacements:
            rec.replacements = "; ".join(info.replacements)
            filled += 1
    return filled
