# -*- coding: utf-8 -*-
"""
Карта комплекта исполнительной документации (правило Н1).

Задача: не читать комплект наугад. Раньше `RepairOcrWorker` распознавал первые
шесть страниц — на комплекте в 131 страницу это лотерея. Здесь строится карта
разделов, и программа идёт сразу в нужный документ (копию акта, удостоверение
о качестве ремонта, блок подписей реестра).

Как это работает:

1. Страница реестра ищется **по структуре, а не по слову «Реестр»**: страницей
   реестра считается та, на которой разбирается таблица минимум с тремя
   строками, где есть номера страниц. Слово OCR коверкает («Ресстр»), и в
   п.2775 из-за этого вторая копия реестра терялась, а смещение уезжало на 1.
2. Реестр в комплекте бывает продублирован — берётся ПОСЛЕДНЯЯ копия.
3. Смещение (правило Н1):
   `смещение = страница последней копии реестра − «№ стр.» самого реестра
   в его же таблице` (если реестра в списке нет — второе слагаемое 0).
4. Смещение **проверяется зондом**: распознаётся предсказанная первая страница
   характерного раздела и сверяется по смыслу. Если не сходится — пробуются
   смещения ±1. Проверка нужна потому, что число копий реестра и наличие
   титульного листа зависят от подрядчика.

Замеры и обоснования — в `для обучени/РАЗМЕТКА.md` (правило Н1, §6, §7).
"""

from __future__ import annotations

import re

from repair.doc_extract import normalize_numero

# Колонки шапки таблицы реестра. Ключ → регулярка по тексту ячейки-заголовка.
# «стр» ищется как отдельное слово: подстрока «стр» есть и в «СТРОЙУРАЛ»,
# и в «Реестр», и якорь колонки уезжал в середину листа.
COL_KEYS = (
    ("п_п", r"п\s*/\s*п"),
    ("название", r"наимен"),          # OCR пишет «Наимененанне»
    ("номер_документа", r"чертеж"),
    ("организация", r"организац"),
    ("листов", r"кол\W*во|листов"),
    ("стр", r"^\W*(?:№|N|М)?\s*стр\.?\s*$"),
)

# Шаблон раздела с копией акта. Искать разделы по названию из реестра вообще
# ненадёжно (OCR коверкает: «Удои оверение о качес ' ве ремон ta»), поэтому
# основные листы находятся по их собственным печатным заголовкам —
# ACT_LANDMARK / CERT_LANDMARK / REPL_LANDMARK ниже.
SEC_ACT = r"акт\s+отн|копия\s+акта|ревизии\s+и\s+отбраковк"

MAX_SCAN = 5            # сколько первых страниц смотреть в поисках реестра
MIN_ROWS = 3            # минимум строк с номерами страниц для полноценной карты
REG_MIN_SCORE = 4       # порог «это страница реестра» (см. registry_score)


# ─── Разбор таблицы реестра по координатам слов ────────────────────────

def _cells(res: dict) -> list[dict]:
    """Ячейки страницы: строка OCR = ячейка таблицы (Windows OCR режет
    таблицу по ячейкам, а не по строкам)."""
    out = []
    for ln in res.get("lines", []):
        words = ln.get("words") or []
        if not words:
            continue
        x0 = min(w["x"] for w in words)
        x1 = max(w["x"] + w["w"] for w in words)
        out.append({
            "text": normalize_numero(ln["text"]).strip(),
            "x0": x0, "x1": x1, "xc": (x0 + x1) / 2,
            "y": sum(w["y"] + w["h"] / 2 for w in words) / len(words),
            "h": max(w["h"] for w in words),
        })
    out.sort(key=lambda c: c["y"])
    return out


def _columns(cells: list[dict], width: int,
             height: int) -> tuple[dict, float, set]:
    """x-центры колонок по шапке таблицы, y строки шапки и какие заголовки
    реально найдены (по ним судят, реестр это или другая таблица)."""
    anchors, ys = {}, {}
    for c in cells:
        if c["y"] > height * 0.45:      # шапка таблицы — в верхней части листа
            continue
        low = c["text"].lower()
        for name, pattern in COL_KEYS:
            # Берём самое нижнее совпадение: слово «наименование» встречается
            # выше в подписи «(наименование объекта)».
            if re.search(pattern, low):
                anchors[name] = c["xc"]
                ys[name] = c["y"]
    found = set(anchors)
    if "название" not in anchors or "стр" not in anchors:
        anchors.setdefault("название", width * 0.40)
        anchors.setdefault("листов", width * 0.76)
        anchors.setdefault("стр", width * 0.88)
    anchors.setdefault("п_п", width * 0.10)
    return anchors, max(ys.values(), default=0), found


def parse_registry(res: dict, pages_total: int = 0) -> dict:
    """
    OCR страницы (с рамками слов) → таблица реестра.

    Возвращает {"строки", "смещение", "реестр_в_списке", "шапка", "качество"}.
    """
    width, height = res.get("size", [1, 1])
    cells = _cells(res)
    anchors, head_y, found_cols = _columns(cells, width, height)
    right_edge = width * 0.62

    def nearest(c, names):
        cand = {k: v for k, v in anchors.items() if k in names}
        if not cand:
            return None
        return min(cand.items(), key=lambda kv: abs(kv[1] - c["xc"]))[0]

    def col_of(c):
        # Назначать колонку по ближайшему центру заголовка нельзя: заголовок
        # «Наименование документов» центрирован, а короткие названия жмутся
        # к левому краю и оказывались ближе к колонке «№ п/п».
        if re.fullmatch(r"\d{1,2}", c["text"]) and c["x1"] <= width * 0.16:
            return "п_п"
        if c["xc"] >= right_edge:
            return nearest(c, ("листов", "стр", "организация")) or "стр"
        return nearest(c, ("название", "номер_документа")) or "название"

    body = [c for c in cells if c["y"] > head_y + 5]
    stop_y = next((c["y"] for c in body
                   if re.search(r"выданы замечани|замечания устранен|"
                                r"объем выполненн", c["text"], re.IGNORECASE)),
                  None)
    if stop_y:
        body = [c for c in body if c["y"] < stop_y - 5]

    named = [c for c in body if col_of(c) == "название" and len(c["text"]) >= 4]
    others = [c for c in body if col_of(c) != "название"]
    if not named:
        return {"строки": [], "смещение": None, "реестр_в_списке": False,
                "шапка": {}, "колонки_найдены": sorted(found_cols),
                "качество": {"оценка": "низкое",
                             "причины": ["таблица не найдена"]}}

    # Числа привязываются к БЛИЖАЙШЕМУ по вертикали названию: в бланках части
    # подрядчиков числа выровнены по верху ячейки, а название по центру, и
    # полосовая привязка отдавала их соседней строке.
    bands: dict[int, list[dict]] = {i: [] for i in range(len(named))}
    for c in others:
        i = min(range(len(named)), key=lambda k: abs(named[k]["y"] - c["y"]))
        if abs(named[i]["y"] - c["y"]) <= max(named[i]["h"] * 2.5, 40):
            bands[i].append(c)

    max_page = pages_total or 999

    def sane(vals, top):
        return [v for v in vals if 1 <= v <= top]

    rows = []
    for i, nc in enumerate(named):
        band = bands[i]
        item = {"название": nc["text"], "_y": nc["y"], "_h": nc["h"]}
        m = re.search(r"\b(\d{1,2})\b",
                      " ".join(c["text"] for c in band if col_of(c) == "п_п"))
        if m:
            item["n"] = int(m.group(1))
        org = [c["text"] for c in band if col_of(c) == "организация"]
        if org:
            item["организация"] = " ".join(org)
        # Границы правдоподобия: в колонки чисел попадают номера сертификатов
        # («№СЕРТ-082-5530/23» → 82, 5530, 23).
        sheets = sane([int(x) for x in re.findall(
            r"\d+", " ".join(c["text"] for c in band
                             if col_of(c) == "листов"))], max_page)
        if sheets:
            item["листов"] = sheets[0]
        pg = sorted((c for c in band if col_of(c) == "стр"),
                    key=lambda c: c["xc"])
        pages = sane([int(x) for c in pg
                      for x in re.findall(r"\d+", c["text"])], max_page)
        if pages:
            item["стр"] = [pages[0], pages[-1]]
        rows.append(item)

    # Перенос названия на вторую строку: у продолжения нет ни номера, ни чисел.
    merged: list[dict] = []
    for item in rows:
        empty = ("n" not in item and not item.get("стр")
                 and not item.get("листов"))
        if (merged and empty
                and 0 < item["_y"] - merged[-1]["_y"] < item["_h"] * 1.8):
            merged[-1]["название"] += " " + item["название"]
            continue
        merged.append(item)
    rows = merged

    # Перенос, «прилипший» к числам соседа: тот же номер и тот же диапазон.
    dedup: list[dict] = []
    for item in rows:
        p = dedup[-1] if dedup else None
        if p and item.get("n") == p.get("n") and item.get("стр") == p.get("стр"):
            p["название"] += " " + item["название"]
            continue
        dedup.append(item)
    rows = dedup

    for i, item in enumerate(rows):
        item.pop("_y", None)
        item.pop("_h", None)
        if "n" not in item:
            prev = rows[i - 1].get("n") if i else None
            item["n"] = (prev + 1) if prev else i + 1
            item["n_восстановлен"] = True

    # Восстановление диапазонов «№ стр.»: разделы идут подряд, поэтому начало
    # следующего = конец предыдущего + 1. OCR теряет то левое, то правое число
    # диапазона, а «Кол-во листов» позволяет диапазон достроить.
    expected = 1
    for item in rows:
        rng = item.get("стр")
        if not rng:
            continue
        start, end = rng
        sheets = item.get("листов")
        if sheets and end - start + 1 != sheets:
            start = end - sheets + 1
            item["стр_восстановлены"] = True
        if start != expected and expected <= end:
            start = expected
            item["стр_восстановлены"] = True
        if sheets and end < start + sheets - 1:
            end = start + sheets - 1
            item["стр_восстановлены"] = True
        item["стр"] = [start, end]
        if not sheets:
            item["листов_вычислено"] = end - start + 1
        expected = end + 1

    _fill_missing_ranges(rows)

    # Строку самого реестра ищем нечётко: у «ТП Инжиниринг» слово «Реестр»
    # уходит в заголовок листа, и в таблице от названия остаётся только
    # «исполнительной документации».
    own = next((r for r in rows if r.get("стр")
                and _similar(r["название"], REGISTRY_TITLE) >= 80), None)
    page = res.get("page", 0)
    last_end = max((r["стр"][1] for r in rows if r.get("стр")), default=0)
    cands = []
    if own:
        cands.append(page - own["стр"][0])
    cands.append(page)
    valid = [c for c in cands if 0 <= c <= 10]
    fits = [c for c in valid
            if not last_end or not pages_total or last_end + c <= pages_total]
    offset = (fits or valid or [page])[0]

    заголовок = " ".join(c["text"] for c in cells
                         if c["y"] <= (head_y or 1e9))[:300]
    return {"строки": rows, "смещение": offset, "смещение_кандидаты": cands,
            "реестр_в_списке": bool(own),
            "реестр_стр": own["стр"] if own else None,
            "шапка": _header(cells, head_y),
            "колонки_найдены": sorted(found_cols),
            "заголовок_листа": заголовок,
            "похожесть_заголовка": _similar(заголовок, REGISTRY_TITLE),
            "качество": quality(rows)}


# Цифры, которые OCR подменяет похожими буквами: «222-1-ЗО9» вместо
# «222-1-309», «Трубопровод №б» вместо «№6».
_LOOKALIKES = str.maketrans({"З": "3", "з": "3", "О": "0", "о": "0",
                             "Б": "6", "б": "6", "l": "1", "I": "1",
                             "|": "1", "О".lower(): "0"})


def _fix_digit_lookalikes(text: str) -> str:
    """Заменить буквы-двойники цифр. Применяется только к копии текста,
    из которой извлекаются числа: для обычных слов такая замена бессмысленна,
    но и безвредна."""
    return text.translate(_LOOKALIKES)


def _fill_missing_ranges(rows: list[dict]) -> list[dict]:
    """
    Достроить пропущенные диапазоны «№ стр.».

    У «ТП Инжиниринг» OCR регулярно не читает номера страниц у первых позиций
    (сам реестр, копия акта, удостоверение о качестве) — эти разделы выпадали
    из карты целиком. Два безопасных вывода:

    * реестр — первая позиция списка, значит его страница 1 (этого достаточно,
      чтобы смещение по правилу Н1 посчиталось верно);
    * если между двумя известными разделами пропущено ровно столько позиций,
      сколько страниц в промежутке, — на каждую приходится по одной странице.

    Где однозначного вывода нет, диапазон остаётся пустым: угадывать нельзя.
    """
    if (rows and not rows[0].get("стр")
            and _similar(rows[0]["название"], REGISTRY_TITLE) >= 80):
        rows[0]["стр"] = [1, 1]
        rows[0]["стр_восстановлены"] = True

    i = 0
    while i < len(rows):
        if rows[i].get("стр"):
            i += 1
            continue
        j = i
        while j < len(rows) and not rows[j].get("стр"):
            j += 1
        prev_end = rows[i - 1]["стр"][1] if i and rows[i - 1].get("стр") else 0
        next_start = rows[j]["стр"][0] if j < len(rows) else None
        сколько = j - i
        if next_start and next_start - prev_end - 1 == сколько:
            for t, r in enumerate(rows[i:j]):
                стр = prev_end + 1 + t
                r["стр"] = [стр, стр]
                r["стр_восстановлены"] = True
        i = j + 1
    return rows


def _header(cells: list[dict], head_y: float) -> dict:
    """Шапка листа реестра: номера и дата акта, линия, установка."""
    text = normalize_numero(" ".join(c["text"] for c in cells
                                     if not head_y or c["y"] <= head_y))
    чистый = _fix_digit_lookalikes(text)
    out = {}
    acts = re.findall(r"(?<![\d-])(\d{3}-\d(?:-\d+)+)(?![\d-])", чистый)
    if acts:
        out["акты"] = list(dict.fromkeys(acts))
    # Не `\b` на конце: в «07.07.2023г.» между цифрой и «г» границы слова нет,
    # и дата не находилась вовсе.
    d = re.search(r"(\d{1,2}[.,]\d{1,2}[.,]\d{4})(?!\d)", чистый)
    if d:
        out["дата_акта"] = d.group(1).replace(",", ".")
    # Слова-якоря ищем в исходном тексте: замена «о»→«0» превращает
    # «Трубопровод» в «Труб0пр0в0д», и якорь перестаёт находиться. Подмену
    # применяем только к захваченному номеру.
    номер = r"([\dЗзОоБбlI|]+(?:-[А-Яа-яЁёA-Za-z]+)?)"
    m = (re.search(r"[Тт]рубопровод\w*\s*№\s*" + номер, text)
         or re.search(r"[Лл]ини\w*\s*№\s*" + номер, text))
    if m:
        out["линия"] = _fix_digit_lookalikes(m.group(1))
    m = re.search(r"уст\.?\s*(\d{2}-\d{2})", text, re.IGNORECASE)
    if m:
        out["установка"] = m.group(1)
    return out


def quality(rows: list[dict]) -> dict:
    """
    Самооценка разбора: разделы реестра идут подряд и без пропусков — это и
    проверяется. По оценке видно, каким строкам можно верить.
    """
    with_pages = [r for r in rows if r.get("стр")]
    q = {"строк": len(rows), "со_страницами": len(with_pages)}
    if not with_pages:
        return dict(q, оценка="низкое",
                    причины=["ни у одной строки нет номера страницы"])

    причины = []
    starts = [r["стр"][0] for r in with_pages]
    if starts != sorted(starts):
        причины.append("диапазоны страниц не монотонны")
    разрывы = sum(1 for a, b in zip(with_pages, with_pages[1:])
                  if b["стр"][0] != a["стр"][1] + 1)
    if разрывы:
        причины.append(f"разрывов между разделами: {разрывы}")
    доля = len(with_pages) / max(1, len(rows))
    if доля < 0.8:
        причины.append(f"страницы распознаны у {доля:.0%} строк")

    q["причины"] = причины
    q["оценка"] = ("высокое" if not причины
                   else ("среднее" if len(причины) == 1 else "низкое"))
    return q


# ─── Построение карты комплекта ───────────────────────────────────────

REGISTRY_TITLE = "реестр исполнительной документации"


def registry_score(reg: dict) -> int:
    """
    Насколько страница похожа на реестр ИД.

    Признака «есть таблица с числами» мало: таблицы есть и в акте, и в схеме
    дефектного участка — реестром признавалась случайная страница, и карта
    уезжала. Самый сильный признак — заголовок листа: у реестра похожесть на
    «Реестр исполнительной документации» 85–100 %, у титульного листа
    («ИСПОЛНИТЕЛЬНАЯ ДОКУМЕНТАЦИЯ») 76–79 %, у остальных страниц ≤ 55 %.

        +4 / +3 / +1   похожесть заголовка ≥ 90 / ≥ 80 / ≥ 70
        +3 / +1        найдены колонки «Наименование…» и «№ стр.» / одна из них
        +2 / +1        не меньше пяти строк с номерами страниц / хотя бы одна
        +1             диапазоны страниц монотонны

    Порог `REG_MIN_SCORE` = 4: реальные реестры выборки набирают 4–9,
    посторонние страницы — не больше 2.
    """
    score = 0
    sim = _similar(reg.get("заголовок_листа") or "", REGISTRY_TITLE)
    score += 4 if sim >= 90 else (3 if sim >= 80 else (1 if sim >= 70 else 0))

    cols = set(reg.get("колонки_найдены") or ())
    if {"название", "стр"} <= cols:
        score += 3
    elif cols & {"название", "стр"}:
        score += 1

    with_pages = [r for r in reg["строки"] if r.get("стр")]
    score += 2 if len(with_pages) >= 5 else (1 if with_pages else 0)
    starts = [r["стр"][0] for r in with_pages]
    if starts and starts == sorted(starts):
        score += 1
    return score


def _looks_like_registry(reg: dict) -> bool:
    """
    Порог «это страница реестра». Похожесть заголовка обязательна: без неё
    реестром признавался лист акта со таблицей замены — там тоже есть числовая
    таблица, и по одним структурным признакам он набирал проходной счёт.
    """
    return (reg.get("похожесть_заголовка", 0) >= 70
            and registry_score(reg) >= REG_MIN_SCORE)


def _usable(reg: dict) -> bool:
    """Хватает ли разбора, чтобы строить карту разделов."""
    return sum(1 for r in reg["строки"] if r.get("стр")) >= MIN_ROWS


ACT_LANDMARK = "акт ревизии и отбраковки трубопровода"
ACT_LANDMARK_MIN = 72
# Заголовок удостоверения печатный и одинаковый у всех подрядчиков — по нему
# раздел находится даже там, где таблица реестра разобралась плохо.
CERT_LANDMARK = "удостоверение о качестве ремонта трубопровода"
CERT_LANDMARK_MIN = 70
CERT_SEARCH_SPAN = 8

# Заголовок таблицы замены. Это главный ориентир для данных о ремонте: у ПСК
# «Стройурал» таблица стоит прямо в удостоверении, а у «ТП Инжиниринг»
# вынесена в приложение к нему («См. Приложение №1») — заголовок одинаковый,
# поэтому одно правило закрывает обе раскладки.
REPL_LANDMARK = "ремонту подвергались следующие участки трубопровода"
REPL_LANDMARK_MIN = 72
REPL_SEARCH_SPAN = 10


def _act_row(rows: list[dict]) -> dict | None:
    """Строка реестра с копией акта — первый раздел комплекта."""
    # Строку самого реестра исключаем сразу: при слипании названий в неё
    # попадает и «Копия акта», и тогда она принималась за раздел акта —
    # смещение уезжало на единицу (п.2701).
    rows = [r for r in rows if _similar(r["название"], REGISTRY_TITLE) < 80]
    for r in rows:
        if r.get("стр") and re.search(SEC_ACT, r["название"], re.IGNORECASE):
            return r
    # Названия приходят с огрехами OCR — если шаблон не сработал, берём
    # раздел с наименьшей страницей: копия акта всегда идёт первой.
    прочие = [r for r in rows if r.get("стр")]
    return min(прочие, key=lambda r: r["стр"][0]) if прочие else None


def _find_act_page(read, reg_page: int, total: int) -> int | None:
    """
    Найти в комплекте страницу с копией акта ревизии и отбраковки.

    Это самый надёжный ориентир: заголовок «Акт ревизии и отбраковки
    трубопровода № …» уникален для комплекта и опознаётся даже с огрехами OCR.
    Смещение, посчитанное от него, не требует угадывания числа копий реестра
    и титульных листов.
    """
    for page in range(reg_page + 1, min(total, reg_page + 5) + 1):
        text = read(page)
        if _similar(text, ACT_LANDMARK) >= ACT_LANDMARK_MIN:
            return page
    return None


def _find_cert_page(read, start: int, total: int,
                    span: int = CERT_SEARCH_SPAN) -> int | None:
    """
    Найти лист «Удостоверение о качестве ремонта трубопровода» перебором
    страниц после копии акта.

    Нужно там, где таблица реестра разобралась плохо: удостоверение — главный
    источник типоразмеров и марки стали, и терять его из-за нечитаемого номера
    страницы нельзя.

    Проход первый — строгий: кроме заголовка требуется слово «Цех», которое
    есть в бланке удостоверения и отсутствует в акте. Только по похожести
    заголовка бывают ложные срабатывания: на скудном тексте листа акта
    partial_ratio даёт высокий процент на пустом месте (п.2701, стр. 5).
    Если строгий проход ничего не дал, повторяем без требования «Цех».
    """
    страницы = range(start, min(total, start + span) + 1)
    тексты = {}
    for page in страницы:
        тексты[page] = read(page)
        текст = тексты[page]
        if (_similar(текст, CERT_LANDMARK) >= CERT_LANDMARK_MIN
                and "цех" in текст.lower()):
            return page
    for page, текст in тексты.items():
        if _similar(текст, CERT_LANDMARK) >= CERT_LANDMARK_MIN:
            return page
    return None


def _find_replacements_page(read, start: int, total: int,
                            span: int = REPL_SEARCH_SPAN) -> int | None:
    """
    Найти лист с таблицей «Ремонту подвергались следующие участки
    трубопровода» — там элементы, типоразмеры, характер ремонта и марка стали.

    У ПСК «Стройурал» это первый лист удостоверения о качестве, у «ТП
    Инжиниринг» — отдельное приложение к нему. Ищется по заголовку, поэтому
    раскладка подрядчика значения не имеет.
    """
    лучший, лучшая = None, 0
    for page in range(start, min(total, start + span) + 1):
        s = _similar(read(page), REPL_LANDMARK)
        if s > лучшая:
            лучший, лучшая = page, s
    return лучший if лучшая >= REPL_LANDMARK_MIN else None


def _similar(text: str, phrase: str) -> int:
    from rapidfuzz import fuzz
    return int(fuzz.partial_ratio(phrase.lower(), text.lower()))


def build_page_map(pdf_path: str, log=None) -> dict | None:
    """
    Карта комплекта ИД: где реестр, какое смещение, какие разделы на каких
    страницах PDF. Возвращает None, если реестр не найден (это не ошибка —
    бывают комплекты без реестра).
    """
    from repair.ocr import ocr_page, plain_text
    from repair.pdf_preview import page_count

    def say(msg):
        if log:
            log(msg)

    total = page_count(pdf_path)
    scanned, кандидаты = {}, []
    for n in range(1, min(MAX_SCAN, total) + 1):
        res = ocr_page(pdf_path, n, try_all=True)
        scanned[n] = res
        cand = parse_registry(res, total)
        if _looks_like_registry(cand):
            кандидаты.append((registry_score(cand), n, cand))
    if not кандидаты:
        say("Реестр ИД не найден на первых страницах — карта не построена")
        return None
    # Реестр бывает продублирован, и смещение по правилу Н1 считается именно
    # от ПОСЛЕДНЕЙ копии — брать «лучшую по разбору» нельзя: у п.2763 первая
    # копия читается чуть чище, и смещение уезжало на 1.
    годные = [c for c in кандидаты if _usable(c[2])]
    _, reg_page, reg = (годные or кандидаты)[-1]

    say(f"Реестр ИД на стр. {reg_page}: строк {len(reg['строки'])}, "
        f"качество разбора «{reg['качество']['оценка']}»")

    def read(page: int) -> str:
        res = scanned.get(page) or ocr_page(pdf_path, page, try_all=True)
        scanned[page] = res
        return plain_text(res)

    if not _usable(reg):
        # Таблицу реестра прочитать не удалось — но нужные листы всё равно
        # находятся по их собственным печатным заголовкам, без карты разделов.
        say("Таблица реестра не разобрана — ищу нужные листы по заголовкам")
        out = {"правило": "Н1", "страница_реестра": reg_page, "смещение": None,
               "разделы": [], "качество": reg["качество"],
               "шапка": reg["шапка"], "страниц": total,
               "причина": "таблица реестра не разобрана"}
        акт = _find_act_page(read, reg_page, total)
        if акт:
            out["акт_страница"] = акт
            say(f"Копия акта — стр. {акт}")
        удост = _find_cert_page(read, (акт or reg_page) + 1, total)
        if удост:
            out["удостоверение_страница"] = удост
            say(f"Удостоверение о качестве — стр. {удост}")
        замена = _find_replacements_page(read, удост or (акт or reg_page), total)
        if замена:
            out["замена_страница"] = замена
            say(f"Таблица «Ремонту подвергались…» — стр. {замена}")
        out["распознано"] = sorted(scanned)
        return out

    # Смещение по формуле Н1 — предположение; ориентир по копии акта — факт.
    offset = reg["смещение"] or 0
    источник = "формула Н1"
    подтверждение = None

    act_row = _act_row(reg["строки"])
    act_page = _find_act_page(read, reg_page, total)
    # Номер страницы копии акта по реестру. Разделы идут подряд, а копия акта —
    # первый раздел, поэтому если реестр входит в собственную нумерацию, акт
    # начинается сразу за ним. Этот вывод надёжнее строки из таблицы: у п.2701
    # номер документа «№1-6» попадал в колонку страниц соседней позиции, та
    # получала «стр. 1», принималась за акт, и смещение уезжало на единицу.
    if reg.get("реестр_в_списке") and reg.get("реестр_стр"):
        стр_акта = reg["реестр_стр"][1] + 1
    elif act_row:
        стр_акта = act_row["стр"][0]
    else:
        стр_акта = 1
    по_акту = (act_page - стр_акта) if act_page else None
    # Титул и копии реестра — единицы страниц. Смещение вне 0..10 означает,
    # что строка реестра прочитана неверно (у п.2838 выходило −21).
    if по_акту is not None and not (0 <= по_акту <= 10):
        say(f"Смещение по копии акта вышло неправдоподобным ({по_акту}) — "
            f"строка реестра прочитана неверно, беру формулу Н1")
        по_акту = None
    if по_акту is not None:
        подтверждение = {"ориентир": "копия акта ревизии и отбраковки",
                         "страница_акта": act_page,
                         "раздел_акта": (act_row["название"][:60] if act_row
                                         else "выведен из порядка разделов"),
                         "стр_по_реестру": стр_акта,
                         "смещение_по_формуле": offset,
                         "подтверждено": True}
        if по_акту != offset:
            say(f"Смещение уточнено по копии акта (стр. {act_page}): "
                f"{offset} → {по_акту}")
        offset, источник = по_акту, "копия акта"
    else:
        say("Копию акта в комплекте найти не удалось — смещение взято по "
            "формуле Н1 и не подтверждено")
        подтверждение = {"ориентир": None, "подтверждено": False,
                         "смещение_по_формуле": offset}

    разделы, отброшено, prev_end = [], 0, 0
    for row in reg["строки"]:
        rng = row.get("стр")
        if not rng:
            continue
        pdf = [rng[0] + offset, rng[1] + offset]
        if not (1 <= pdf[0] <= pdf[1] <= total) or pdf[0] <= prev_end:
            отброшено += 1
            continue
        разделы.append({"название": row["название"], "листов": row.get("листов"),
                        "стр": rng, "pdf": pdf})
        prev_end = pdf[1]

    out = {"правило": "Н1", "страница_реестра": reg_page, "смещение": offset,
           "смещение_источник": источник,
           "разделы": разделы, "качество": reg["качество"],
           "шапка": reg["шапка"], "страниц": total,
           "распознано": sorted(scanned)}
    if отброшено:
        out["отброшено_разделов"] = отброшено
    if подтверждение:
        out["проверка_смещения"] = подтверждение

    # Удостоверение о качестве — источник замены и марки стали. Если таблица
    # реестра его не дала или разобралась ненадёжно, ищем лист напрямую по
    # печатному заголовку.
    из_таблицы = quality_cert_section(out)
    надёжно = (из_таблицы is not None
               and reg["качество"]["оценка"] == "высокое")
    if not надёжно:
        начало = (act_page + 1) if act_page else (reg_page + 1)
        найдено = _find_cert_page(read, начало, total)
        if найдено:
            out["удостоверение_страница"] = найдено
            if из_таблицы and из_таблицы["pdf"][0] != найдено:
                say(f"Удостоверение о качестве: по реестру стр. "
                    f"{из_таблицы['pdf'][0]}, по заголовку — {найдено}; "
                    f"беру найденное по заголовку")
            else:
                say(f"Удостоверение о качестве найдено по заголовку: "
                    f"стр. {найдено}")
        elif из_таблицы:
            out["удостоверение_страница"] = из_таблицы["pdf"][0]
    elif из_таблицы:
        out["удостоверение_страница"] = из_таблицы["pdf"][0]

    # Таблица замены — то, ради чего комплект и читается. У части подрядчиков
    # она вынесена из удостоверения в приложение, поэтому ищем её отдельно.
    начало_замены = out.get("удостоверение_страница") or (act_page or reg_page)
    замена = _find_replacements_page(read, начало_замены, total)
    if замена:
        out["замена_страница"] = замена
        if замена != out.get("удостоверение_страница"):
            say(f"Таблица «Ремонту подвергались…» — на стр. {замена} "
                f"(вынесена из удостоверения в приложение)")
    else:
        say("Таблицу «Ремонту подвергались…» найти не удалось — данные о "
            "замене нужно внести вручную")
    return out


def section_pages(page_map: dict, pattern: str, limit: int = 8) -> list[int]:
    """Страницы PDF разделов, чьё название матчит регулярку."""
    out = []
    for sec in (page_map or {}).get("разделы", []):
        if not re.search(pattern, sec["название"], re.IGNORECASE):
            continue
        first, last = sec["pdf"]
        out.extend(range(first, min(last, first + limit - 1) + 1))
    return out


def find_section(page_map: dict, pattern: str) -> dict | None:
    """Первый раздел, чьё название матчит регулярку."""
    for sec in (page_map or {}).get("разделы", []):
        if re.search(pattern, sec["название"], re.IGNORECASE):
            return sec
    return None


# Ядро названия раздела для нечёткого поиска: искать «удостоверение» нельзя
# (в реестре есть «Копия удостоверения сварщика»), а точное «о качестве»
# ломается на OCR-огрехах вида «о каче•лве».
CERT_PHRASE = "о качестве ремонта"
CERT_MIN = 70


def find_section_like(page_map: dict, phrase: str,
                      min_score: int = CERT_MIN) -> dict | None:
    """
    Раздел, чьё название ближе всего к фразе (нечётко, rapidfuzz).

    Нужен потому, что названия в реестре приходят из OCR с выпавшими буквами:
    «Удостоверение о каче•лве ремонта трубопровода», «Удои оверение о качес ' ве
    ремон ta трубопровода» — регулярка их не берёт.
    """
    лучший, лучшая = None, 0
    for sec in (page_map or {}).get("разделы", []):
        s = _similar(sec["название"], phrase)
        if s > лучшая:
            лучший, лучшая = sec, s
    return лучший if лучшая >= min_score else None


def quality_cert_section(page_map: dict) -> dict | None:
    """Раздел «Удостоверение о качестве ремонта трубопровода»."""
    return find_section_like(page_map, CERT_PHRASE)
