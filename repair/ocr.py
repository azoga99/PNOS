# -*- coding: utf-8 -*-
"""
Распознавание текста через Windows OCR (этап Э4, слой 2).

Используется встроенный движок Windows 10/11 (Windows.Media.Ocr) через
pywinrt — ничего тяжёлого в поставку не добавляется (пакеты winrt-* весят
единицы мегабайт; при сборке EXE добавить их в окружение).

Качество на печатном тексте актов хорошее; рукописные вставки (день в грифе,
подписи) распознаются ненадёжно — поэтому OCR только ПРЕДЗАПОЛНЯЕТ поля,
подтверждает сотрудник по превью (принцип «ручной ввод равноправен»).

Про повороты. В реальных комплектах ИД повёрнута каждая восьмая страница
(замер по обучающей выборке «для обучени»: 24 из 196 — 270°, 90° и 180°).
Windows OCR ведёт себя при этом по-разному: на 180° выдаёт мусор, а на 90/270°
**молча возвращает пустой список строк** — страница выглядит чистой, и раздел
документа теряется. Поэтому распознавание идёт с автоопределением ориентации
(`ocr_auto` / `ocr_page`), а пустой ответ трактуется как «страница, вероятно,
повёрнута», а не «текста нет».

Функции с координатами слов (`ocr_lines`) нужны для разбора таблиц: реестр ИД
Windows OCR отдаёт по ячейкам, и восстановить строки можно только по рамкам.
"""

from __future__ import annotations

import re

# Слово от трёх русских букв — признак осмысленного текста.
_WORD_RE = re.compile(r"[А-Яа-яЁё]{3,}")
# Частые слова документов ремонта: сильный сигнал верной ориентации.
_HINTS = ("акт", "трубопровод", "документ", "реестр", "лист", "стр",
          "ремонт", "установк", "линия", "контрол", "испытан", "сварочн",
          "удостоверение", "качеств", "исполнительн", "механик", "инженер")
# Ниже этого порога результат считается негодным и пробуются 90/270°.
_SCORE_MIN = 120


def available() -> bool:
    """Доступен ли Windows OCR с русским языком на этой машине."""
    try:
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.globalization import Language
        return (OcrEngine.try_create_from_language(Language("ru")) is not None
                or OcrEngine.try_create_from_user_profile_languages() is not None)
    except Exception:
        return False


async def _ocr_bytes_boxes_async(png_bytes: bytes, lang: str) -> list[dict]:
    """PNG-байты → строки со словами и их рамками."""
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.globalization import Language

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream.get_output_stream_at(0))
    writer.write_bytes(png_bytes)
    await writer.store_async()
    await writer.flush_async()

    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()

    engine = OcrEngine.try_create_from_language(Language(lang)) \
        or OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        raise RuntimeError("Windows OCR недоступен (нет языкового пакета)")

    result = await engine.recognize_async(bitmap)
    out = []
    for line in result.lines:
        words = []
        for w in line.words:
            r = w.bounding_rect
            words.append({"t": w.text, "x": round(r.x), "y": round(r.y),
                          "w": round(r.width), "h": round(r.height)})
        out.append({"text": line.text, "words": words})
    return out


async def _ocr_bytes_async(png_bytes: bytes, lang: str) -> list[str]:
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.globalization import Language

    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream.get_output_stream_at(0))
    writer.write_bytes(png_bytes)
    await writer.store_async()
    await writer.flush_async()

    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()

    engine = OcrEngine.try_create_from_language(Language(lang)) \
        or OcrEngine.try_create_from_user_profile_languages()
    if engine is None:
        raise RuntimeError("Windows OCR недоступен (нет языкового пакета)")

    result = await engine.recognize_async(bitmap)
    return [line.text for line in result.lines]


def ocr_png_bytes(png_bytes: bytes, lang: str = "ru") -> list[str]:
    """PNG-байты → строки распознанного текста."""
    import asyncio
    return asyncio.run(_ocr_bytes_async(png_bytes, lang))


def ocr_array(arr, lang: str = "ru") -> list[str]:
    """numpy-массив RGB (страница из pdf_preview) → строки текста."""
    return ocr_png_bytes(_to_png(arr), lang)


# ─── С координатами слов и автоповоротом ──────────────────────────────

def _to_png(arr) -> bytes:
    import io
    from PIL import Image

    buf = io.BytesIO()
    Image.fromarray(arr).save(buf, format="PNG")
    return buf.getvalue()


def ocr_lines(arr, lang: str = "ru") -> list[dict]:
    """numpy-массив RGB → [{"text": ..., "words": [{"t","x","y","w","h"}]}]."""
    import asyncio
    return asyncio.run(_ocr_bytes_boxes_async(_to_png(arr), lang))


def score_lines(lines: list[dict]) -> int:
    """
    Оценка осмысленности распознанного. Чем выше, тем вероятнее, что
    ориентация верна: считаются длинные русские слова плюс бонус за слова,
    характерные для документов ремонта.
    """
    text = " ".join(ln["text"] for ln in lines)
    if not text.strip():
        return 0
    score = sum(len(m.group(0)) for m in _WORD_RE.finditer(text))
    low = text.lower()
    return score + 25 * sum(1 for h in _HINTS if h in low)


def _rot(arr, deg: int):
    import numpy as np
    return arr if deg == 0 else np.rot90(arr, k=deg // 90)


def ocr_auto(arr, lang: str = "ru", try_all: bool = False) -> dict:
    """
    Распознать изображение, сам подобрав ориентацию.

    Возвращает {"rotation", "score", "scores", "size", "lines"}. Сначала
    пробуются 0° и 180°; если обе дают слабый результат — 90° и 270°
    (страницы-альбомы, на которых OCR обычно молчит).
    """
    h, w = arr.shape[:2]
    results: dict[int, dict] = {}

    def attempt(deg: int) -> dict:
        if deg not in results:
            try:
                lines = ocr_lines(_rot(arr, deg), lang)
            except Exception:
                lines = []
            results[deg] = {"lines": lines, "score": score_lines(lines)}
        return results[deg]

    for deg in ((0, 180, 90, 270) if try_all else (0, 180)):
        attempt(deg)
        best = max(results, key=lambda d: results[d]["score"])
        if not try_all and results[best]["score"] >= _SCORE_MIN:
            break
    best = max(results, key=lambda d: results[d]["score"])
    if results[best]["score"] < _SCORE_MIN:
        for deg in (90, 270):
            attempt(deg)
        best = max(results, key=lambda d: results[d]["score"])

    return {
        "rotation": best,
        "score": results[best]["score"],
        "scores": {d: r["score"] for d, r in sorted(results.items())},
        "size": [w, h] if best in (0, 180) else [h, w],
        "lines": results[best]["lines"],
    }


def ocr_page(pdf_path: str, page_no: int, scale: float = 2.2,
             lang: str = "ru", try_all: bool = False) -> dict:
    """Страница PDF (1-базный номер) → результат `ocr_auto` + номер страницы."""
    from repair.pdf_preview import render_page

    res = ocr_auto(render_page(pdf_path, page_no - 1, scale=scale), lang,
                   try_all=try_all)
    res["page"] = page_no
    res["scale"] = scale
    return res


def ocr_zone(pdf_path: str, page_no: int, top: float, bottom: float,
             rotation: int = 0, scale: float = 3.0, lang: str = "ru") -> dict:
    """
    Распознать горизонтальную полосу страницы (доли высоты, 0..1).

    Зоны нужны потому, что полностраничный OCR смешивает таблицу с подписями
    на чертеже: типоразмер «108х6-89х6» на листе акта читается как
    «пн:108х649х6б9», а тот же кроп нижней полосы — верно.
    """
    from repair.pdf_preview import render_page

    arr = _rot(render_page(pdf_path, page_no - 1, scale=scale), rotation)
    h = arr.shape[0]
    crop = arr[int(h * top):int(h * bottom), :, :]
    lines = ocr_lines(crop, lang)
    return {"page": page_no, "rotation": rotation, "scale": scale,
            "полоса": [top, bottom], "size": [crop.shape[1], crop.shape[0]],
            "score": score_lines(lines), "lines": lines}


def ocr_zone_multi(pdf_path: str, page_no: int, top: float, bottom: float,
                   rotation: int = 0, lang: str = "ru",
                   scales: tuple = (2.2, 3.0, 3.5)) -> dict:
    """
    Распознать зону в НЕСКОЛЬКО проходов с разными масштабами рендера.

    Windows OCR нестабилен по масштабу: на одном и том же скане ячейка
    «108х6-89х6» читается верно при 1.5/2.2/3.0/3.5 и ломается при 2.0/2.5
    («пн:108х649х6б9»). Один проход, даже кропом, не даёт гарантии — поэтому
    критичные мелкие поля берутся согласованием нескольких проходов
    (`best_match`).

    Возвращает {"варианты": [{"scale", "text", "lines"}], "page", "полоса"}.
    """
    варианты = []
    for s in scales:
        res = ocr_zone(pdf_path, page_no, top, bottom, rotation=rotation,
                       scale=s, lang=lang)
        варианты.append({"scale": s, "lines": res["lines"],
                         "text": " ".join(ln["text"] for ln in res["lines"])})
    return {"page": page_no, "rotation": rotation, "полоса": [top, bottom],
            "варианты": варианты}


def best_match(variants: list[dict], pattern) -> str | None:
    """
    Наиболее частое совпадение регулярки по вариантам распознавания.

    Так из нескольких проходов выбирается значение, которое OCR прочитал
    одинаково чаще всего: одиночная ошибка масштаба остаётся в меньшинстве.
    """
    import collections

    rx = re.compile(pattern) if isinstance(pattern, str) else pattern
    counter: collections.Counter = collections.Counter()
    for v in variants:
        for m in rx.finditer(v.get("text", "")):
            counter[m.group(0)] += 1
    if not counter:
        return None
    return counter.most_common(1)[0][0]


def rows_from_lines(lines: list[dict]) -> list[str]:
    """
    Склеить ячейки таблицы в строки таблицы.

    Windows OCR отдаёт таблицу по ячейкам: в акте «Переход» и «108х6-89х6»
    приходят разными строками, и запись о замене выглядит как две половинки.
    Ячейки с близкими вертикальными центрами — одна строка таблицы; внутри
    строки они выстраиваются слева направо.
    """
    cells = []
    for ln in lines:
        words = ln.get("words") or []
        if not words:
            continue
        cells.append({
            "y": sum(w["y"] + w["h"] / 2 for w in words) / len(words),
            "h": max(w["h"] for w in words),
            "x": min(w["x"] for w in words),
            "t": ln["text"],
        })
    cells.sort(key=lambda c: (c["y"], c["x"]))

    rows, cur = [], []
    for c in cells:
        if cur and abs(c["y"] - cur[0]["y"]) > max(cur[0]["h"] * 0.7, 8):
            rows.append(cur)
            cur = []
        cur.append(c)
    if cur:
        rows.append(cur)
    return [" ".join(c["t"] for c in sorted(r, key=lambda c: c["x"]))
            for r in rows]


def plain_text(res: dict) -> str:
    """Результат ocr_* → текст построчно."""
    return "\n".join(ln["text"] for ln in res.get("lines", []))
