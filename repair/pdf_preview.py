# -*- coding: utf-8 -*-
"""
Превью страниц PDF для формы реестра (этап Э4, слой 1 — без распознавания).

Сотрудник вписывает в реестр дату грифа УТВЕРЖДАЮ, дату «Сдал» и типоразмеры,
глядя на страницы документа прямо в программе — не мечась между окнами.

Рендер — pypdfium2 (лёгкая зависимость ~3 МБ; при сборке EXE добавить в
зависимости). Документ кэшируется, чтобы листание страниц было мгновенным.
"""

from __future__ import annotations

# кэш последнего открытого документа: {path: (mtime, pdf)}
_cache: dict = {}


def _open(path: str):
    import os
    import pypdfium2 as pdfium

    mtime = os.path.getmtime(path)
    hit = _cache.get(path)
    if hit and hit[0] == mtime:
        return hit[1]
    # держим только один документ (файлы бывают по 50+ МБ)
    for _, (_, old) in list(_cache.items()):
        try:
            old.close()
        except Exception:
            pass
    _cache.clear()
    pdf = pdfium.PdfDocument(path)
    _cache[path] = (mtime, pdf)
    return pdf


def page_count(path: str) -> int:
    return len(_open(path))


def render_page(path: str, index: int, scale: float = 2.0):
    """
    Страница PDF (0-базный индекс) → numpy-массив RGB.
    scale 2.0 ≈ 144 dpi — гриф и подписи читаются при зуме.
    """
    import numpy as np

    pdf = _open(path)
    page = pdf[index]
    try:
        bmp = page.render(scale=scale)
        return np.asarray(bmp.to_pil().convert("RGB"))
    finally:
        page.close()


def close_all() -> None:
    for _, (_, pdf) in list(_cache.items()):
        try:
            pdf.close()
        except Exception:
            pass
    _cache.clear()
