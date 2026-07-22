# -*- coding: utf-8 -*-
"""
Распознавание текста через Windows OCR (этап Э4, слой 2).

Используется встроенный движок Windows 10/11 (Windows.Media.Ocr) через
pywinrt — ничего тяжёлого в поставку не добавляется (пакеты winrt-* весят
единицы мегабайт; при сборке EXE добавить их в окружение).

Качество на печатном тексте актов хорошее; рукописные вставки (день в грифе,
подписи) распознаются ненадёжно — поэтому OCR только ПРЕДЗАПОЛНЯЕТ поля,
подтверждает сотрудник по превью (принцип «ручной ввод равноправен»).
"""

from __future__ import annotations


def available() -> bool:
    """Доступен ли Windows OCR с русским языком на этой машине."""
    try:
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.globalization import Language
        return (OcrEngine.try_create_from_language(Language("ru")) is not None
                or OcrEngine.try_create_from_user_profile_languages() is not None)
    except Exception:
        return False


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
    import io
    from PIL import Image

    buf = io.BytesIO()
    Image.fromarray(arr).save(buf, format="PNG")
    return ocr_png_bytes(buf.getvalue(), lang)
