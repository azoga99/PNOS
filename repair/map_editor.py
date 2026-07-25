# -*- coding: utf-8 -*-
"""
Ядро автоправки коррозионной карты (этап Э3).

Карта — PNG, отрендеренный из векторного редактора: разметка (номера сечений
и значения замеров) красная, линии трубопровода чёрные, фон белый. Отсюда два
приёма, доказанные прототипом на п.2775:

  1) стирать только красные пиксели — чёрные линии под подписью не страдают;
  2) новое значение не рисовать шрифтом, а копировать красные глифы готовой
     подписи-донора с этой же карты — шрифт совпадает идеально.

Подпись сечения: «NN (V.V)». Номер и значение разделены наибольшим пробелом
внутри подписи — так выделяется часть «(V.V)» без распознавания текста.

Слой чистый (numpy + Pillow), Qt не знает; страница даст координаты кликов.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

import numpy as np


@dataclass
class BBox:
    x1: int
    y1: int
    x2: int  # исключительно
    y2: int

    def width(self) -> int:
        return self.x2 - self.x1

    def height(self) -> int:
        return self.y2 - self.y1

    def pad(self, p: int, w: int, h: int) -> "BBox":
        return BBox(max(0, self.x1 - p), max(0, self.y1 - p),
                    min(w, self.x2 + p), min(h, self.y2 + p))


def red_mask(arr: np.ndarray) -> np.ndarray:
    """Маска красной разметки (тёмно-красный чертёжный: R доминирует)."""
    r = arr[:, :, 0].astype(int)
    g = arr[:, :, 1].astype(int)
    b = arr[:, :, 2].astype(int)
    return (r > 90) & (r - g > 40) & (r - b > 40)


def _near_white(arr: np.ndarray) -> np.ndarray:
    return (arr[:, :, 0] > 205) & (arr[:, :, 1] > 205) & (arr[:, :, 2] > 205)


def _line_mask(arr: np.ndarray) -> np.ndarray:
    """Чёрная линия трубопровода: тёмный нейтральный пиксель (не цветной)."""
    r = arr[:, :, 0].astype(int)
    g = arr[:, :, 1].astype(int)
    b = arr[:, :, 2].astype(int)
    mx = np.maximum(np.maximum(r, g), b)
    mn = np.minimum(np.minimum(r, g), b)
    return (mx < 95) & (mx - mn < 45)


def markup_mask(arr: np.ndarray) -> np.ndarray:
    """
    Только цветная разметка любого цвета: красная, синяя, оранжевая.

    В отличие от `ink_mask` не захватывает серую кайму сглаживания чёрных
    линий. Для поиска подписей это принципиально: по маске со сглаживанием
    рамка подписи прирастает к трубопроводу и перестаёт быть похожей на
    текстовую строку.
    """
    r = arr[:, :, 0].astype(int)
    g = arr[:, :, 1].astype(int)
    b = arr[:, :, 2].astype(int)
    mx = np.maximum(np.maximum(r, g), b)
    mn = np.minimum(np.minimum(r, g), b)
    return (mx > 90) & (mx - mn > 45)


def ink_mask(arr: np.ndarray) -> np.ndarray:
    """
    Вся цветная разметка (красные глифы + их сглаженная кайма + цветовой шум
    PNG), но НЕ фон и НЕ чёрные линии. Стирание по этой маске убирает старое
    значение начисто, без синеватых «призраков» по краям.
    """
    return ~_near_white(arr) & ~_line_mask(arr)


def label_bbox_at(arr: np.ndarray, x: int, y: int,
                  reach_x: int = 22, reach_y: int = 8,
                  max_grow: int = 60,
                  mask: np.ndarray | None = None) -> BBox | None:
    """
    Рамка подписи (одной текстовой строки) вокруг точки клика.
    Рамка растёт, пока в зоне досягаемости за её краями есть пиксели разметки.
    Досягаемость раздельная: по горизонтали reach_x больше пробела
    «номер—значение» (~10–15 px), по вертикали reach_y меньше межстрочного
    расстояния — иначе рамка слипается с соседними подписями (реальный случай
    «16 (6.0)» на карте п.2775).

    По умолчанию берётся красная разметка. Готовую маску можно передать
    снаружи: на картах есть подписи не только красные, но и синие с
    оранжевыми — для их поиска нужна маска всей цветной разметки (`ink_mask`).
    """
    h, w = arr.shape[:2]
    m = red_mask(arr) if mask is None else mask

    # стартовая привязка: ближайший красный пиксель в окне вокруг клика
    win = 24
    sub = m[max(0, y - win):min(h, y + win), max(0, x - win):min(w, x + win)]
    ys, xs = np.nonzero(sub)
    if len(xs) == 0:
        return None
    x1 = max(0, x - win) + int(xs.min())
    y1 = max(0, y - win) + int(ys.min())
    x2 = max(0, x - win) + int(xs.max()) + 1
    y2 = max(0, y - win) + int(ys.max()) + 1

    for _ in range(max_grow):
        grown = False
        lx = max(0, x1 - reach_x)
        if lx < x1 and m[y1:y2, lx:x1].any():
            _, xs_ = np.nonzero(m[y1:y2, lx:x1])
            x1 = lx + int(xs_.min())
            grown = True
        rx = min(w, x2 + reach_x)
        if rx > x2 and m[y1:y2, x2:rx].any():
            _, xs_ = np.nonzero(m[y1:y2, x2:rx])
            x2 = x2 + int(xs_.max()) + 1
            grown = True
        ty = max(0, y1 - reach_y)
        if ty < y1 and m[ty:y1, x1:x2].any():
            ys_, _ = np.nonzero(m[ty:y1, x1:x2])
            y1 = ty + int(ys_.min())
            grown = True
        by = min(h, y2 + reach_y)
        if by > y2 and m[y2:by, x1:x2].any():
            ys_, _ = np.nonzero(m[y2:by, x1:x2])
            y2 = y2 + int(ys_.max()) + 1
            grown = True
        if not grown:
            break
    return BBox(x1, y1, x2, y2)


def _glyph_columns(m: np.ndarray) -> list[tuple[int, int]]:
    """Группы непустых столбцов маски → интервалы глифов [(c1, c2), …]."""
    cols = m.any(axis=0)
    groups = []
    start = None
    for i, v in enumerate(cols):
        if v and start is None:
            start = i
        elif not v and start is not None:
            groups.append((start, i))
            start = None
    if start is not None:
        groups.append((start, len(cols)))
    return groups


def find_value_span(arr: np.ndarray, bbox: BBox) -> BBox | None:
    """
    Часть «(V.V)» внутри рамки подписи «NN (V.V)»: значение начинается после
    НАИБОЛЬШЕГО пробела между глифами (пробел «номер—скобка» шире межглифовых).
    """
    m = red_mask(arr[bbox.y1:bbox.y2, bbox.x1:bbox.x2])
    groups = _glyph_columns(m)
    if len(groups) < 2:
        return None
    gaps = [(groups[i + 1][0] - groups[i][1], i) for i in range(len(groups) - 1)]
    max_gap, idx = max(gaps)
    if max_gap <= 1:
        return None
    val_groups = groups[idx + 1:]
    c1, c2 = val_groups[0][0], val_groups[-1][1]
    ys, _ = np.nonzero(m[:, c1:c2])
    if len(ys) == 0:
        return None
    return BBox(bbox.x1 + c1, bbox.y1 + int(ys.min()),
                bbox.x1 + c2, bbox.y1 + int(ys.max()) + 1)


def erase_red(arr: np.ndarray, bbox: BBox, pad: int = 3) -> int:
    """
    Стереть значение начисто: всю цветную разметку (красное + сглаженная кайма
    + цветовой шум) в рамке с небольшим запасом → белый; чёрные линии
    сохраняются. Возвращает число стёртых пикселей.
    """
    h, w = arr.shape[:2]
    bb = bbox.pad(pad, w, h)
    zone = arr[bb.y1:bb.y2, bb.x1:bb.x2]
    m = ink_mask(zone)
    zone[m] = (255, 255, 255)
    return int(m.sum())


def copy_red(arr: np.ndarray, src: BBox, dst_x: int, dst_y: int) -> int:
    """
    Скопировать глифы значения из рамки-источника в позицию (левый-нижний
    угол назначения), вместе со сглаживанием — чтобы вставленное значение
    выглядело «родным». Чёрные линии донора не переносятся. Возвращает
    число перенесённых пикселей.
    """
    patch = arr[src.y1:src.y2, src.x1:src.x2].copy()
    m = ink_mask(patch)
    h, w = patch.shape[:2]
    y1 = dst_y - h
    region = arr[y1:y1 + h, dst_x:dst_x + w]
    region[m] = patch[m]
    return int(m.sum())


def _ink_color(arr: np.ndarray, bbox: BBox) -> tuple[int, int, int]:
    """
    Цвет разметки подписи: медиана «самых цветных» её пикселей. Берётся из
    самой подписи, поэтому оранжевые сечения остаются оранжевыми, а красные —
    красными (на схемах заменяемые сечения часто выделены другим цветом).
    """
    zone = arr[bbox.y1:bbox.y2, bbox.x1:bbox.x2].reshape(-1, 3).astype(int)
    if not len(zone):
        return (158, 40, 36)
    # «чернильные» пиксели: заметно цветные либо тёмные (без фона)
    saturation = zone.max(axis=1) - zone.min(axis=1)
    ink = zone[(saturation > 40) | (zone.max(axis=1) < 200)]
    if not len(ink):
        ink = zone
    # медиана по чернилам вместе со сглаживанием: тонкие штрихи исходного
    # чертёжного шрифта визуально светлее сплошной заливки Arial
    return tuple(int(v) for v in np.median(ink, axis=0))


def _font_path(italic: bool = True) -> str | None:
    names = ["ariali.ttf", "arial.ttf"] if italic else ["arial.ttf", "ariali.ttf"]
    for n in names:
        p = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", n)
        if os.path.exists(p):
            return p
    return None


def _font_for_text(text: str, target_h: int, italic: bool = True):
    """
    Arial такого кегля, чтобы отрисованный ТЕКСТ ЦЕЛИКОМ (со скобками) был по
    высоте как исходное значение. Сравнение like-for-like: скобки выше цифр,
    поэтому подбирать по высоте цифр нельзя — получится крупнее оригинала.
    """
    from PIL import ImageFont

    path = _font_path(italic)
    if path is None:
        return ImageFont.load_default(), 0

    size = max(6, int(target_h * 0.9))
    best, best_d = None, 10 ** 9
    for _ in range(40):
        f = ImageFont.truetype(path, size)
        box = f.getbbox(text)
        h = box[3] - box[1]
        d = abs(h - target_h)
        if d < best_d:
            best, best_d, best_size = f, d, size
        if h == target_h or size <= 6:
            return f, size
        size += 1 if h < target_h else -1
        size = max(6, size)
    return best, best_size


def tighten_to_ink(arr: np.ndarray, box: BBox, pad: int = 2) -> BBox | None:
    """
    Ужать рамку до фактических цветных глифов внутри неё.

    Нужно, когда рамка пришла извне (например, от OCR): она бывает свободнее
    глифов, а для стирания и подбора размера шрифта требуются точные границы.
    """
    h, w = arr.shape[:2]
    bb = box.pad(pad, w, h)
    m = markup_mask(arr[bb.y1:bb.y2, bb.x1:bb.x2])
    ys, xs = np.nonzero(m)
    if len(xs) == 0:
        return None
    return BBox(bb.x1 + int(xs.min()), bb.y1 + int(ys.min()),
                bb.x1 + int(xs.max()) + 1, bb.y1 + int(ys.max()) + 1)


def render_value(arr: np.ndarray, label_xy: tuple[int, int], text: str,
                 italic: bool = True) -> dict:
    """
    Правка «клик + ввод»: заменить значение подписи на введённый текст.
    Цвет берётся из самой подписи, высота цифр — как у исходного значения,
    шрифт Arial (курсив, как на схемах). Возвращает отчёт с рамками.
    """
    lb = label_bbox_at(arr, *label_xy)
    if lb is None:
        raise ValueError("рядом с кликом не найдена подпись")
    lv = find_value_span(arr, lb)
    if lv is None:
        raise ValueError("в подписи не выделяется значение «(…)»")
    return render_value_in(arr, lv, text, italic=italic, label=lb)


def render_value_in(arr: np.ndarray, value_box: BBox, text: str,
                    italic: bool = True, label: BBox | None = None) -> dict:
    """
    Заменить значение в ЗАРАНЕЕ ИЗВЕСТНОЙ рамке.

    Отдельно от `render_value` потому, что при автоматической правке рамка
    берётся из распознанной подписи, а не выращивается от клика. Наращивание
    рядом с подписью цепляет красный крестик замера: рамка «раздувается», и
    значение пишется чёрным, огромным и не на своём месте (наступил на это
    на карте п.2763, сечение 115).
    """
    from PIL import Image, ImageDraw

    lv = tighten_to_ink(arr, value_box) or value_box
    lb = label if label is not None else lv
    color = _ink_color(arr, lv)
    font, size = _font_for_text(text, lv.height(), italic)

    erased = erase_red(arr, lv)

    # текст рисуем на прозрачном слое и вставляем по левому-нижнему углу
    pad = 6
    tmp = Image.new("RGBA", (lv.width() * 3 + 2 * pad, lv.height() * 3 + 2 * pad),
                    (255, 255, 255, 0))
    d = ImageDraw.Draw(tmp)
    d.text((pad, pad), text, font=font, fill=tuple(color) + (255,))
    bbox_t = tmp.getbbox()
    if bbox_t is None:
        raise ValueError("не удалось отрисовать текст")
    glyphs = tmp.crop(bbox_t)

    gh, gw = glyphs.size[1], glyphs.size[0]
    x0 = lv.x1
    y0 = max(0, lv.y2 - gh)
    h, w = arr.shape[:2]
    gw = min(gw, w - x0)
    gh = min(gh, h - y0)
    patch = np.array(glyphs.convert("RGBA"))[:gh, :gw]
    alpha = patch[:, :, 3:4].astype(float) / 255.0
    region = arr[y0:y0 + gh, x0:x0 + gw].astype(float)
    arr[y0:y0 + gh, x0:x0 + gw] = (patch[:, :, :3] * alpha
                                   + region * (1 - alpha)).astype(np.uint8)

    return {"label": lb, "value_old": lv, "erased_px": erased,
            "color": color, "font_size": size, "text": text,
            "written_box": BBox(x0, y0, x0 + gw, y0 + gh)}


def replace_value(arr: np.ndarray, label_xy: tuple[int, int],
                  donor_xy: tuple[int, int]) -> dict:
    """
    Высокоуровневая операция «два клика»: подпись сечения + подпись-донор
    с нужным значением. Значение подписи стирается (только красное), на его
    место копируются глифы значения донора (выравнивание по левому нижнему
    углу стёртого). Возвращает отчёт с рамками (для показа до/после).
    Исключение ValueError — подпись/значение не найдены.
    """
    lb = label_bbox_at(arr, *label_xy)
    if lb is None:
        raise ValueError("рядом с кликом не найдена красная подпись")
    lv = find_value_span(arr, lb)
    if lv is None:
        raise ValueError("в подписи не выделяется значение «(…)»")

    db = label_bbox_at(arr, *donor_xy)
    if db is None:
        raise ValueError("рядом со вторым кликом не найдена подпись-донор")
    dv = find_value_span(arr, db)
    if dv is None:
        raise ValueError("в подписи-доноре не выделяется значение «(…)»")

    erased = erase_red(arr, lv)
    copied = copy_red(arr, dv, lv.x1, lv.y2)
    return {"label": lb, "value_old": lv, "donor_value": dv,
            "erased_px": erased, "copied_px": copied}
