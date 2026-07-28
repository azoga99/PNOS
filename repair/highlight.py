# -*- coding: utf-8 -*-
"""
Поиск красного выделения «заменяемый участок» на схеме.

Инженер ОТН отмечает дефектный участок жирной красной дугой — на схеме внутри
акта отбраковки, а иногда и прямо на исходнике коррозионной карты. Это самая
прямая подсказка, какое сечение подлежит замене: остальная красная разметка
(цифры, подписи, рамка) нарисована тонкими штрихами, а выделение — широкой
полосой.

Отсюда способ: взять красную маску и «съесть» её края на радиус, заметно
больший толщины обычного штриха. Тонкое исчезает, широкая полоса остаётся.

Радиус считается от размера изображения, а не задаётся числом: одна и та же
схема приходит и страницей PDF в разном масштабе, и PNG-картой на 4200 px.

Что с найденным выделением делают дальше, зависит от источника. На исходнике
ОТН из «Первички» координаты переносятся на рабочую карту один в один — их и
используют для поиска сечения. Схема внутри акта — другой чертёж; связать её с
картой можно было бы через подписанные рядом замеры, но проверка на п.2763
показала, что OCR читает их неверно (6.9 → «9.6», настоящее 2.6 пропущено) и
уводит на чужое сечение. Поэтому схему акта программа только показывает.

Слой чистый: только numpy, ни Qt, ни OCR, ни PDF.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np

from repair.map_editor import BBox, red_mask

#: Делитель для радиуса эрозии: k ≈ min(высота, ширина) / 1200.
#: Обычный штрих разметки 2–4 px, полоса выделения — 10–20 px.
EROSION_DIVISOR = 1200

#: Радиус меньше двух брать нельзя: окно 3×3 переживает штрих в 3 px, и обычная
#: разметка схемы начинает считаться выделением.
MIN_EROSION = 2

#: Меньше этого числа выживших пикселей — считаем шумом, а не выделением.
MIN_MARKER_PX = 40

#: Выделение — это полоса вдоль трубы, а не пятно: длинная сторона должна быть
#: заметно больше радиуса эрозии. Отсекает толстые точки и углы рамки — на
#: исходнике п.2800 такое пятно (12×6 px) иначе принималось за выделение.
MIN_MARKER_EXTENT = 8


@dataclass
class Marker:
    """Найденное выделение: рамка, центр и сколько пикселей выжило эрозию."""
    bbox: BBox
    center: tuple[int, int]
    pixels: int

    @property
    def relative_center(self) -> tuple[float, float]:
        """Центр в долях от размера изображения — для переноса между схемами."""
        return self._rel

    _rel: tuple[float, float] = (0.0, 0.0)


def erosion_radius(shape) -> int:
    """Радиус эрозии для изображения такого размера."""
    h, w = shape[:2]
    return max(MIN_EROSION, round(min(h, w) / EROSION_DIVISOR))


def thick_mask(mask: np.ndarray, k: int) -> np.ndarray:
    """
    Оставить только пиксели, вокруг которых маска сплошная на радиус k.

    Обычная эрозия квадратным окном; scipy ради неё тянуть незачем — сдвиги
    numpy на изображении карты отрабатывают за доли секунды.
    """
    out = mask.copy()
    for dy in range(-k, k + 1):
        for dx in range(-k, k + 1):
            if dy == 0 and dx == 0:
                continue
            out &= np.roll(np.roll(mask, dy, axis=0), dx, axis=1)
    return out


def _components(ys: np.ndarray, xs: np.ndarray, gap: int = 8) -> list[list[int]]:
    """
    Группы связных точек (соседи — в пределах `gap` пикселей), крупные первыми.

    Точек после эрозии остаются сотни, поэтому обходим их напрямую, без
    разметки всего изображения.
    """
    точки = list(zip(ys.tolist(), xs.tolist()))
    не_посещены = set(range(len(точки)))
    группы: list[list[int]] = []
    while не_посещены:
        старт = не_посещены.pop()
        группа = [старт]
        стек = [старт]
        while стек:
            i = стек.pop()
            yi, xi = точки[i]
            рядом = [j for j in не_посещены
                     if abs(точки[j][0] - yi) <= gap and abs(точки[j][1] - xi) <= gap]
            for j in рядом:
                не_посещены.discard(j)
                группа.append(j)
                стек.append(j)
        группы.append(группа)
    группы.sort(key=len, reverse=True)
    return группы


def find_markers(arr: np.ndarray, min_px: int = MIN_MARKER_PX) -> list[Marker]:
    """
    Все выделения на изображении схемы, крупные первыми.

    Их бывает несколько: акт 222-1-91 (п.2688) требует заменить 4 погонных
    метра трубы, и на схеме отмечены три разных места. Брать только самое
    крупное значило бы потерять остальные.
    """
    if arr is None or getattr(arr, "ndim", 0) != 3:
        return []
    k = erosion_radius(arr.shape)
    толстое = thick_mask(red_mask(arr), k)
    ys, xs = np.nonzero(толстое)
    if len(xs) < min_px:
        return []

    h, w = arr.shape[:2]
    out: list[Marker] = []
    for группа in _components(ys, xs):
        if len(группа) < min_px:
            break                       # дальше только мельче
        gy, gx = ys[группа], xs[группа]
        bbox = BBox(int(gx.min()), int(gy.min()),
                    int(gx.max()) + 1, int(gy.max()) + 1)
        if max(bbox.width(), bbox.height()) < MIN_MARKER_EXTENT * k:
            continue                    # толстая точка, а не полоса вдоль трубы
        cx, cy = int(gx.mean()), int(gy.mean())
        m = Marker(bbox=bbox, center=(cx, cy), pixels=len(группа))
        m._rel = (cx / w, cy / h)
        out.append(m)
    return out


def find_marker(arr: np.ndarray, min_px: int = MIN_MARKER_PX) -> Marker | None:
    """Самое крупное выделение или None — короткая форма `find_markers`."""
    найдено = find_markers(arr, min_px)
    return найдено[0] if найдено else None


def union_bbox(markers) -> BBox | None:
    """Рамка, охватывающая все выделения — чтобы показать их разом."""
    if not markers:
        return None
    return BBox(min(m.bbox.x1 for m in markers), min(m.bbox.y1 for m in markers),
                max(m.bbox.x2 for m in markers), max(m.bbox.y2 for m in markers))


def labels_near(labels, point: tuple[int, int], radius: int) -> list:
    """
    Подписи сечений в круге вокруг точки, ближние первыми.

    Радиус нужен приличный: подписи стоят не на трубе, а сбоку на выносках, и
    ближайшая к выделению подпись не обязательно относится к нему. Поэтому
    отдаём список кандидатов, а не одну подпись.
    """
    import math

    рядом = []
    for l in labels or []:
        d = math.hypot(l.point[0] - point[0], l.point[1] - point[1])
        if d <= radius:
            рядом.append((d, l))
    рядом.sort(key=lambda p: p[0])
    return [l for _, l in рядом]
