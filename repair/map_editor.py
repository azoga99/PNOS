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

import numpy as np
from dataclasses import dataclass


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


def label_bbox_at(arr: np.ndarray, x: int, y: int,
                  reach_x: int = 22, reach_y: int = 8,
                  max_grow: int = 60) -> BBox | None:
    """
    Рамка красной подписи (одной текстовой строки) вокруг точки клика.
    Рамка растёт, пока в зоне досягаемости за её краями есть красные пиксели.
    Досягаемость раздельная: по горизонтали reach_x больше пробела
    «номер—значение» (~10–15 px), по вертикали reach_y меньше межстрочного
    расстояния — иначе рамка слипается с соседними подписями (реальный случай
    «16 (6.0)» на карте п.2775).
    """
    h, w = arr.shape[:2]
    m = red_mask(arr)

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


def erase_red(arr: np.ndarray, bbox: BBox) -> int:
    """Стереть красные пиксели в рамке (→ белый). Возвращает число пикселей."""
    zone = arr[bbox.y1:bbox.y2, bbox.x1:bbox.x2]
    m = red_mask(zone)
    zone[m] = (255, 255, 255)
    return int(m.sum())


def copy_red(arr: np.ndarray, src: BBox, dst_x: int, dst_y: int) -> int:
    """
    Скопировать красные пиксели из рамки-источника в позицию (левый-нижний
    угол назначения). Возвращает число перенесённых пикселей.
    """
    patch = arr[src.y1:src.y2, src.x1:src.x2].copy()
    m = red_mask(patch)
    h, w = patch.shape[:2]
    y1 = dst_y - h
    region = arr[y1:y1 + h, dst_x:dst_x + w]
    region[m] = patch[m]
    return int(m.sum())


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
