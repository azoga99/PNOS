# -*- coding: utf-8 -*-
"""
Тесты ядра автоправки карты (Э3) на синтетическом изображении:
красная подпись «47 (3.4)», чёрная линия через подпись, донор «16 (6.0)».
"""

import numpy as np
import pytest

from repair.map_editor import (
    red_mask, label_bbox_at, find_value_span, erase_red, copy_red,
    replace_value, BBox,
)

RED = (158, 40, 36)
BLACK = (20, 20, 20)
WHITE = (255, 255, 255)


def _blank(w=400, h=120):
    return np.full((h, w, 3), 255, dtype=np.uint8)


def _draw_glyph(arr, x, y, w=8, h=16, color=RED):
    """Примитивный «глиф» — прямоугольная рамка пикселей."""
    arr[y:y + h, x:x + 2] = color
    arr[y:y + h, x + w - 2:x + w] = color
    arr[y:y + 2, x:x + w] = color
    arr[y + h - 2:y + h, x:x + w] = color


def _draw_label(arr, x, y, n_num=2, n_val=5):
    """Подпись «NN (V.V)»: n_num глифов, широкий пробел, n_val глифов."""
    gx = x
    for _ in range(n_num):
        _draw_glyph(arr, gx, y)
        gx += 11          # межглифовый зазор 3px
    gx += 12              # большой пробел перед «(»
    val_x1 = gx
    for _ in range(n_val):
        _draw_glyph(arr, gx, y)
        gx += 11
    return val_x1, gx - 3  # рамка значения по x


def test_red_mask_separates_colors():
    arr = _blank()
    arr[10, 10] = RED
    arr[10, 20] = BLACK
    m = red_mask(arr)
    assert m[10, 10] and not m[10, 20]


def test_label_bbox_and_value_span():
    arr = _blank()
    vx1, vx2 = _draw_label(arr, 50, 40)
    bb = label_bbox_at(arr, 60, 48)
    assert bb is not None
    assert bb.x1 == 50 and bb.y1 == 40

    span = find_value_span(arr, bb)
    assert span is not None
    assert span.x1 == vx1              # значение начинается после большого пробела
    assert span.x2 == vx2              # правый край последнего глифа


def test_erase_red_keeps_black_line():
    arr = _blank()
    _draw_label(arr, 50, 40)
    arr[47:49, 40:200] = BLACK         # линия сквозь подпись
    bb = label_bbox_at(arr, 60, 48)
    span = find_value_span(arr, bb)
    erased = erase_red(arr, span)
    assert erased > 0
    assert not red_mask(arr[span.y1:span.y2, span.x1:span.x2]).any()
    # линия жива
    assert (arr[47, span.x1:span.x2] == BLACK).all(axis=-1).any()


def test_copy_red_transfers_glyphs():
    arr = _blank()
    _draw_glyph(arr, 300, 80)
    src = BBox(300, 80, 308, 96)
    n = copy_red(arr, src, 100, 60)
    assert n > 0
    assert red_mask(arr[44:60, 100:108]).sum() == n


def test_replace_value_end_to_end():
    arr = _blank(500, 200)
    _draw_label(arr, 50, 40, n_val=5)          # правимая подпись
    _draw_label(arr, 50, 140, n_val=4)         # донор (другое значение — 4 глифа)
    arr[47:49, 30:250] = BLACK                 # линия через первую подпись

    rep = replace_value(arr, (60, 48), (60, 148))
    assert rep["erased_px"] > 0 and rep["copied_px"] > 0

    # в зоне старого значения теперь глифы донора: ширина по донору (4 глифа)
    lv = rep["value_old"]
    dv = rep["donor_value"]
    new_zone = red_mask(arr[lv.y2 - dv.height():lv.y2, lv.x1:lv.x1 + dv.width()])
    assert new_zone.sum() == rep["copied_px"]
    # линия не пострадала
    assert (arr[47, lv.x1:lv.x2] == BLACK).all(axis=-1).any()
    # донор на месте (копировали, не вырезали)
    assert red_mask(arr[dv.y1:dv.y2, dv.x1:dv.x2]).any()


def test_replace_value_no_label_raises():
    arr = _blank()
    with pytest.raises(ValueError):
        replace_value(arr, (10, 10), (20, 20))
