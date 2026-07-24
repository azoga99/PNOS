# -*- coding: utf-8 -*-
"""Тесты правки карты вводом значения (Arial + цвет из подписи)."""

import numpy as np
import pytest

from repair.map_editor import render_value, red_mask, ink_mask, _ink_color, BBox

ORANGE = (255, 144, 60)
DARKRED = (158, 40, 36)
BLACK = (20, 20, 20)


def _blank(w=400, h=140):
    return np.full((h, w, 3), 255, dtype=np.uint8)


def _draw_glyph(arr, x, y, w=8, h=16, color=DARKRED):
    arr[y:y + h, x:x + 2] = color
    arr[y:y + h, x + w - 2:x + w] = color
    arr[y:y + 2, x:x + w] = color
    arr[y + h - 2:y + h, x:x + w] = color


def _draw_label(arr, x, y, n_num=2, n_val=5, color=DARKRED):
    gx = x
    for _ in range(n_num):
        _draw_glyph(arr, gx, y, color=color)
        gx += 11
    gx += 12
    for _ in range(n_val):
        _draw_glyph(arr, gx, y, color=color)
        gx += 11


def test_render_writes_text_and_keeps_line():
    arr = _blank()
    _draw_label(arr, 40, 50)
    arr[57:59, 20:300] = BLACK          # линия сквозь подпись
    rep = render_value(arr, (52, 58), "(6.0)")

    assert rep["text"] == "(6.0)"
    assert rep["erased_px"] > 0
    # что-то нарисовано в зоне значения
    wb = rep["written_box"]
    assert ink_mask(arr[wb.y1:wb.y2, wb.x1:wb.x2]).any()
    # чёрная линия цела
    assert (arr[57, 20:40] == BLACK).all()


def test_color_taken_from_label_orange():
    """Оранжевая подпись остаётся оранжевой (не перекрашивается в красный)."""
    arr = _blank()
    _draw_label(arr, 40, 50, color=ORANGE)
    rep = render_value(arr, (52, 58), "(6.0)")
    r, g, b = rep["color"]
    assert r > 200 and 90 < g < 220 and b < 160     # оранжевый диапазон
    assert not (g < 80 and b < 80)                  # не тёмно-красный


def test_color_taken_from_label_dark_red():
    arr = _blank()
    _draw_label(arr, 40, 50, color=DARKRED)
    rep = render_value(arr, (52, 58), "(6.0)")
    r, g, b = rep["color"]
    assert r > 120 and g < 130 and b < 130


def test_font_size_scales_with_label():
    """Кегль подбирается под высоту подписи: крупнее подпись — крупнее шрифт."""
    small = _blank(400, 140)
    _draw_label(small, 40, 50, n_val=5)
    r_small = render_value(small, (52, 58), "(6.0)")

    big = _blank(700, 220)
    for i in range(2):
        _draw_glyph(big, 40 + i * 22, 60, w=16, h=32)
    gx = 40 + 2 * 22 + 24
    for i in range(5):
        _draw_glyph(big, gx + i * 22, 60, w=16, h=32)
    r_big = render_value(big, (60, 76), "(6.0)")

    assert r_big["font_size"] > r_small["font_size"]


def test_no_label_raises():
    arr = _blank()
    with pytest.raises(ValueError):
        render_value(arr, (10, 10), "(6.0)")


def test_ink_color_ignores_white_background():
    arr = _blank()
    arr[10:20, 10:20] = ORANGE
    c = _ink_color(arr, BBox(0, 0, 40, 40))
    assert c[0] > 200 and c[2] < 160
