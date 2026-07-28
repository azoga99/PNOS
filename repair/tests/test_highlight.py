# -*- coding: utf-8 -*-
"""
Поиск красного выделения «заменяемый участок» (repair.highlight).

Синтетика: белый лист, тонкие красные штрихи (обычная разметка схемы) и одна
жирная полоса — то самое выделение инженера ОТН.
"""

import numpy as np
import pytest

from repair.highlight import (find_marker, erosion_radius, thick_mask,
                              labels_near, EROSION_DIVISOR, MIN_EROSION)
from repair.map_editor import BBox, red_mask
from repair.map_labels import MapLabel

КРАСНЫЙ = (200, 30, 40)
ЧЁРНЫЙ = (20, 20, 20)


def _лист(w=1400, h=1200):
    return np.full((h, w, 3), 255, dtype=np.uint8)


def _штрих(arr, x, y, длина=120, толщина=3):
    arr[y:y + толщина, x:x + длина] = КРАСНЫЙ


def _полоса(arr, x, y, длина=90, толщина=16):
    arr[y:y + толщина, x:x + длина] = КРАСНЫЙ


# ─── Радиус эрозии ────────────────────────────────────────────────────

def test_radius_scales_with_image():
    """Одна схема приходит в разных масштабах — радиус должен ехать за ними."""
    assert erosion_radius((2970, 4200, 3)) == round(2970 / EROSION_DIVISOR)
    assert erosion_radius((6000, 6000, 3)) == 5


def test_radius_never_below_minimum():
    """
    Окно 3×3 переживает штрих в 3 px — тогда обычная разметка схемы начинает
    считаться выделением (на этом сломался первый вариант).
    """
    assert erosion_radius((1685, 1191, 3)) == MIN_EROSION
    assert erosion_radius((100, 100, 3)) == MIN_EROSION


# ─── Отсев тонкого ────────────────────────────────────────────────────

def test_thin_strokes_are_not_a_marker():
    """Цифры и выноски схемы нарисованы тонко — выделением их считать нельзя."""
    arr = _лист()
    for i in range(12):
        _штрих(arr, 100 + i * 90, 200 + i * 70)
    assert find_marker(arr) is None


def test_thick_band_is_found():
    arr = _лист()
    _штрих(arr, 100, 200)
    _полоса(arr, 600, 700)
    m = find_marker(arr)
    assert m is not None
    # центр полосы: x 600…690, y 700…716
    assert 600 <= m.center[0] <= 690
    assert 700 <= m.center[1] <= 716


def test_marker_ignores_black_lines():
    """Трубопровод нарисован чёрным и толстым — он не выделение."""
    arr = _лист()
    arr[500:520, 100:900] = ЧЁРНЫЙ
    assert find_marker(arr) is None


def test_largest_band_wins():
    """Рамка чертежа местами тоже толстая — берём наибольшее пятно."""
    arr = _лист()
    _полоса(arr, 50, 50, длина=25, толщина=14)      # мелкое пятно
    _полоса(arr, 600, 700, длина=200, толщина=20)   # настоящее выделение
    m = find_marker(arr)
    assert m is not None
    assert m.center[0] > 500


def test_min_px_guard():
    """Крошечное пятно — шум, а не выделение."""
    arr = _лист()
    _полоса(arr, 600, 700, длина=12, толщина=12)
    assert find_marker(arr) is None


def test_relative_center_for_transfer():
    arr = _лист(1000, 500)
    _полоса(arr, 400, 200, длина=200, толщина=20)
    m = find_marker(arr)
    rx, ry = m.relative_center
    assert 0.45 < rx < 0.65
    assert 0.35 < ry < 0.50


def test_thick_mask_erodes():
    m = np.zeros((40, 40), dtype=bool)
    m[10:30, 10:30] = True
    e = thick_mask(m, 2)
    assert e[20, 20]
    assert not e[10, 10]        # угол съеден


def test_no_marker_on_blank():
    assert find_marker(_лист()) is None
    assert find_marker(None) is None


# ─── Подписи вокруг выделения ─────────────────────────────────────────

def _подпись(number, x, y):
    return MapLabel(number=number, point=(x, y),
                    bbox=BBox(x - 10, y - 8, x + 10, y + 8))


def test_labels_near_sorted_by_distance():
    labels = [_подпись(28, 2709, 2378), _подпись(29, 2446, 2345),
              _подпись(31, 2433, 2209), _подпись(99, 100, 100)]
    рядом = labels_near(labels, (2625, 2307), 260)
    assert [l.number for l in рядом] == [28, 29, 31]


def test_labels_near_empty_when_far():
    assert labels_near([_подпись(1, 0, 0)], (5000, 5000), 100) == []
    assert labels_near(None, (0, 0), 100) == []
