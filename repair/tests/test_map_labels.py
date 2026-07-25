# -*- coding: utf-8 -*-
"""
Тесты поиска подписей сечений на коррозионной карте.

Проверка идёт на реальной карте готового пункта — на синтетике такой разбор
смысла не имеет. Если карт на машине нет или недоступен Windows OCR, тесты
пропускаются.
"""

import os

import numpy as np
import pytest

from repair.map_labels import (MapLabel, find_labels, labels_by_number,
                               MAX_OCR_DIM, MAX_SECTION)
from repair.map_editor import BBox

_CANDIDATES = [
    r"C:\Users\volod\Desktop\Готовые — копия\п.2763(ШАР)\Первичка\2763.png",
    r"C:\Users\volod\Desktop\Готовые\п.2763(ШАР)\Первичка\2763.png",
]
REAL_MAP = next((p for p in _CANDIDATES if os.path.isfile(p)), _CANDIDATES[0])

# Подписи, снятые с карты п.2763 глазами.
ГЛАЗАМИ = [59, 60, 61, 102, 103, 104, 105, 106, 110, 113, 114, 115, 116, 117]


def _map():
    from repair.ocr import available
    if not available():
        pytest.skip("Windows OCR на этой машине недоступен")
    if not os.path.isfile(REAL_MAP):
        pytest.skip("реальных карт на этой машине нет")
    from PIL import Image
    return np.array(Image.open(REAL_MAP).convert("RGB"))


@pytest.fixture(scope="module")
def labels():
    return find_labels(_map())


def test_finds_most_labels(labels):
    """На карте около 127 сечений — находиться должно большинство."""
    assert len(labels) >= 90
    номера = {l.number for l in labels}
    assert len(номера) >= 90
    assert max(номера) <= MAX_SECTION


def test_finds_labels_checked_by_eye(labels):
    номера = {l.number for l in labels}
    не_найдено = [n for n in ГЛАЗАМИ if n not in номера]
    # Часть подписей OCR не берёт — их сотрудник правит кликом, как раньше.
    assert len(не_найдено) <= 2, f"потеряно слишком много: {не_найдено}"


def test_no_numbers_from_title_block(labels):
    """
    В штампе и таблице элементов есть числа («219», «20», «2.58»), но подписью
    считается только «число + значение в скобках» — мусор попадать не должен.
    """
    номера = [l.number for l in labels]
    assert 219 not in номера
    assert 2970 not in номера


def test_label_points_inside_bbox(labels):
    for l in labels[:20]:
        x, y = l.point
        assert l.bbox.x1 <= x <= l.bbox.x2
        assert l.bbox.y1 <= y <= l.bbox.y2


def test_labels_are_spread_over_map(labels):
    """Подписи разбросаны по всей карте, а не сгрудились в одном углу."""
    ys = [l.point[1] for l in labels]
    assert max(ys) - min(ys) > 1000


def test_labels_by_number_keeps_all_duplicates():
    """Один номер может встретиться несколько раз — все места сохраняются."""
    подписи = [
        MapLabel(number=7, point=(10, 10), bbox=BBox(0, 0, 20, 20)),
        MapLabel(number=7, point=(10, 900), bbox=BBox(0, 890, 20, 910)),
        MapLabel(number=8, point=(50, 50), bbox=BBox(40, 40, 60, 60)),
    ]
    по_номеру = labels_by_number(подписи)
    assert [l.point[1] for l in по_номеру[7]] == [10, 900]
    assert len(по_номеру[8]) == 1


def test_oversized_scales_are_skipped():
    """
    Windows OCR не принимает изображения крупнее MaxImageDimension. Масштабы,
    выводящие карту за предел, должны отбрасываться до вызова OCR — иначе
    падение WinError (карта 2970×4200 при масштабе 2.5).
    """
    arr = np.full((500, 500, 3), 255, dtype=np.uint8)
    сообщения = []
    итог = find_labels(arr, scales=(30.0,), log=сообщения.append)
    assert итог == []
    assert any(str(MAX_OCR_DIM) in m for m in сообщения)
