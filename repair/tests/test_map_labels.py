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
    r"C:\Users\volod\Desktop\Сданы на ПНОС — копия"
    r"\п.2763 (Всем отправлено)\Первичка\2763.png",
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
    """
    На карте около 127 сечений — находиться должно большинство.

    Порог держится у достигнутого: до перехода на признак «номер нарисован
    цветной краской» здесь находилось 106 сечений, стало 110. Просядет ниже
    105 — значит распознавание сломали.
    """
    assert len(labels) >= 105
    номера = {l.number for l in labels}
    assert len(номера) >= 105
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


# ─── Признак подписи: цвет номера, а не читаемость значения ────────────

def test_value_like_accepts_cyrillic_misreads():
    """
    Значение в мелком курсиве OCR читает кириллицей — это не повод терять
    подпись: сам текст значения не используется, его вырезает геометрия.
    Примеры сняты с карт п.2688 и п.2775.
    """
    from repair.map_labels import _value_like

    for токен in ("(3.0)", "(ЗО", "(ЗЯ", "(ЗУ", "(за", "В.Ы", "ИЗ)", "(6.4]",
                  "(12.4)"):
        assert _value_like(токен), токен


def test_value_like_rejects_title_block_text():
    """
    «116 т-и» на карте п.2778 — это «Т-11 в Т-12» из основной надписи, набранное
    тем же бордовым. Проходило как подпись сечения №116, и правка ушла бы в штамп.
    """
    from repair.map_labels import _value_like

    for токен in ("т-и", "Т-12", "", "  ", "108", "12"):
        assert not _value_like(токен), токен


def test_value_like_rejects_glued_neighbour():
    """
    «(4.7) 78 (6.1)» на карте п.2775: OCR потерял скобки у чужого значения и
    выдал номер «47» со «значением» «786.1» — четыре цифры, столько не бывает.
    """
    from repair.map_labels import _value_like

    assert not _value_like("786.1")


def test_ink_color_is_mode_not_average():
    """
    Цвет подписи — самый частый цвет чернил.

    Карта квантована, и вокруг чёрных глифов рассыпан цветной дизер: у надписи
    «10(1)» на карте п.2704 чёрных пикселей 207, а цветных в ореоле 467.
    Проверка «цветных больше чёрных» на этом ломалась.
    """
    from repair.map_labels import ink_color, is_markup_number

    arr = np.full((40, 40, 3), 255, np.uint8)
    arr[10:30, 10:30] = (0, 0, 0)                 # тело глифа — чёрное
    arr[8:32:2, 8:32:3] = (144, 58, 0)            # цветной ореол, но вразнобой
    arr[9:31:3, 9:31:2] = (58, 144, 219)
    box = BBox(5, 5, 35, 35)
    assert ink_color(arr, box) == (0, 0, 0)
    assert not is_markup_number(arr, box)


def test_markup_number_accepts_colored_label():
    """Бордовые, оранжевые и синие подписи — все свои."""
    from repair.map_labels import is_markup_number

    for цвет in ((128, 0, 0), (255, 144, 60), (0, 0, 255), (0, 144, 255)):
        arr = np.full((40, 40, 3), 255, np.uint8)
        arr[10:30, 10:30] = цвет
        assert is_markup_number(arr, BBox(5, 5, 35, 35)), цвет


def test_pixels_ink_keeps_bright_colors():
    """
    Порог «чернил» — по минимальному каналу.

    У оранжевой (255, 144, 60) и ярко-красной (255, 0, 0) подписи максимум
    равен 255: проверка по максимуму выбрасывала их целиком, и решение
    принимал тёмный дизер вокруг.
    """
    from repair.map_labels import pixels_ink

    пиксели = np.array([[255, 144, 60], [255, 0, 0], [128, 0, 0], [0, 0, 0],
                        [255, 255, 255], [255, 255, 219]])
    assert list(pixels_ink(пиксели)) == [True, True, True, True, False, False]
