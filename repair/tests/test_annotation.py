# -*- coding: utf-8 -*-
"""
Тесты снятия отметки ремонта с коррозионной карты.

Синтетика проверяет главное свойство: чертёж под полосой не страдает, а
законное красное примечание в стороне остаётся на месте. Реальные карты
проверяют то, ради чего всё написано, — что после стирания на листе не
остаётся ни пикселя выделения; при их отсутствии тесты пропускаются.
"""

import os

import numpy as np
import pytest

from repair.annotation import (Annotation, bright_red, erase_annotation,
                               find_annotation)

БЕЛЫЙ = (255, 255, 255)
КРАСНЫЙ = (255, 0, 0)
ЧЁРНЫЙ = (0, 0, 0)


def _лист(h=400, w=600) -> np.ndarray:
    return np.full((h, w, 3), 255, np.uint8)


def _текст(arr, y1, y2, x1, цвет=КРАСНЫЙ, букв=6, шаг=20, штрих=4):
    """
    «Надпись» тонкими штрихами.

    Толщина принципиальна: разметка чертежа рисуется штрихом в несколько
    пикселей и эрозию в `find_markers` не переживает — этим выделение от неё и
    отличается. Сплошные блоки вместо букв сделали бы тест неверным.
    """
    for i in range(букв):
        x = x1 + i * шаг
        arr[y1:y2, x:x + штрих] = цвет


def _с_выделением(подпись=True) -> np.ndarray:
    """
    Лист: чёрная труба, красная полоса ПОД ней, выноска, подчёркивание и
    надпись — как рисует ОТН.
    """
    arr = _лист()
    arr[180:230, 60:260] = КРАСНЫЙ          # полоса выделения (жирная)
    arr[198:212, 40:300] = ЧЁРНЫЙ           # труба поверх полосы
    arr[156:182, 256:260] = КРАСНЫЙ         # выноска от полосы вверх
    arr[152:156, 258:430] = КРАСНЫЙ         # подчёркивание надписи
    if подпись:
        _текст(arr, 110, 145, 280)
    return arr


# ── маска заливки ────────────────────────────────────────────────────

@pytest.mark.parametrize("цвет, ожидание", [
    ((255, 0, 0), True),        # заливка выделения
    ((250, 20, 20), True),
    ((128, 0, 0), False),       # бордовые подписи сечений
    ((255, 144, 60), False),    # оранжевые подписи
    ((214, 102, 0), False),     # рыжая кайма у оранжевого крестика
    ((0, 0, 255), False),       # синие подписи
    ((0, 0, 0), False),
    ((255, 255, 255), False),
])
def test_bright_red_отличает_заливку(цвет, ожидание):
    arr = np.full((2, 2, 3), цвет, np.uint8)
    assert bool(bright_red(arr).all()) is ожидание


# ── поиск и стирание ─────────────────────────────────────────────────

def test_без_выделения_ничего_не_находится():
    assert find_annotation(_лист()) is None


def test_законное_красное_примечание_не_считается_выделением():
    """Тонкий красный текст в штампе — не полоса; файл трогать нельзя."""
    arr = _лист()
    _текст(arr, 350, 380, 40, букв=8, шаг=18)
    assert find_annotation(arr) is None


def test_полоса_находится_и_стирается():
    arr = _с_выделением(подпись=False)
    ann = find_annotation(arr)
    assert ann is not None and ann.marker_px > 0
    erase_annotation(arr, ann)
    assert not bright_red(arr).any()


def test_линия_трубы_остаётся_целой():
    """Полоса лежит под чертежом: после стирания труба не рвётся."""
    arr = _с_выделением(подпись=False)
    труба_до = np.all(arr[198:212, 40:300] == ЧЁРНЫЙ, axis=2)
    erase_annotation(arr, find_annotation(arr))
    труба_после = np.all(arr[198:212, 40:300] == ЧЁРНЫЙ, axis=2)
    assert np.array_equal(труба_до, труба_после)
    assert труба_после.all()


def test_подпись_рядом_убирается_рамкой():
    arr = _с_выделением(подпись=True)
    ann = find_annotation(arr)
    assert ann.caption is not None and ann.caption_boxed
    erase_annotation(arr, ann)
    assert not bright_red(arr).any()


def test_подпись_с_чужой_краской_рамкой_не_стирается():
    """Если рядом с надписью чужой чертёж — рамкой бить нельзя."""
    arr = _с_выделением(подпись=True)
    arr[118:130, 306:420] = (0, 0, 255)     # синяя подпись поверх надписи
    ann = find_annotation(arr)
    assert ann.caption is not None and not ann.caption_boxed
    erase_annotation(arr, ann)
    assert np.all(arr[118:130, 306:420] == (0, 0, 255))


def test_далёкое_примечание_переживает_стирание():
    """Красный текст в штампе далеко от выделения не должен пострадать."""
    arr = _с_выделением(подпись=False)
    _текст(arr, 350, 380, 380, букв=8, шаг=18)
    было = arr[350:380, 380:530].copy()
    erase_annotation(arr, find_annotation(arr))
    assert np.array_equal(arr[350:380, 380:530], было)


def test_erase_annotation_пустой_вход():
    assert erase_annotation(_лист(), None) == 0


# ── реальные карты ───────────────────────────────────────────────────

_С_ОТМЕТКОЙ = [
    r"C:\Users\volod\Desktop\п.2687\Первичка\2687.png",
    r"C:\Users\volod\Desktop\п.2687\Первичка\2687..png",
    r"C:\Users\volod\Desktop\п.2687(запросить акт отбр)\Первичка\2687.png",
]
_БЕЗ_ОТМЕТКИ = [
    r"C:\Users\volod\Desktop\Сданы на ПНОС — копия"
    r"\п.2788 (Всем отправлено)\Первичка\2788.png",
    r"C:\Users\volod\Desktop\Сданы на ПНОС — копия"
    r"\п.2792 (Всем отправлено)\Первичка\2792.png",
]


def _карта(path):
    if not os.path.isfile(path):
        pytest.skip("реальных карт на этой машине нет")
    from PIL import Image
    return np.array(Image.open(path).convert("RGB"))


@pytest.mark.parametrize("path", _С_ОТМЕТКОЙ)
def test_реальная_карта_очищается_начисто(path):
    arr = _карта(path)
    было = arr.copy()
    ann = find_annotation(arr)
    assert ann is not None, "выделение на карте п.2687 должно находиться"
    стёрто = erase_annotation(arr, ann)
    assert стёрто > 0
    # ни одного пикселя заливки не осталось
    assert not bright_red(arr).any()
    # изменилось только то, что попало в маску
    изменилось = np.any(было != arr, axis=2)
    assert np.array_equal(изменилось, изменилось & ann.mask)
    # правка локальна: меньше процента листа
    assert стёрто < 0.01 * arr.shape[0] * arr.shape[1]


@pytest.mark.parametrize("path", _С_ОТМЕТКОЙ[:2])
def test_после_стирания_не_остаётся_ореола(path):
    """
    Ни одного цветного пикселя там, где была надпись.

    Глифы сглажены цветным дизером, и он выходит за пределы красного: при
    фиксированном запасе рамки под подчёркиванием оставалась россыпь из
    двух десятков розовых точек — на листе их видно.
    """
    arr = _карта(path)
    erase_annotation(arr, find_annotation(arr))
    зона = arr[1980:2080, 2740:3340].astype(int)
    не_белые = ~np.all(зона >= 248, axis=2)
    assert int(не_белые.sum()) == 0


@pytest.mark.parametrize("path", _БЕЗ_ОТМЕТКИ)
def test_пункт_без_ремонта_не_трогается(path):
    """У этих пунктов чистым красным написано примечание в штампе."""
    arr = _карта(path)
    assert bright_red(arr).any(), "на этих листах красное примечание есть"
    assert find_annotation(arr) is None
