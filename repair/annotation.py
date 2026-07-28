# -*- coding: utf-8 -*-
"""
Снятие отметки ремонта с коррозионной карты.

Инженер ОТН помечает заменяемый участок жирной красной полосой, тянет от неё
тонкую выноску и подписывает «Коррозионный износ». После ремонта эта отметка
уже неверна: участок заменён, номиналы на карте исправлены — а красное
по-прежнему кричит про износ. Карта попадает в отчёт ЭПБ картинкой один в
один (в п.2687 `word/media/image6.png` совпадает с `2687.png` попиксельно),
поэтому отметку надо убирать именно с PNG.

**Почему это удаётся сделать начисто.** Карта — не скан, а экспорт из
векторного редактора с квантованной палитрой (383 цвета на весь лист п.2687).
Полоса залита ровно одним цветом (255, 0, 0) и **лежит под чертежом**: чёрная
линия отвода нарисована поверх неё и не прерывается. Кольцо пикселей вокруг
полосы содержит только чужую краску (чёрный, оранжевый, бордовый, синий) — ни
одного пикселя-смеси. Значит стирание по цвету не оставляет ни розовой каймы,
ни разрыва в трубе, и дорисовывать ничего не нужно.

Надпись — другое дело: она сглажена и осыпается цветным дизером, который по
цвету не отловить. Зато она стоит на пустом белом поле, поэтому стирается
рамкой целиком. Условие проверяется перед стиранием (`_foreign_ink`): если в
рамке оказалась чужая краска, рамкой не стираем.

**Чего здесь принципиально нет — правила «убрать весь красный».** На семи
сданных пунктах (2778, 2779, 2785, 2786, 2788, 2792, 2793) тем же чистым
красным написано законное примечание в штампе — «Отсутствует журнальная
запись о ремонте 2018 года». Поэтому отсчёт всегда идёт от найденного
выделения (`repair.highlight.find_markers`): нет полосы — файл не трогаем.

Слой чистый: только numpy, ни Qt, ни Pillow, ни COM.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np

from repair.highlight import find_markers
from repair.map_editor import BBox

#: Заливка выделения: яркий насыщенный красный. Пороги подобраны так, чтобы
#: мимо прошла вся остальная краска карты: бордовые подписи (128, 0, 0) —
#: по R, оранжевые (255, 144, 60) и рыжая кайма (214, 102, 0) — по G.
HL_MIN_R = 200
HL_MAX_GB = 90

#: На сколько пикселей от выделения искать его подпись. Надпись «Коррозионный
#: износ» в п.2687 стоит вплотную к подчёркиванию, которое само является частью
#: выделения; запас взят с большим избытком.
CAPTION_NEAR = 90

#: Поля рамки надписи при стирании.
CAPTION_PAD = 8

#: Насколько рамке надписи разрешено разъехаться в поисках чистой границы.
MAX_CAPTION_GROW = 40

#: Почти белый фон — всё, что светлее, считаем бумагой.
WHITE_MIN = 248


@dataclass
class Annotation:
    """Что именно предстоит стереть с листа."""
    mask: np.ndarray                      # булева маска пикселей под стирание
    marker_px: int = 0                    # полоса + выноска + подчёркивание
    caption: BBox | None = None           # рамка надписи, если она найдена
    caption_boxed: bool = False           # рамку можно стереть целиком
    bbox: BBox | None = None              # общая рамка отметки
    markers: list = field(default_factory=list)

    @property
    def pixels(self) -> int:
        return int(self.mask.sum())


def bright_red(arr: np.ndarray) -> np.ndarray:
    """Маска заливки выделения — только яркий чистый красный."""
    r = arr[:, :, 0].astype(int)
    g = arr[:, :, 1].astype(int)
    b = arr[:, :, 2].astype(int)
    return (r >= HL_MIN_R) & (g <= HL_MAX_GB) & (b <= HL_MAX_GB)


#: Разрыв, через который куски отметки считаются одним целым. Выноска местами
#: истончается до пунктира, поэтому строгая связность по соседним пикселям
#: разорвала бы её на части.
GAP = 6


def _labels(mask: np.ndarray, gap: int = GAP) -> np.ndarray:
    """
    Разметка связных кусков маски: массив номеров, −1 вне маски.

    Обход идёт по КООРДИНАТАМ пикселей через сетку соседства, а не сдвигами
    всего изображения. На листе карты 4200×2970 пикселей заливки тысячи, а
    самих пикселей 12 млн: наращивание маски сдвигами numpy потребовало бы
    сотни проходов по всему массиву и считалось бы минутами.
    """
    ys, xs = np.nonzero(mask)
    out = np.full(mask.shape, -1, np.int32)
    n = len(xs)
    if not n:
        return out
    сетка: dict = {}
    for i in range(n):
        сетка.setdefault((ys[i] // gap, xs[i] // gap), []).append(i)

    не_посещены = np.zeros(n, bool)
    метка = 0
    for старт in range(n):
        if не_посещены[старт]:
            continue
        стек = [старт]
        не_посещены[старт] = True
        while стек:
            i = стек.pop()
            yi, xi = ys[i], xs[i]
            out[yi, xi] = метка
            for cy in range(yi // gap - 1, yi // gap + 2):
                for cx in range(xi // gap - 1, xi // gap + 2):
                    for j in сетка.get((cy, cx), ()):
                        if (not не_посещены[j]
                                and abs(ys[j] - yi) <= gap
                                and abs(xs[j] - xi) <= gap):
                            не_посещены[j] = True
                            стек.append(j)
        метка += 1
    return out


def _label_at(lab: np.ndarray, x: int, y: int, win: int = 12) -> int:
    """Номер куска в точке или в ближайшей её окрестности; −1 — не найден."""
    h, w = lab.shape
    y1, y2 = max(0, y - win), min(h, y + win + 1)
    x1, x2 = max(0, x - win), min(w, x + win + 1)
    зона = lab[y1:y2, x1:x2]
    свои = зона[зона >= 0]
    if not свои.size:
        return -1
    значения, счёт = np.unique(свои, return_counts=True)
    return int(значения[int(np.argmax(счёт))])


def _bbox_of(mask: np.ndarray) -> BBox | None:
    ys, xs = np.nonzero(mask)
    if not len(xs):
        return None
    return BBox(int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1)


def _is_white(zone: np.ndarray) -> np.ndarray:
    return ((zone[:, :, 0] >= WHITE_MIN) & (zone[:, :, 1] >= WHITE_MIN)
            & (zone[:, :, 2] >= WHITE_MIN))


def _grow_clean_box(arr: np.ndarray, box: BBox, limit: int = MAX_CAPTION_GROW) -> BBox:
    """
    Раздвинуть рамку надписи до чистого белого поля с каждой стороны.

    Фиксированного запаса не хватает: глифы сглажены цветным дизером, и его
    ореол выходит за рамку на непредсказуемые несколько пикселей (на карте
    п.2687 — на 5 px вниз от подчёркивания). Поэтому каждая сторона едет
    наружу, пока приграничная линия не окажется полностью белой, и
    останавливается, едва в рамку начинает попадать чужая краска.
    """
    h, w = arr.shape[:2]
    x1, y1, x2, y2 = box.x1, box.y1, box.x2, box.y2
    for _ in range(limit):
        двигали = False
        # (новая координата, срез приграничной линии)
        стороны = [
            ("y1", y1 - 1, (slice(max(0, y1 - 1), y1), slice(x1, x2))),
            ("y2", y2 + 1, (slice(y2, min(h, y2 + 1)), slice(x1, x2))),
            ("x1", x1 - 1, (slice(y1, y2), slice(max(0, x1 - 1), x1))),
            ("x2", x2 + 1, (slice(y1, y2), slice(x2, min(w, x2 + 1)))),
        ]
        for имя, новое, срез in стороны:
            if имя == "y1" and новое < 0:
                continue
            if имя == "x1" and новое < 0:
                continue
            if имя == "y2" and новое > h:
                continue
            if имя == "x2" and новое > w:
                continue
            линия = arr[срез]
            if not линия.size or _is_white(линия).all():
                continue                    # граница уже чистая
            проба = BBox(новое if имя == "x1" else x1,
                         новое if имя == "y1" else y1,
                         новое if имя == "x2" else x2,
                         новое if имя == "y2" else y2)
            if _foreign_ink(arr, проба):
                continue                    # дальше чужой чертёж — стоп
            x1, y1, x2, y2 = проба.x1, проба.y1, проба.x2, проба.y2
            двигали = True
        if not двигали:
            break
    return BBox(x1, y1, x2, y2)


def _foreign_ink(arr: np.ndarray, box: BBox) -> int:
    """
    Сколько в рамке пикселей чужой краски — не фон и не красное семейство.

    Красным семейством считаем всё с высоким R: сама надпись, её сглаживание и
    цветной дизер вокруг глифов имеют R близкий к максимуму, а чёрная линия,
    синяя и бордовая подписи — нет. Ноль на выходе означает, что в рамке нет
    ничего, кроме надписи, и её можно стирать целиком.
    """
    zone = arr[box.y1:box.y2, box.x1:box.x2]
    if not zone.size:
        return 0
    r = zone[:, :, 0].astype(int)
    g = zone[:, :, 1].astype(int)
    b = zone[:, :, 2].astype(int)
    white = (r >= WHITE_MIN) & (g >= WHITE_MIN) & (b >= WHITE_MIN)
    red_family = r >= HL_MIN_R
    return int((~white & ~red_family).sum())


def find_annotation(arr: np.ndarray, near: int = CAPTION_NEAR) -> Annotation | None:
    """
    Найти отметку ремонта на листе карты: полоса, выноска и подпись к ним.

    Возвращает None, если жирного выделения на листе нет — тогда файл не
    трогается вовсе. Это и есть защита от стирания законных красных примечаний
    в штампе: они выделением не являются.
    """
    if arr is None or getattr(arr, "ndim", 0) != 3:
        return None
    метки = find_markers(arr)
    if not метки:
        return None

    красное = bright_red(arr)
    h, w = arr.shape[:2]

    # 1. Связная область выделения: полоса + тонкая выноска + подчёркивание.
    lab = _labels(красное)
    свои = {n for n in (_label_at(lab, *m.center) for m in метки) if n >= 0}
    if not свои:
        return None
    полоса = np.isin(lab, list(свои))
    mask = полоса.copy()
    рамка_полосы = _bbox_of(полоса)

    # 2. Подпись рядом: остальное красное в окрестности выделения. Она отдельным
    #    куском, поэтому ростом области её не достать.
    ext = рамка_полосы.pad(near, w, h)
    окрестность = np.zeros((h, w), bool)
    окрестность[ext.y1:ext.y2, ext.x1:ext.x2] = True
    подпись_px = красное & окрестность & ~полоса

    caption, boxed = None, False
    if подпись_px.any():
        caption = _bbox_of(подпись_px).pad(CAPTION_PAD, w, h)
        if _foreign_ink(arr, caption) == 0:
            caption = _grow_clean_box(arr, caption)
            # В рамке только надпись — стираем её целиком вместе со сглаживанием.
            mask[caption.y1:caption.y2, caption.x1:caption.x2] = True
            boxed = True
        else:
            # Чужая краска рядом: ограничиваемся точным стиранием по цвету.
            mask |= подпись_px

    return Annotation(mask=mask, marker_px=int(полоса.sum()), caption=caption,
                      caption_boxed=boxed, bbox=_bbox_of(mask), markers=метки)


def erase_annotation(arr: np.ndarray, ann: Annotation) -> int:
    """Стереть отметку (правка массива на месте). Возвращает число пикселей."""
    if ann is None or not ann.mask.any():
        return 0
    arr[ann.mask] = (255, 255, 255)
    return ann.pixels
