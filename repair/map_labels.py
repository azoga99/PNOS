# -*- coding: utf-8 -*-
"""
Поиск подписей сечений на коррозионной карте.

Зачем: чтобы сотрудник не искал сечение глазами. Если известно, где на карте
подпись каждого сечения, программа сама подводит к нужному месту, сама
подставляет новое значение и сама показывает «до/после».

Как это работает и почему именно так:

* **Карта распознаётся целиком, одним вызовом OCR.** Windows OCR возвращает
  слова вместе с рамками — этого достаточно: подпись сечения выглядит как
  «112 (5.9)», то есть число, а за ним значение в скобках. На карте п.2763
  (2970×4200, около 125 сечений) полный проход занимает полсекунды и находит
  14 из 16 подписей, проверенных глазами.
* **Выращивать рамки по цветной маске оказалось хуже.** Первая версия искала
  подписи по маске разметки и растила рамки из затравок: рамки слипались с
  трубопроводом и соседними подписями, номера читались неверно, и в выдаче
  один и тот же номер появлялся десятки раз. Этот путь отброшен.
* **Несколько масштабов — для полноты.** На каждом масштабе OCR теряет свои
  подписи (при 1.0 не нашлись 110 и 112, при 1.5 — 112 и 116), поэтому
  результаты объединяются.
* **Значение в скобках распознаётся плохо** («(5.9)» → «(59)») и как текст не
  используется: его всё равно предстоит заменить, а вырезает его геометрия
  (`find_value_span`).

Слой чистый (numpy + Pillow + Windows OCR), Qt не знает.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import numpy as np

from repair.map_editor import BBox

# Номер сечения: 1–3 цифры отдельным словом.
_NUM_RE = re.compile(r"^(\d{1,3})$")
# Значение замера: «(5.9)», OCR часто теряет точку или закрывающую скобку.
_VAL_RE = re.compile(r"^\(\s*(\d{1,3})[.,]?(\d)?\)?$")
# Номера сечений не бывают больше этого — отсекает числа из штампа и таблицы.
MAX_SECTION = 500
# Масштабы распознавания: на каждом OCR теряет свои подписи, объединение даёт
# заметно больше (на карте п.2763: 100 подписей при двух масштабах, 107 при трёх).
SCALES = (1.0, 1.5, 2.0)
# У Windows OCR есть предел размера изображения (MaxImageDimension): карта
# 2970×4200 при масштабе 2.5 его превышает, и вызов падает с WinError.
MAX_OCR_DIM = 9800
# Две подписи с одним номером ближе этого расстояния — одна и та же подпись,
# найденная на разных масштабах.
SAME_LABEL_PX = 30


@dataclass
class MapLabel:
    """Подпись сечения на карте: «NN (V.V)»."""
    number: int
    point: tuple[int, int]        # точка внутри номера — для render_value
    bbox: BBox                    # рамка подписи целиком
    value_bbox: BBox | None = None   # рамка части «(V.V)» — её и заменяют
    value_text: str = ""          # как OCR прочитал значение (справочно)
    scales: list = field(default_factory=list)   # на каких масштабах найдена

    @property
    def center(self) -> tuple[int, int]:
        return ((self.bbox.x1 + self.bbox.x2) // 2,
                (self.bbox.y1 + self.bbox.y2) // 2)


def _word_box(wd: dict, scale: float) -> BBox:
    return BBox(int(wd["x"] / scale), int(wd["y"] / scale),
                int((wd["x"] + wd["w"]) / scale),
                int((wd["y"] + wd["h"]) / scale))


def _labels_at_scale(arr: np.ndarray, scale: float) -> list[MapLabel]:
    from PIL import Image
    from repair.ocr import ocr_lines

    img = Image.fromarray(arr)
    if scale != 1.0:
        img = img.resize((int(img.width * scale), int(img.height * scale)),
                         Image.LANCZOS)
    out: list[MapLabel] = []
    for line in ocr_lines(np.array(img)):
        words = line.get("words") or []
        for i, wd in enumerate(words):
            m = _NUM_RE.match(wd["t"].strip())
            if not m:
                continue
            номер = int(m.group(1))
            if not (1 <= номер <= MAX_SECTION):
                continue
            # Подпись обязана иметь значение в скобках: иначе это число из
            # штампа, таблицы элементов или обозначения диаметра.
            if i + 1 >= len(words):
                continue
            зн = words[i + 1]["t"].strip()
            if not _VAL_RE.match(зн):
                continue
            nb = _word_box(wd, scale)
            vb = _word_box(words[i + 1], scale)
            out.append(MapLabel(
                number=номер,
                point=((nb.x1 + nb.x2) // 2, (nb.y1 + nb.y2) // 2),
                bbox=BBox(min(nb.x1, vb.x1), min(nb.y1, vb.y1),
                          max(nb.x2, vb.x2), max(nb.y2, vb.y2)),
                value_bbox=vb, value_text=зн, scales=[scale]))
    return out


def find_labels(arr: np.ndarray, scales=SCALES, log=None) -> list[MapLabel]:
    """
    Найти подписи сечений на карте. Возвращает список, отсортированный сверху
    вниз. Один номер может встретиться несколько раз — это нормально
    (одинаковые участки), выбирать место правки должен сотрудник.
    """
    h, w = arr.shape[:2]
    предел = MAX_OCR_DIM / max(h, w)
    годные = [s for s in scales if s <= предел]
    if log and len(годные) < len(scales):
        log(f"Масштабы больше {предел:.1f} пропущены: карта {w}×{h}, "
            f"Windows OCR не берёт изображения крупнее {MAX_OCR_DIM} px")

    найдено: list[MapLabel] = []
    for scale in годные:
        порция = _labels_at_scale(arr, scale)
        добавлено = 0
        for lb in порция:
            двойник = next(
                (x for x in найдено if x.number == lb.number
                 and abs(x.point[0] - lb.point[0]) <= SAME_LABEL_PX
                 and abs(x.point[1] - lb.point[1]) <= SAME_LABEL_PX), None)
            if двойник is not None:
                двойник.scales.append(scale)
                continue
            найдено.append(lb)
            добавлено += 1
        if log:
            log(f"Карта, масштаб {scale}: подписей {len(порция)}, "
                f"новых {добавлено}")

    найдено.sort(key=lambda l: (l.bbox.y1, l.bbox.x1))
    if log:
        log(f"Карта: всего подписей {len(найдено)}, "
            f"разных сечений {len({l.number for l in найдено})}")
    return найдено


def labels_by_number(labels: list[MapLabel]) -> dict[int, list[MapLabel]]:
    """
    Подписи по номеру сечения.

    Список, а не одна подпись: один и тот же номер на карте встречается
    несколько раз (одинаковые переходы, повторяющиеся участки), и выбирать,
    какое именно место править, должен сотрудник.
    """
    out: dict[int, list[MapLabel]] = {}
    for lb in labels:
        out.setdefault(lb.number, []).append(lb)
    for номер in out:
        out[номер].sort(key=lambda l: (l.bbox.y1, l.bbox.x1))
    return out
