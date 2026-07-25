# -*- coding: utf-8 -*-
"""
Подбор сечений по таблице замены.

Из документов ремонта известно, ЧТО заменено: «Переход 108х6-89х6»,
«Труба 219х10 L=2940мм», «Отвод П90 219х12 4шт.». В листе «Импорт» у каждого
сечения есть тип элемента, исполнительная толщина и диаметр. Отсюда программа
может показать, какие сечения подходят под запись о замене, и предложить новый
номинал — сотруднику остаётся выбрать место, а не искать его перебором.

Главное ограничение, ради которого всё и написано так: **однозначного ответа
обычно нет**. В пункте 2775 «Труба 108» встречается 26 раз, «Отвод 108» — 11,
«Переход 159 / 108» — 4. Поэтому функция возвращает СПИСОК кандидатов, а
выбор остаётся за человеком — по коррозионной карте, где видно, какой именно
участок заменён.

Слой чистый: ни Qt, ни Excel, ни OCR.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from repair.model import Section

# Типы элементов, которые могут стоять в таблице замены.
ELEMENTS = ("Труба", "Отвод", "Переход", "Тройник", "Фланец", "Заглушка",
            "Бобышка", "Задвижка", "Опора")
# Порог нечёткого сравнения названий: OCR глотает буквы («Пе еход»).
ELEMENT_MIN = 75
# Типоразмер: «108х6», «219х10», «89х3.5», «32х4»
_SIZE_RE = re.compile(r"(\d{2,3})\s*[хx×]\s*(\d{1,2}(?:[.,]\d)?)")


@dataclass
class Suggestion:
    """Одна строка таблицы замены и подходящие ей сечения."""
    строка: str
    элемент: str = ""
    размеры: list = field(default_factory=list)     # [(диаметр, толщина), …]
    кандидаты: list = field(default_factory=list)   # list[Section]

    @property
    def диаметр(self) -> str:
        """Диаметр в том виде, как он записан на листе «Импорт»."""
        return " / ".join(d for d, _ in self.размеры)

    @property
    def номинал(self) -> str:
        """Предлагаемый новый номинал: толщина новой трубы."""
        толщины = [t.replace(",", ".") for _, t in self.размеры]
        if not толщины:
            return ""
        if len(толщины) == 1 or len(set(толщины)) == 1:
            return толщины[0]
        return " / ".join(толщины)      # переход с разными стенками

    @property
    def однозначно(self) -> bool:
        return len(self.кандидаты) == 1


def _similar(a: str, b: str) -> int:
    from rapidfuzz import fuzz
    return int(fuzz.partial_ratio(a.lower(), b.lower()))


def parse_replacement(line: str) -> Suggestion:
    """
    Разобрать строку таблицы замены → элемент и типоразмеры.

    Понимает записи вида «Переход 108х6-89х6», «Труба 219х10 L=2940мм»,
    «Отвод П90 219х12 4шт.». Мусор OCR не отсеивается здесь: если размеры не
    разобрались, кандидатов просто не будет.
    """
    s = Suggestion(строка=(line or "").strip())
    if not s.строка:
        return s

    лучший, лучшая = "", 0
    for эл in ELEMENTS:
        оценка = _similar(эл, s.строка)
        if оценка > лучшая:
            лучший, лучшая = эл, оценка
    if лучшая >= ELEMENT_MIN:
        s.элемент = лучший

    for m in _SIZE_RE.finditer(s.строка):
        s.размеры.append((m.group(1), m.group(2)))
    return s


def _diameters(текст: str) -> set:
    return set(re.findall(r"\d{2,3}", текст or ""))


def suggest_sections(lines, sections: list[Section],
                     exclude=("Задвижка",)) -> list[Suggestion]:
    """
    Для каждой строки таблицы замены найти подходящие сечения.

    Совпадение: тип элемента (нечётко — OCR коверкает названия) и полный
    набор диаметров. Задвижки исключены: по правилам их номиналы не обновляют.
    """
    if isinstance(lines, str):
        lines = [lines]

    итог = []
    for line in lines:
        s = parse_replacement(line)
        нужные = {d for d, _ in s.размеры}
        if s.элемент and нужные:
            for sec in sections:
                if any(_similar(e, sec.element) >= 90 for e in exclude):
                    continue
                if _similar(s.элемент, sec.element) < ELEMENT_MIN:
                    continue
                if _diameters(sec.diameter) == нужные:
                    s.кандидаты.append(sec)
        s.кандидаты.sort(key=lambda x: x.number)
        итог.append(s)
    return итог


def describe(s: Suggestion) -> str:
    """Короткое пояснение для журнала."""
    if not s.элемент or not s.размеры:
        return f"«{s.строка[:50]}» — типоразмер не разобран"
    номера = ", ".join(str(c.number) for c in s.кандидаты)
    if not s.кандидаты:
        return (f"{s.элемент} {s.диаметр} — подходящих сечений не найдено")
    if s.однозначно:
        return (f"{s.элемент} {s.диаметр} → сечение {номера}, "
                f"номинал {s.номинал}")
    return (f"{s.элемент} {s.диаметр} → возможные сечения: {номера}; "
            f"номинал {s.номинал} — выберите место по карте")
