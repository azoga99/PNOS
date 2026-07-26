# -*- coding: utf-8 -*-
"""
Ядро пересчёта: применение номинальных толщин к заменённым сечениям.

Правило (подтверждено ретро-тестами на п.2701 и п.2775):
    заменённый участок — труба новая, мерить нечего → всем точкам замеров,
    минимуму и исполнительной толщине проставляется номинал новой трубы;
    диаметр сохраняется (кроме перехода — он двусторонний); тип не меняется.

Поля правки раздельные (`SectionEdit`): Тисп, диаметр и замер вводятся по
отдельности. Единственное поле «новый номинал» этого не позволяло: у перехода
Тисп записана двусторонне («6 / 6»), и число 6.0 затирало её половиной —
на п.2763 сечение 114 «6 / 6 · 108 / 89» превращалось в «6 · 108 / 89».
Замер по умолчанию выводится из Тисп, но его можно задать явно.

Функция чистая: не пишет в файлы, только считает и возвращает новые сечения
и список изменений для экрана предпросмотра.
"""

from __future__ import annotations

from dataclasses import dataclass

from repair.model import Section, Change


def _fmt_measure(nominal: float) -> str:
    """Замер выводится с одним знаком: 12.0, 8.0, 5.0."""
    return f"{nominal:.1f}"


def _fmt_thickness(nominal: float) -> str:
    """Исполнительная толщина: целое без дробной части (12), иначе как есть (3.5)."""
    return str(int(nominal)) if float(nominal) == int(nominal) else str(nominal)


@dataclass
class SectionEdit:
    """
    Правка одного сечения. Пустое поле = «не менять».

    thickness — Тисп в том виде, как она стоит на листе: «6», «12 / 10»;
    diameter  — диаметр: «108», «108 / 89»;
    measure   — значение всех точек замеров и минимума: «6.0». Пусто —
                выводится из thickness (труба новая, мерить нечего).
    """
    thickness: str = ""
    diameter: str = ""
    measure: str = ""

    def is_empty(self) -> bool:
        return not (self.thickness.strip() or self.diameter.strip()
                    or self.measure.strip())


def measure_from_thickness(thickness: str) -> str:
    """
    Замер новой трубы = её номинальная толщина одним знаком: «6» → «6.0».

    У перехода («12 / 10») берётся первое число: сечение относится к одной
    стороне перехода, вторая живёт своим сечением.
    """
    first = (thickness or "").replace(",", ".").split("/")[0].strip()
    try:
        return f"{float(first):.1f}"
    except ValueError:
        return ""


def apply_edits(
    sections: list[Section],
    edits: dict,
    *,
    by_row: bool = False,
) -> tuple[list[Section], list[Change]]:
    """
    sections — все сечения листа «Импорт» (исходные);
    edits    — {ключ: SectionEdit}; ключ — № сечения (by_row=False) или строка
               листа «Импорт» (by_row=True — живой путь: номера не уникальны).

    Возвращает (новые сечения, список изменений для предпросмотра).
    """
    new_sections: list[Section] = []
    changes: list[Change] = []

    for sec in sections:
        key = sec.row if by_row else sec.number
        e = edits.get(key)
        if e is None or e.is_empty():
            new_sections.append(sec)
            continue

        thick = e.thickness.strip() or sec.thickness_nom
        diam = e.diameter.strip() or sec.diameter
        measure = e.measure.strip() or measure_from_thickness(e.thickness)

        if measure:
            # номинал только в заполненные точки: пустая точка не «дорисовывается»
            measures = [measure if old else "" for old in sec.measures]
            tmin = measure
        else:
            measures, tmin = list(sec.measures), sec.tmin

        new_sec = Section(
            number=sec.number,
            element=sec.element,
            thickness_nom=thick,
            diameter=diam,
            measures=measures,
            tmin=tmin,
            row=sec.row,
        )
        new_sections.append(new_sec)

        # Изменение для предпросмотра: сжатое «было → станет».
        before = (f"Тисп {sec.thickness_nom} · {sec.diameter} · "
                  f"замеры {'/'.join(sec.measures)}")
        after = (f"Тисп {thick} · {diam} · "
                 f"замеры {'/'.join(measures)}")
        changes.append(Change(
            target="Импорт",
            where=f"сечение {sec.number}",
            before=before,
            after=after,
        ))

    return new_sections, changes


def apply_nominals(
    sections: list[Section],
    replaced: dict[int, float],
    perehod: dict[int, tuple[str, str]] | None = None,
    *,
    by_row: bool = False,
) -> tuple[list[Section], list[Change]]:
    """
    sections — все сечения листа «Импорт» (исходные).
    replaced — {ключ: новый номинал толщины, мм} из отметки сотрудника.
    perehod  — {ключ: (толщина_строкой, диаметр_строкой)} для переходов
               (двусторонние значения, напр. ("12 / 10", "273 / 219")).

    Ключ — № сечения (by_row=False, как в фикстурах/отчётах) или строка листа
    «Импорт» (by_row=True — живой путь: номера сечений не уникальны).

    Короткая форма поверх `apply_edits`: остаётся для ретро-тестов на готовых
    отчётах, где заменённое сечение задаётся одним числом.

    Возвращает (новые сечения, список изменений).
    """
    perehod = perehod or {}
    edits: dict = {}
    for key, nominal in replaced.items():
        m = _fmt_measure(nominal)
        if key in perehod:
            thick, diam = perehod[key]
            edits[key] = SectionEdit(thickness=thick, diameter=diam, measure=m)
        else:
            edits[key] = SectionEdit(thickness=_fmt_thickness(nominal), measure=m)
    return apply_edits(sections, edits, by_row=by_row)


def corrosion_rate_for_replaced(min_rate: float) -> float:
    """
    Скорость коррозии заменённого участка: утонения нет → минимальный порог.
    (Остаточный ресурс = (Sф − Sотбр) / rate; срок эксплуатации t не меняется.)
    """
    return min_rate
