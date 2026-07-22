# -*- coding: utf-8 -*-
"""
Ядро пересчёта: применение номинальных толщин к заменённым сечениям.

Правило (подтверждено ретро-тестами на п.2701 и п.2775):
    заменённый участок — труба новая, мерить нечего → всем точкам замеров,
    минимуму и исполнительной толщине проставляется номинал новой трубы;
    диаметр сохраняется (кроме перехода — он двусторонний); тип не меняется.

Функция чистая: не пишет в файлы, только считает и возвращает новые сечения
и список изменений для экрана предпросмотра.
"""

from __future__ import annotations

from repair.model import Section, Change


def _fmt_measure(nominal: float) -> str:
    """Замер выводится с одним знаком: 12.0, 8.0, 5.0."""
    return f"{nominal:.1f}"


def _fmt_thickness(nominal: float) -> str:
    """Исполнительная толщина: целое без дробной части (12), иначе как есть (3.5)."""
    return str(int(nominal)) if float(nominal) == int(nominal) else str(nominal)


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

    Возвращает (новые сечения, список изменений).
    """
    perehod = perehod or {}
    new_sections: list[Section] = []
    changes: list[Change] = []

    for sec in sections:
        key = sec.row if by_row else sec.number
        if key not in replaced:
            new_sections.append(sec)
            continue

        nominal = replaced[key]

        if key in perehod:
            new_thick, new_diam = perehod[key]
        else:
            new_thick, new_diam = _fmt_thickness(nominal), sec.diameter

        m = _fmt_measure(nominal)
        # номинал только в заполненные точки: пустая точка не «дорисовывается»
        new_measures = [m if old else "" for old in sec.measures]

        new_sec = Section(
            number=sec.number,
            element=sec.element,
            thickness_nom=new_thick,
            diameter=new_diam,
            measures=new_measures,
            tmin=m,
            row=sec.row,
        )
        new_sections.append(new_sec)

        # Изменение для предпросмотра: сжатое «было → станет».
        before = (f"Тисп {sec.thickness_nom} · {sec.diameter} · "
                  f"замеры {'/'.join(sec.measures)}")
        after = (f"Тисп {new_thick} · {new_diam} · "
                 f"замеры {'/'.join(new_measures)}")
        changes.append(Change(
            target="Импорт",
            where=f"сечение {sec.number}",
            before=before,
            after=after,
        ))

    return new_sections, changes


def corrosion_rate_for_replaced(min_rate: float) -> float:
    """
    Скорость коррозии заменённого участка: утонения нет → минимальный порог.
    (Остаточный ресурс = (Sф − Sотбр) / rate; срок эксплуатации t не меняется.)
    """
    return min_rate
