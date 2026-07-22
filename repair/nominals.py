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
) -> tuple[list[Section], list[Change]]:
    """
    sections — все сечения листа «Импорт» (исходные).
    replaced — {№ сечения: новый номинал толщины, мм} из отметки сотрудника.
    perehod  — {№ сечения: (толщина_строкой, диаметр_строкой)} для переходов
               (двусторонние значения, напр. ("12 / 10", "273 / 219")).

    Возвращает (новые сечения, список изменений).
    """
    perehod = perehod or {}
    new_sections: list[Section] = []
    changes: list[Change] = []

    for sec in sections:
        if sec.number not in replaced:
            new_sections.append(sec)
            continue

        nominal = replaced[sec.number]

        if sec.number in perehod:
            new_thick, new_diam = perehod[sec.number]
        else:
            new_thick, new_diam = _fmt_thickness(nominal), sec.diameter

        m = _fmt_measure(nominal)
        new_measures = [m] * len(sec.measures)

        new_sec = Section(
            number=sec.number,
            element=sec.element,
            thickness_nom=new_thick,
            diameter=new_diam,
            measures=new_measures,
            tmin=m,
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
