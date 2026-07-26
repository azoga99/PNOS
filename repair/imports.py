# -*- coding: utf-8 -*-
"""
Слой листа «Импорт» ПНОС_ТТП.xlsm: чтение сечений и запись номиналов.

Раскладка листа (строка 1 — шапка, данные со строки 2):
    A Номер · B Сечение · C Элемент · D Тисп · E Диам ·
    F…K Т1…Т6 · L Тмин · далее — материал и служебные колонки (не трогаем).

Важно (реальные данные п.2701): номера сечений в «Импорте» НЕ уникальны —
задвижка №6 и отвод №6 живут в соседних строках. Поэтому адресация всегда
по строке листа (Section.row), а подбор строк по номерам сечений отдаёт
предупреждения о коллизиях.

Принцип записи: пишем только значения в ячейки данных, формулы и правый
формульный блок листа не трогаем — никаких #REF!. Диф считается чистой
функцией (тестируется без Excel), COM только применяет готовые операции.
"""

from __future__ import annotations

from dataclasses import dataclass

from repair.config import REPAIR_CONFIG
from repair.model import Section

SHEET_IMPORT = "Импорт"
SHEET_IMPORT_PR = "Импорт (ПР)"

COL_SECTION = 2   # B
COL_ELEMENT = 3   # C
COL_THICK = 4     # D  Тисп
COL_DIAM = 5      # E
COL_T1 = 6        # F  (…K = Т6)
N_MEASURES = 6
COL_TMIN = 12     # L


@dataclass
class WriteOp:
    """Одна операция записи в лист: строка, колонка, значение."""
    row: int
    col: int
    value: object


def _cell_str(v) -> str:
    """Значение ячейки → строка в стиле отчёта («11.2», «12», «12 / 10»)."""
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        # 10.0 → «10.0»: замеры в отчёте всегда с одним знаком
        return f"{v:.1f}"
    return str(v).strip()


def _to_excel_value(s: str) -> str:
    """
    Строка → значение для Excel. Всегда ТЕКСТ.

    Замеры на листе «Импорт» лежат текстом с точкой («6.9»), формат ячеек «@»,
    а блок отчёта (колонки AH…AS) переносит значение как есть: `=IF(F115="","",
    F115)`. Число 6.0 в такой ячейке печатается в отчёте как «6» — рядом с
    «6.9» соседних строк это читается как ошибка замера. Расчётные листы
    («Табл», «ТаблТв», «Табл2») берут значение через VALUE() и текст понимают.
    """
    return (s or "").strip()


# ─── Чтение (openpyxl, без Excel) ─────────────────────────────────────

def read_import_sections(xlsm_path: str, sheet: str = SHEET_IMPORT) -> list[Section]:
    """
    Прочитать сечения листа «Импорт». Строки без числового «Сечения» (B)
    пропускаются; чтение до конца данных.
    """
    import openpyxl  # локальный импорт — не тянуть при старте приложения

    wb = openpyxl.load_workbook(xlsm_path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        out: list[Section] = []
        for r, row in enumerate(
                ws.iter_rows(min_row=2, max_col=COL_TMIN, values_only=True),
                start=2):
            vals = list(row) + [None] * (COL_TMIN - len(row))
            b = vals[COL_SECTION - 1]
            try:
                number = int(float(str(b).strip()))
            except (TypeError, ValueError):
                continue
            measures = [
                _cell_str(vals[COL_T1 - 1 + i]) for i in range(N_MEASURES)
            ]
            # хвостовые пустые точки отбрасываем, внутренние пустоты храним
            while measures and not measures[-1]:
                measures.pop()
            out.append(Section(
                number=number,
                element=_cell_str(vals[COL_ELEMENT - 1]),
                thickness_nom=_cell_str(vals[COL_THICK - 1]),
                diameter=_cell_str(vals[COL_DIAM - 1]),
                measures=measures,
                tmin=_cell_str(vals[COL_TMIN - 1]),
                row=r,
            ))
        return out
    finally:
        wb.close()


# ─── Подбор строк по отметке сотрудника ───────────────────────────────

def rows_for_numbers(
    sections: list[Section],
    numbers: dict[int, float],
    config: dict | None = None,
) -> tuple[dict[int, float], list[str]]:
    """
    {№ сечения: номинал} → {строка листа: номинал} + предупреждения.

    Коллизии (один номер — несколько строк) решаются в пользу «не мелких»
    элементов; мелкие (SMALL_ELEMENTS) и исключаемые (EXCLUDE_FROM_TABLE3,
    т.е. задвижки) пропускаются с предупреждением — решение за сотрудником.
    """
    cfg = config or REPAIR_CONFIG
    skip_types = {t.lower() for t in cfg["SMALL_ELEMENTS"]} | \
                 {t.lower() for t in cfg["EXCLUDE_FROM_TABLE3"]}

    by_number: dict[int, list[Section]] = {}
    for s in sections:
        by_number.setdefault(s.number, []).append(s)

    rows: dict[int, float] = {}
    warnings: list[str] = []

    for num, nominal in numbers.items():
        cands = by_number.get(num, [])
        if not cands:
            warnings.append(f"сечение {num}: не найдено на листе «Импорт»")
            continue
        usable = [s for s in cands if s.element.lower() not in skip_types]
        skipped = [s for s in cands if s.element.lower() in skip_types]
        for s in skipped:
            warnings.append(
                f"сечение {num}: строка {s.row} ({s.element}) пропущена — "
                f"мелкий/исключаемый элемент, при необходимости отметьте вручную")
        if not usable:
            continue
        if len(usable) > 1:
            warnings.append(
                f"сечение {num}: {len(usable)} строк на листе "
                f"({', '.join(str(s.row) for s in usable)}) — применено ко всем, проверьте")
        for s in usable:
            rows[s.row] = nominal
    return rows, warnings


# ─── Диф и запись ─────────────────────────────────────────────────────

def build_write_ops(old: list[Section], new: list[Section]) -> list[WriteOp]:
    """
    Сравнить старые и новые сечения (по строкам) и вернуть операции записи
    только для изменившихся ячеек: Тисп (D), Диам (E), Т1…Т6 (F…K), Тмин (L).
    Пустые точки замеров не трогаются.
    """
    old_by_row = {s.row: s for s in old if s.row}
    ops: list[WriteOp] = []

    for n in new:
        o = old_by_row.get(n.row)
        if o is None or o is n:
            continue
        if n.thickness_nom != o.thickness_nom:
            ops.append(WriteOp(n.row, COL_THICK, _to_excel_value(n.thickness_nom)))
        if n.diameter != o.diameter:
            ops.append(WriteOp(n.row, COL_DIAM, _to_excel_value(n.diameter)))
        for i in range(min(len(n.measures), N_MEASURES)):
            nv = n.measures[i]
            ov = o.measures[i] if i < len(o.measures) else ""
            if nv and nv != ov:
                ops.append(WriteOp(n.row, COL_T1 + i, _to_excel_value(nv)))
        if n.tmin and n.tmin != o.tmin:
            ops.append(WriteOp(n.row, COL_TMIN, _to_excel_value(n.tmin)))
    return ops


def apply_ops_com(worksheet, ops: list[WriteOp]) -> int:
    """
    Применить операции к листу открытой COM-книги. Возвращает число записей.
    Сохранение книги — на вызывающей стороне.

    Формат ячейки принудительно текстовый: без него Excel сам превращает
    строку «6.0» в число 6, и отчёт печатает замер без десятых.
    """
    for op in ops:
        cell = worksheet.Cells(op.row, op.col)
        cell.NumberFormat = "@"
        cell.Value = op.value
    return len(ops)
