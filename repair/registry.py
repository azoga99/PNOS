# -*- coding: utf-8 -*-
"""
Реестр ремонтов — видимый лист «Ремонт» в ПНОС_ТТП.xlsm.

Единственное хранилище модуля (принцип П-1: файлы остаются обычными).
Лист можно заполнить и поправить руками без программы — модуль его подхватит;
и наоборот, всё, что модуль записал, человек видит и правит в Excel.

Раскладка (строка 1 — шапка, данные со строки 2; одна строка = один комплект
ИД или акт без ИД):

    A  № записи
    B  Статус ИД: «есть» / «ожидается» / «не будет»
    C  Дата выполнения работ («Сдал» из реестра ИД), ДД.ММ.ГГГГ
    D  Номера актов через «;»
    E  Даты актов (гриф УТВЕРЖДАЮ) через «;»
    F  Замена (типоразмеры из таблиц актов)
    G  Материал замены
    H  Файлы-источники через «;»
    I  Примечание сотрудника

Чтение — openpyxl (без Excel, безопасно, тестируемо). Запись — через COM
в открытую книгу (макросы .xlsm сохраняются гарантированно); чистая
сериализация вынесена отдельно и покрыта тестами без Excel.
"""

from __future__ import annotations

from repair.model import (
    ActRef, RepairRecord, ID_PRESENT, ID_EXPECTED, ID_NEVER,
)

SHEET_NAME = "Ремонт"
LIST_SEP = ";"

HEADER = [
    "№", "Статус ИД", "Дата выполнения", "Номера актов", "Даты актов",
    "Замена", "Материал", "Файлы-источники", "Примечание",
]

_VALID_STATUS = {ID_PRESENT, ID_EXPECTED, ID_NEVER}


# ─── Чистая сериализация (тестируется без Excel) ──────────────────────

def _split(cell: str) -> list[str]:
    return [p.strip() for p in str(cell or "").split(LIST_SEP) if p.strip()]


def record_to_row(record: RepairRecord, n: int) -> list[str]:
    """RepairRecord → значения ячеек A…I (все — строки)."""
    return [
        str(n),
        record.id_status,
        record.date_done,
        LIST_SEP.join(a.number for a in record.acts if a.is_valid()),
        LIST_SEP.join(a.date for a in record.acts if a.is_valid()),
        record.replacements,
        record.material,
        LIST_SEP.join(record.source_files),
        record.note,
    ]


def row_to_record(row: list) -> RepairRecord | None:
    """
    Значения ячеек A…I → RepairRecord.
    Пустая строка (нет ни статуса, ни актов) → None.
    Неизвестный статус трактуется как «ожидается» — данные не теряем.
    """
    vals = [("" if v is None else str(v).strip()) for v in row]
    vals += [""] * (len(HEADER) - len(vals))

    status = vals[1] if vals[1] in _VALID_STATUS else (ID_EXPECTED if vals[1] else "")
    numbers = _split(vals[3])
    dates = _split(vals[4])

    if not status and not numbers:
        return None

    acts = [
        ActRef(number=num, date=(dates[i] if i < len(dates) else ""))
        for i, num in enumerate(numbers)
    ]
    return RepairRecord(
        id_status=status or ID_EXPECTED,
        date_done=vals[2],
        acts=acts,
        replacements=vals[5],
        material=vals[6],
        source_files=_split(vals[7]),
        note=vals[8],
    )


def records_to_rows(records: list[RepairRecord]) -> list[list[str]]:
    return [record_to_row(r, i) for i, r in enumerate(records, start=1)]


def rows_to_records(rows: list[list]) -> list[RepairRecord]:
    out = []
    for row in rows:
        rec = row_to_record(row)
        if rec is not None:
            out.append(rec)
    return out


# ─── Чтение из .xlsm (openpyxl, без Excel) ────────────────────────────

def read_registry(xlsm_path: str) -> list[RepairRecord]:
    """
    Прочитать лист «Ремонт». Листа нет → пустой список (ремонт не заводили).
    """
    import openpyxl  # локальный импорт — не тянуть при старте приложения

    wb = openpyxl.load_workbook(xlsm_path, read_only=True, data_only=True)
    try:
        if SHEET_NAME not in wb.sheetnames:
            return []
        ws = wb[SHEET_NAME]
        rows = []
        for row in ws.iter_rows(min_row=2, max_col=len(HEADER), values_only=True):
            rows.append(list(row))
        return rows_to_records(rows)
    finally:
        wb.close()


# ─── Запись через COM (книга уже открыта воркером) ────────────────────

def write_registry_com(workbook, records: list[RepairRecord]) -> None:
    """
    Записать реестр на лист «Ремонт» открытой COM-книги (создать при
    отсутствии). Пишутся только значения ячеек; формул на листе нет.
    Сохранение книги — на вызывающей стороне (воркер решает, когда Save).
    """
    ws = None
    for sh in workbook.Sheets:
        if str(sh.Name).strip().lower() == SHEET_NAME.lower():
            ws = sh
            break
    if ws is None:
        ws = workbook.Sheets.Add(After=workbook.Sheets(workbook.Sheets.Count))
        ws.Name = SHEET_NAME

    # Шапка.
    for c, title in enumerate(HEADER, start=1):
        ws.Cells(1, c).Value = title

    rows = records_to_rows(records)
    # Очистить прежние данные с запасом (старых строк могло быть больше).
    clear_to = max(len(rows), 200) + 1
    ws.Range(ws.Cells(2, 1), ws.Cells(clear_to, len(HEADER))).ClearContents()

    for r, row in enumerate(rows, start=2):
        for c, val in enumerate(row, start=1):
            cell = ws.Cells(r, c)
            cell.NumberFormat = "@"   # текст: даты не должны «поплыть»
            cell.Value = val
