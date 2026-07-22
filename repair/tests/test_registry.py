# -*- coding: utf-8 -*-
"""
Тесты реестра ремонтов (лист «Ремонт»).

Сериализация — чистая, гоняется без Excel. Чтение — через openpyxl на
временном файле. Запись COM — на фейковом объекте книги (проверяем, какие
значения ушли бы в ячейки).
"""

import openpyxl

from repair.model import ActRef, RepairRecord, ID_PRESENT, ID_EXPECTED
from repair.registry import (
    SHEET_NAME, HEADER,
    record_to_row, records_to_rows, rows_to_records,
    read_registry, write_registry_com,
)


def _sample_records():
    return [
        RepairRecord(
            id_status=ID_PRESENT, date_done="13.04.2026",
            acts=[ActRef("222-1-84", "24.02.2026")],
            replacements="273x10; 219x8; 325x12",
            material="Ст.20",
            source_files=["36-30, №6, 222-1-84.pdf"],
        ),
        RepairRecord(
            id_status=ID_EXPECTED,
            acts=[ActRef("222-1-121", "24.03.2026"), ActRef("222-1-143", "30.03.2026")],
            note="ИД обещали к маю",
        ),
    ]


# ─── Сериализация туда-обратно ────────────────────────────────────────

def test_roundtrip():
    records = _sample_records()
    rows = records_to_rows(records)
    back = rows_to_records(rows)

    assert len(back) == 2
    a, b = back
    assert a.id_status == ID_PRESENT
    assert a.date_done == "13.04.2026"
    assert [x.number for x in a.acts] == ["222-1-84"]
    assert [x.date for x in a.acts] == ["24.02.2026"]
    assert a.replacements == "273x10; 219x8; 325x12"
    assert a.material == "Ст.20"
    assert a.source_files == ["36-30, №6, 222-1-84.pdf"]

    assert b.id_status == ID_EXPECTED
    assert not b.has_id
    assert [x.number for x in b.acts] == ["222-1-121", "222-1-143"]
    assert [x.date for x in b.acts] == ["24.03.2026", "30.03.2026"]
    assert b.note == "ИД обещали к маю"


def test_row_numbering():
    rows = records_to_rows(_sample_records())
    assert rows[0][0] == "1"
    assert rows[1][0] == "2"


def test_empty_and_partial_rows():
    assert rows_to_records([["", "", "", "", "", "", "", "", ""]]) == []
    assert rows_to_records([[None] * 9]) == []
    # строка, набитая руками не полностью: только номер акта
    recs = rows_to_records([["1", "", "", "222-1-5", "", "", "", "", ""]])
    assert len(recs) == 1
    assert recs[0].id_status == ID_EXPECTED       # разумный дефолт
    assert recs[0].acts[0].number == "222-1-5"
    assert recs[0].acts[0].date == ""


def test_unknown_status_becomes_expected():
    """Опечатка в статусе не роняет чтение и не теряет строку."""
    recs = rows_to_records([["1", "ждём", "", "222-1-5", "01.02.2026",
                             "", "", "", ""]])
    assert len(recs) == 1
    assert recs[0].id_status == ID_EXPECTED


def test_dates_align_with_numbers():
    """Дат меньше, чем номеров — недостающие пустые, ничего не съезжает."""
    recs = rows_to_records([["1", "ожидается", "", "222-1-1;222-1-2",
                             "01.01.2026", "", "", "", ""]])
    acts = recs[0].acts
    assert [a.number for a in acts] == ["222-1-1", "222-1-2"]
    assert [a.date for a in acts] == ["01.01.2026", ""]


# ─── Чтение из файла (openpyxl) ───────────────────────────────────────

def _write_xlsx(path, with_sheet=True, rows=()):
    wb = openpyxl.Workbook()
    if with_sheet:
        ws = wb.create_sheet(SHEET_NAME)
        ws.append(HEADER)
        for r in rows:
            ws.append(r)
    wb.save(path)


def test_read_registry_roundtrip(tmp_path):
    path = str(tmp_path / "т.xlsx")
    _write_xlsx(path, rows=records_to_rows(_sample_records()))
    recs = read_registry(path)
    assert len(recs) == 2
    assert recs[0].acts[0].number == "222-1-84"
    assert recs[1].note == "ИД обещали к маю"


def test_read_registry_no_sheet(tmp_path):
    path = str(tmp_path / "т.xlsx")
    _write_xlsx(path, with_sheet=False)
    assert read_registry(path) == []


# ─── Запись через COM (фейковая книга) ────────────────────────────────

class _FakeCell:
    def __init__(self):
        self.Value = None
        self.NumberFormat = None


class _FakeRange:
    def __init__(self):
        self.cleared = False
    def ClearContents(self):
        self.cleared = True


class _FakeSheet:
    def __init__(self, name):
        self.Name = name
        self._cells = {}
        self.cleared_ranges = []
    def Cells(self, r, c):
        key = (r, c)
        if key not in self._cells:
            self._cells[key] = _FakeCell()
        return self._cells[key]
    def Range(self, c1, c2):
        rng = _FakeRange()
        self.cleared_ranges.append(rng)
        return rng
    def value_at(self, r, c):
        cell = self._cells.get((r, c))
        return None if cell is None else cell.Value


class _FakeSheets(list):
    def __call__(self, index):
        return self[index - 1]
    @property
    def Count(self):
        return len(self)
    def Add(self, After=None):
        sh = _FakeSheet("Лист")
        self.append(sh)
        return sh


class _FakeWorkbook:
    def __init__(self, sheet_names):
        self.Sheets = _FakeSheets([_FakeSheet(n) for n in sheet_names])


def test_write_creates_sheet_and_header():
    wb = _FakeWorkbook(["паспорт", "Импорт"])
    write_registry_com(wb, _sample_records())

    sheet = next(s for s in wb.Sheets if s.Name == SHEET_NAME)
    assert [sheet.value_at(1, c) for c in range(1, len(HEADER) + 1)] == HEADER
    # первая запись: статус и номер акта на местах
    assert sheet.value_at(2, 2) == "есть"
    assert sheet.value_at(2, 4) == "222-1-84"
    # даты пишутся текстом
    assert sheet.Cells(2, 3).NumberFormat == "@"
    # старые данные зачищены
    assert sheet.cleared_ranges and all(r.cleared for r in sheet.cleared_ranges)


def test_write_reuses_existing_sheet():
    wb = _FakeWorkbook(["паспорт", SHEET_NAME])
    write_registry_com(wb, _sample_records())
    assert sum(1 for s in wb.Sheets if s.Name == SHEET_NAME) == 1
