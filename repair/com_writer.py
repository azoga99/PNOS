# -*- coding: utf-8 -*-
"""
Запись результатов ремонта в ПНОС_ТТП.xlsm через Excel COM.

Почему COM, а не openpyxl: книга — .xlsm с макросами и сложной структурой;
COM гарантированно сохраняет всё как есть (правило ПНОС). Пишутся только
значения ячеек — формулы не трогаются.

Листы «Импорт» и «Импорт (ПР)» зеркальны (подтверждено пользователем) —
операции применяются к обоим одинаково.

Функция самодостаточна (CoInitialize внутри) — можно звать из скрипта или из
QThread-воркера; в воркере поверх неё добавится DialogKiller, как в остальных
этапах ПНОС.
"""

from __future__ import annotations

import os

from repair.imports import apply_ops_com, WriteOp, SHEET_IMPORT, SHEET_IMPORT_PR
from repair.model import RepairRecord
from repair.registry import write_registry_com


def apply_to_workbook(
    xlsm_path: str,
    ops: list[WriteOp],
    records: list[RepairRecord] | None = None,
    log=None,
) -> dict:
    """
    Открыть книгу, применить операции к «Импорт» и «Импорт (ПР)», записать
    реестр на лист «Ремонт» (если передан), сохранить, закрыть.

    Возвращает счётчики: {"import": N, "import_pr": N, "registry": M}.
    Исключение = запись не удалась (вызывающая сторона делает откат из бэкапа).
    """
    import pythoncom
    import win32com.client

    def _log(msg: str):
        if log:
            log(msg)

    pythoncom.CoInitialize()
    xl = None
    wb = None
    try:
        _log("Запуск Excel в фоне...")
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AskToUpdateLinks = False
        xl.EnableEvents = False       # макросы книги не должны стрелять при записи
        xl.ScreenUpdating = False

        path = os.path.normpath(xlsm_path)
        _log(f"Открытие книги: {os.path.basename(path)}")
        wb = xl.Workbooks.Open(path, UpdateLinks=0,
                               IgnoreReadOnlyRecommended=True, Notify=False)

        counts = {"import": 0, "import_pr": 0, "registry": 0}

        counts["import"] = apply_ops_com(wb.Sheets(SHEET_IMPORT), ops)
        _log(f"Лист «{SHEET_IMPORT}»: записано ячеек {counts['import']}")

        # Зеркальный лист: те же строки и колонки.
        counts["import_pr"] = apply_ops_com(wb.Sheets(SHEET_IMPORT_PR), ops)
        _log(f"Лист «{SHEET_IMPORT_PR}»: записано ячеек {counts['import_pr']}")

        if records is not None:
            write_registry_com(wb, records)
            counts["registry"] = len(records)
            _log(f"Лист «Ремонт»: записей {counts['registry']}")

        wb.Save()
        _log("Книга сохранена.")
        return counts

    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if xl is not None:
            try:
                xl.ScreenUpdating = True
                xl.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
