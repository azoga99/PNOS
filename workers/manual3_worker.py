# -*- coding: utf-8 -*-
"""
Шаг 3: Ручной контроль — запись введённых в форму данных в Excel.

Пользователь заполняет поля прямо в интерфейсе, а этот воркер невидимо
открывает ПНОС_ТТП.xlsm каждого пункта и записывает значения на лист
«паспорт» в заданные ячейки (см. CONFIG["MANUAL_FORM"]).

Даты пишутся как ТЕКСТ (NumberFormat = "@"), чтобы Excel не переформатировал
их по-своему — в ячейке остаётся ровно введённая строка ДД.ММ.ГГГГ.
"""

import os

from PySide6.QtCore import QThread, Signal

import pythoncom
import win32com.client

from config import CONFIG
from workers.stage2_worker import find_file_exact_ext
from workers.stage3_worker import DialogKiller, safe_close_com
from status_manager import set_stage_status


class Manual3Worker(QThread):
    """Записывает данные формы ручного контроля в ПНОС_ТТП.xlsm по пунктам."""

    log = Signal(str)
    progress = Signal(int)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    report_ready = Signal(list)

    def __init__(self, entries: list, parent=None):
        """
        entries — список словарей вида:
            {"folder": <путь к папке пункта>, "num": "2758", "values": {...}}
        values содержит ключи: purpose, commission_year, specialist, date,
                               epb_done, epb_org, reg_num, epb_date, sheets
        """
        super().__init__(parent)
        self.entries = entries or []
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        cfg = CONFIG["MANUAL_FORM"]
        sheet_name = cfg["SHEET"]

        self.log.emit("═" * 40)
        self.log.emit("ШАГ 3: Запись данных ручного контроля в Excel")
        self.log.emit("═" * 40)

        report = []
        if not self.entries:
            self.log.emit("⚠ Нет пунктов для записи.")
            self.finished_ok.emit(False)
            return

        pythoncom.CoInitialize()
        dk = DialogKiller(log_callback=self.log.emit)
        xl_app = None
        ok_count = 0
        err_count = 0

        try:
            dk.start()
            self.log.emit("\n⚙ Инициализация Excel в фоне...")
            xl_app = win32com.client.Dispatch("Excel.Application")
            xl_app.Visible = False
            xl_app.DisplayAlerts = False
            xl_app.AskToUpdateLinks = False
            xl_app.ScreenUpdating = False
            xl_app.EnableEvents = False
            self.info.emit("Excel готов к записи данных", "done")
            self.progress.emit(5)

            total = len(self.entries)
            for i, entry in enumerate(self.entries):
                if self._is_cancelled:
                    self.log.emit("\n❌ Отменено пользователем.")
                    break

                folder = entry["folder"]
                num = entry.get("num", os.path.basename(folder))
                values = entry.get("values", {})
                row = {"№ Пункта": num, "Записано": False, "Результат": "Ожидание"}
                self.log.emit(f"\n── п.{num} ──")
                self.info.emit(f"Запись данных п.{num}...", "wait")

                work_wb = None
                try:
                    work_path = find_file_exact_ext(folder, (".xls", ".xlsx", ".xlsm"))
                    if not work_path:
                        self.log.emit("   ✗ Excel-файл (ПНОС_ТТП.xlsm) не найден")
                        row["Результат"] = "❌ Нет Excel"
                        err_count += 1
                        report.append(row)
                        continue

                    work_wb = xl_app.Workbooks.Open(
                        work_path, UpdateLinks=0,
                        IgnoreReadOnlyRecommended=True, Notify=False)
                    ws = work_wb.Sheets(sheet_name)

                    self._write_values(ws, cfg, values, folder)

                    work_wb.Save()
                    safe_close_com(work_wb, save=False)
                    work_wb = None

                    set_stage_status(folder, "manual3", True)
                    # Сохраняем черновик, чтобы при повторном открытии форма показала введённое
                    set_stage_status(folder, "manual3_draft", values)

                    self.log.emit("   ✓ Данные записаны в лист «паспорт»")
                    self.info.emit(f"п.{num} — данные записаны", "done")
                    row["Записано"] = True
                    row["Результат"] = "✅ Успешно"
                    ok_count += 1

                except Exception as ex:
                    self.log.emit(f"   ✗ Ошибка записи: {ex}")
                    row["Результат"] = f"❌ {str(ex)[:40]}"
                    err_count += 1
                    if work_wb is not None:
                        safe_close_com(work_wb, save=False)
                finally:
                    report.append(row)

                self.progress.emit(5 + int((i + 1) / total * 90))

            self.progress.emit(100)
            self.log.emit(f"\n{'═' * 40}")
            self.log.emit(f"✅ Готово. Записано: {ok_count}, ошибок: {err_count}")
            self.report_ready.emit(report)
            self.finished_ok.emit(err_count == 0 and not self._is_cancelled)

        except Exception as e:
            import traceback
            self.log.emit(f"\n❌ Критическая ошибка: {e}")
            self.log.emit(traceback.format_exc())
            self.report_ready.emit(report)
            self.finished_ok.emit(False)
        finally:
            try:
                if xl_app is not None:
                    xl_app.Quit()
            except Exception:
                pass
            dk.stop()
            pythoncom.CoUninitialize()

    def _write_values(self, ws, cfg, values: dict, folder: str = ""):
        """Записывает значения формы в ячейки листа «паспорт»."""
        def put_text(cell_addr, text):
            rng = ws.Range(cell_addr)
            rng.NumberFormat = "@"   # текстовый формат — Excel не переформатирует
            rng.Value = "" if text is None else str(text)

        # B14 — предназначение (обязательное)
        put_text(cfg["PURPOSE_CELL"], values.get("purpose", ""))

        # B19 — год ввода в эксплуатацию (необязательное). В ячейке лежит полная
        # дата, отчёт берёт из неё только год (данные!C12 = YEAR(паспорт!B19)),
        # поэтому день и месяц сохраняем: сначала из текущего значения ячейки,
        # затем из F051 паспорта, и лишь если взять неоткуда — 01.01.
        # Пустое поле ячейку не трогает: этап 4 подставит значение из паспорта.
        year = str(values.get("commission_year", "") or "").strip()
        cell_addr = cfg.get("COMMISSION_CELL")
        if year and cell_addr:
            import passport_data
            current = ws.Range(cell_addr).Value
            date_text = passport_data.commission_value(
                year, current, passport_data.commission_date(folder))
            if date_text:
                put_text(cell_addr, date_text)
                self.log.emit(f"   ✓ Год ввода в эксплуатацию: {date_text}")

        # B22 — ФИО специалиста (обязательное)
        put_text(cfg["SPECIALIST_CELL"], values.get("specialist", ""))

        # B33 — дата (обязательная), как текст ДД.ММ.ГГГГ
        put_text(cfg["DATE_CELL"], values.get("date", ""))

        # B52 — организация прошлой ЭПБ (необязательное)
        org = (values.get("epb_org") or "").strip()
        put_text(cfg["ORG_CELL"], org)

        # B51 — да/нет: авто по наличию организации в B52
        put_text(cfg["EPB_DONE_CELL"], "да" if org else "нет")

        # B53 — рег. номер (необязательное), как есть
        put_text(cfg["REG_NUM_CELL"], values.get("reg_num", ""))

        # B54 — дата прошлой ЭПБ (необязательная), текст ДД.ММ.ГГГГ
        put_text(cfg["EPB_DATE_CELL"], values.get("epb_date", ""))

        # B55 — количество листов (необязательное): число, если задано
        sheets_raw = str(values.get("sheets", "") or "").strip()
        cell = ws.Range(cfg["SHEETS_CELL"])
        if not sheets_raw:
            cell.Value = ""
        else:
            try:
                cell.NumberFormat = "0"
                cell.Value = int(sheets_raw)
            except ValueError:
                cell.NumberFormat = "@"
                cell.Value = sheets_raw
