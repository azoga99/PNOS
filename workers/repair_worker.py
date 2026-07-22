# -*- coding: utf-8 -*-
"""
Воркеры страницы «Ремонт» (модуль repair, ветка repair-module).

RepairScanWorker  — фоновое сканирование пункта: детекция документов ремонта,
                    привязка по именам файлов, чтение реестра (лист «Ремонт»)
                    и сечений листа «Импорт» через openpyxl. Без COM.
RepairApplyWorker — фоновое применение: бэкап → COM-запись в книгу
                    («Импорт», «Импорт (ПР)», «Ремонт») → строка «Сведения
                    о ремонте» в отчёт Word → журнал; при ошибке — автооткат
                    (см. repair.orchestrator).

Тяжёлые импорты (openpyxl, docx, win32com) — только внутри run(), чтобы не
ломать быстрый старт приложения.
"""

from __future__ import annotations

import os

from PySide6.QtCore import QThread, Signal


def find_point_report(point_folder: str) -> str | None:
    """
    Отчёт ЗЭПБ в корне пункта: .docx, не шаблон ПНОС_ТТП, предпочтительно
    без «ПЕЧАТЬ» (как в find_report_file этапа 5, без импорта воркера).
    """
    cands = []
    for f in os.listdir(point_folder):
        p = os.path.join(point_folder, f)
        low = f.lower()
        if (os.path.isfile(p) and low.endswith(".docx")
                and not f.startswith("~$") and "пнос_ттп" not in low):
            cands.append(p)
    if not cands:
        return None
    non_print = [p for p in cands if "печать" not in os.path.basename(p).lower()]
    return sorted(non_print or cands)[0]


def find_point_xlsm(point_folder: str) -> str | None:
    for f in sorted(os.listdir(point_folder)):
        p = os.path.join(point_folder, f)
        if os.path.isfile(p) and f.lower().endswith((".xlsm", ".xlsx", ".xls")) \
                and not f.startswith("~$"):
            return p
    return None


class RepairScanWorker(QThread):
    """Сканирование пункта для страницы «Ремонт» (только чтение)."""

    log = Signal(str)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    # результат: dict с ключами docs, verdicts, records, sections, unit, line,
    # xlsm, report, svedenia_now
    result = Signal(dict)

    def __init__(self, point_folder: str, parent=None):
        super().__init__(parent)
        self.point_folder = point_folder

    def run(self):
        try:
            from repair.detect import (find_repair_docs, parse_doc_filename,
                                       match_point, build_records_from_docs)
            from repair.registry import read_registry
            from repair.imports import read_import_sections
            from repair.word_writer import read_svedenia

            folder = self.point_folder
            out: dict = {"folder": folder}

            docs = find_repair_docs(folder)
            out["docs"] = docs
            if not docs.has_repair:
                self.info.emit("Документов ремонта не найдено — состояние С0", "info")
                self.result.emit(out)
                self.finished_ok.emit(True)
                return

            xlsm = find_point_xlsm(folder)
            report = find_point_report(folder)
            out["xlsm"], out["report"] = xlsm, report

            # Карточка пункта для светофора: установка (паспорт!B12),
            # № трубопровода (паспорт!B17). Пусто — вердикт unknown.
            unit = line = ""
            if xlsm:
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
                    ws = wb["паспорт"]
                    unit = str(ws["B12"].value or "").strip()
                    line = str(ws["B17"].value or "").strip()
                    wb.close()
                except Exception as e:
                    self.log.emit(f"паспорт не прочитан: {e}")
            out["unit"], out["line"] = unit, line

            verdicts = []
            for f in docs.all_files:
                p = parse_doc_filename(f)
                verdicts.append((f, p, match_point(p, unit, line)))
            out["verdicts"] = verdicts

            # Реестр: из листа «Ремонт», если он уже есть; иначе черновик
            # из имён документов.
            records = read_registry(xlsm) if xlsm else []
            out["registry_existed"] = bool(records)
            if not records:
                records = build_records_from_docs(docs)
            out["records"] = records

            out["sections"] = read_import_sections(xlsm) if xlsm else []
            out["svedenia_now"] = read_svedenia(report) if report else None

            n_ok = sum(1 for _, _, v in verdicts if v == "ok")
            self.info.emit(
                f"Документов: {len(verdicts)} (совпало {n_ok}) · "
                f"сечений: {len(out['sections'])}", "done")
            self.result.emit(out)
            self.finished_ok.emit(True)

        except Exception as e:
            import traceback
            self.log.emit(traceback.format_exc())
            self.info.emit(f"Ошибка сканирования: {str(e)[:80]}", "error")
            self.finished_ok.emit(False)


class RepairVerifyWorker(QThread):
    """Верификация пункта (этап Э2): только чтение, без COM."""

    log = Signal(str)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    result = Signal(list)            # list[Finding]

    def __init__(self, point_folder: str, parent=None):
        super().__init__(parent)
        self.point_folder = point_folder

    def run(self):
        try:
            from repair.verifier import verify_point
            findings = verify_point(self.point_folder)
            self.result.emit(findings)
            n_err = sum(1 for f in findings if f.level == "error")
            n_warn = sum(1 for f in findings if f.level == "warn")
            if n_err:
                self.info.emit(f"Верификация: ошибок {n_err}, "
                               f"предупреждений {n_warn}", "error")
            elif n_warn:
                self.info.emit(f"Верификация: предупреждений {n_warn}", "warn")
            else:
                self.info.emit("Верификация: пункт согласован", "done")
            self.finished_ok.emit(True)
        except Exception as e:
            import traceback
            self.log.emit(traceback.format_exc())
            self.info.emit(f"Ошибка верификации: {str(e)[:80]}", "error")
            self.finished_ok.emit(False)


class RepairApplyWorker(QThread):
    """Применение согласованных изменений (COM Excel + python-docx Word)."""

    log = Signal(str)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    result = Signal(object)          # ApplyResult

    def __init__(self, point_folder: str, xlsm_path: str,
                 report_path: str | None, ops: list, records: list,
                 svedenia_text: str | None, parent=None):
        super().__init__(parent)
        self.point_folder = point_folder
        self.xlsm_path = xlsm_path
        self.report_path = report_path
        self.ops = ops
        self.records = records
        self.svedenia_text = svedenia_text

    def run(self):
        try:
            from repair.orchestrator import apply_repair
            from repair.com_writer import apply_to_workbook
            from repair.word_writer import write_svedenia

            report = self.report_path
            sved = self.svedenia_text

            def writer(xlsm, ops, records, log=None):
                counts = apply_to_workbook(xlsm, ops, records, log=log)
                if report and sved is not None:
                    if write_svedenia(report, sved):
                        counts["word"] = 1
                        if log:
                            log("Отчёт Word: строка «Сведения о ремонте» записана")
                    else:
                        raise RuntimeError(
                            "в отчёте не найдена строка «Сведения о ремонте»")
                return counts

            extra = [report] if report else []
            res = apply_repair(self.point_folder, self.xlsm_path, self.ops,
                               self.records, writer, extra_files=extra,
                               log=self.log.emit)
            self.result.emit(res)
            if res.ok:
                self.info.emit("Применено. Бэкап и журнал созданы.", "done")
            else:
                self.info.emit(f"Ошибка, выполнен откат: {res.error[:80]}", "error")
            self.finished_ok.emit(res.ok)

        except Exception as e:
            import traceback
            self.log.emit(traceback.format_exc())
            self.info.emit(f"Ошибка применения: {str(e)[:80]}", "error")
            self.finished_ok.emit(False)
