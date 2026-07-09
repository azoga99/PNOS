# -*- coding: utf-8 -*-
"""
Правка готовых отчётов (страница «Правка отчётов»).

Пакетно вносит в каждый Word-файл отчёта, лежащий в корне папки пункта
(включая копии «ПЕЧАТЬ», без захода в подпапки), четыре правки:
1. Дата на титуле (поле MERGEFIELD «ДатаОтчета», формат «06 июля 2026 г.»).
2. Две строки квалификационных удостоверений Жиганьшина.
3. Ячейка «Положение о производственном контроле…» в Таблице 1.
4. Скан приказа в приложении «ПРИКАЗ О НАЗНАЧЕНИИ ГРУППЫ ЭКСПЕРТА» —
   заменяется картинкой из выбранной папки с сохранением размеров.

Режим автономный: не связан с конвейером, статусами .pnos_status.json
и страницами «Обзор» / «Итоговый отчёт».
"""

import os
import re

import pythoncom
import win32com.client

from PySide6.QtCore import QThread, Signal

from config import CONFIG
from utils import title_date_text
from workers.stage4_worker import safe_close_com

# Константы Word
WD_GOTO_PAGE = 1
WD_GOTO_ABSOLUTE = 1
WD_FIND_STOP = 0          # Wrap: не продолжать с начала документа
WD_WITH_IN_TABLE = 12     # Information(...): курсор внутри таблицы
WD_COLLAPSE_END = 0


def range_find(rng, text: str, wildcards: bool = False) -> bool:
    """Поиск в диапазоне через свойства Find.

    ВАЖНО: параметры Execute(FindText=..., MatchWildcards=...) в связке
    pywin32 + Word здесь молча игнорируются (проверено: wildcard-шаблоны
    не находились), поэтому все настройки задаются свойствами Find.
    При успехе rng сжимается до найденного текста.
    """
    f = rng.Find
    f.ClearFormatting()
    f.Text = text
    f.MatchWildcards = wildcards
    if not wildcards:
        f.MatchCase = False
        f.MatchWholeWord = False
    f.Forward = True
    f.Wrap = WD_FIND_STOP
    return bool(f.Execute())


class FixReportsWorker(QThread):
    """Пакетная правка готовых Word-отчётов через Word COM."""
    log = Signal(str)
    info = Signal(str, str)   # (сообщение, категория: info/done/wait/warn/error)
    progress = Signal(int)
    finished_ok = Signal(bool)

    def __init__(self, root_folder: str, date_str: str, scan_path: str, parent=None):
        super().__init__(parent)
        # normpath обязателен: QFileDialog отдаёт путь с прямыми слэшами
        # (C:/…), а Word в Documents.Open такой путь не принимает —
        # падает с «К сожалению, файл не найден…»
        self.root_folder = os.path.normpath(root_folder)
        self.date_str = date_str          # «06.07.2026»
        self.scan_path = os.path.normpath(scan_path)
        self._is_cancelled = False
        self.fix_cfg = CONFIG["FIX_REPORTS"]

    def cancel(self):
        self._is_cancelled = True

    # ─── Сбор файлов ───────────────────────────────────────────────

    def _collect_files(self):
        """[(метка пункта, полный путь к docx), ...].

        Берутся Word-файлы из корня выбранной папки (метка «корень»)
        и из корня папок пунктов п.* — без захода в подпапки."""
        skip_names = {n.lower() for n in CONFIG["TEMPLATE_FILES"]}

        def docx_in(folder):
            for f in sorted(os.listdir(folder)):
                if f.startswith("~$") or not f.lower().endswith(".docx"):
                    continue
                if f.lower() in skip_names:
                    continue  # шаблон ПНОС_ТТП.docx — не отчёт
                if os.path.isfile(os.path.join(folder, f)):
                    yield f

        result = [("корень", os.path.join(self.root_folder, f))
                  for f in docx_in(self.root_folder)]
        for d in sorted(os.listdir(self.root_folder)):
            p_folder = os.path.join(self.root_folder, d)
            if os.path.isdir(p_folder) and d.startswith("п."):
                result.extend((d, os.path.join(p_folder, f))
                              for f in docx_in(p_folder))
        return result

    # ─── Правка 1: дата на титуле ──────────────────────────────────

    def _fix_title_date(self, doc, new_text: str) -> str:
        field_name = self.fix_cfg["DATE_FIELD"]
        replaced = 0
        try:
            for fld in doc.Fields:
                try:
                    code = fld.Code.Text or ""
                except Exception:
                    continue
                # Имя поля сравниваем точно: есть соседние поля
                # «ДатаОтчетаДень1» / «ДатаОтчетаМес1» / «ДатаОтчетаГод»
                m = re.search(r'MERGEFIELD\s+"?([^"\s\\]+)', code)
                if m and m.group(1) == field_name:
                    fld.Result.Text = new_text
                    replaced += 1
        except Exception as e:
            self.log.emit(f"   [дата] ошибка перебора полей: {e}")
        if replaced:
            return f"✓ дата (поле ×{replaced})"

        # Фолбэк: поле разорвано — ищем дату словами на первой странице
        end_p1 = doc.Content.End
        try:
            page2 = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=2)
            if page2.Start > 0:
                end_p1 = page2.Start
        except Exception:
            pass
        rng = doc.Range(0, end_p1)
        if range_find(rng, "[0-9]{2} [а-я]@ [0-9]{4} г.", wildcards=True) \
                and rng.Start < end_p1:
            rng.Text = new_text
            return "✓ дата (текст)"
        return "⚠ дата не найдена"

    # ─── Правка 2: удостоверения ───────────────────────────────────

    def _replace_paragraph(self, doc, old_key: str, new_line: str, done_key: str) -> str:
        """Находит абзац по уникальному фрагменту и заменяет его текст целиком."""
        rng = doc.Content
        if range_find(rng, old_key):
            para = rng.Paragraphs(1).Range.Duplicate
            para.End = para.End - 1  # не трогаем знак абзаца
            para.Text = new_line
            return "✓"
        if range_find(doc.Content, done_key):
            return "— уже"
        return "⚠ не найдено"

    def _fix_certificates(self, doc) -> str:
        results = []
        for item in self.fix_cfg["CERT_FIXES"]:
            results.append(self._replace_paragraph(
                doc, item["OLD_KEY"], item["NEW_LINE"], item["DONE_KEY"]))
        if all(r == "✓" for r in results):
            return "✓ удостоверения"
        if all(r == "— уже" for r in results):
            return "— удостоверения уже заменены"
        return f"⚠ удостоверения: ЗС {results[0]}, ТУ {results[1]}"

    # ─── Правка 3: положение в Таблице 1 ───────────────────────────

    def _fix_table(self, doc) -> str:
        cfg = self.fix_cfg["TABLE_FIX"]
        rng = doc.Content
        if not range_find(rng, cfg["OLD_NUM"]):
            if range_find(doc.Content, cfg["NEW_NUM"]):
                return "— положение уже заменено"
            return "⚠ положение не найдено"
        if not rng.Information(WD_WITH_IN_TABLE):
            return "⚠ положение найдено вне таблицы"
        cell_rng = rng.Cells(1).Range
        cell_end = cell_rng.End

        cr = cell_rng.Duplicate
        if range_find(cr, cfg["OLD_NUM"]) and cr.Start < cell_end:
            cr.Text = cfg["NEW_NUM"]

        # Дата утверждения — последняя строка ячейки, «от ДД.ММ.ГГГГ г.»
        cr2 = cell_rng.Duplicate
        if range_find(cr2, "от [0-9]{2}.[0-9]{2}.[0-9]{4} г.", wildcards=True) \
                and cr2.Start < cell_end:
            cr2.Text = cfg["NEW_DATE"]
            return "✓ положение"
        return "⚠ положение: номер заменён, дата не найдена"

    # ─── Правка 4: скан приказа ────────────────────────────────────

    def _fix_order_scan(self, doc) -> str:
        heading = self.fix_cfg["ORDER_HEADING"]
        search = doc.Content
        # Первое вхождение обычно строка СОДЕРЖАНИЯ — берём то,
        # сразу после которого стоит картинка.
        while range_find(search, heading):
            head_end = search.End
            target = None
            target_start = None
            for shp in doc.InlineShapes:
                try:
                    s = shp.Range.Start
                except Exception:
                    continue
                if head_end <= s <= head_end + 300:
                    if target_start is None or s < target_start:
                        target, target_start = shp, s
            if target is not None:
                try:
                    w, h = target.Width, target.Height
                    anchor = target.Range.Duplicate
                    target.Delete()
                    new_shp = doc.InlineShapes.AddPicture(
                        FileName=self.scan_path, LinkToFile=False,
                        SaveWithDocument=True, Range=anchor)
                    new_shp.LockAspectRatio = 0
                    new_shp.Width = w
                    new_shp.Height = h
                    return "✓ приказ"
                except Exception as e:
                    self.log.emit(f"   [приказ] ошибка замены: {e}")
                    return "⚠ приказ: ошибка замены"
            search.Collapse(WD_COLLAPSE_END)
        return "⚠ приказ: страница с картинкой не найдена"

    # ─── Открытие документа ────────────────────────────────────────

    def _open_doc(self, wd_app, path: str):
        """Открывает документ; при отказе Word пробует короткое 8.3-имя
        (обходит длинные пути и спецсимволы)."""
        try:
            doc = wd_app.Documents.Open(
                os.path.normpath(path),
                ConfirmConversions=False, AddToRecentFiles=False)
            if doc is not None:
                return doc
        except Exception as e:
            self.log.emit(f"   [открытие] не открылся обычный путь: {e}")
        import win32api
        short = win32api.GetShortPathName(os.path.normpath(path))
        return wd_app.Documents.Open(
            short, ConfirmConversions=False, AddToRecentFiles=False)

    # ─── Основной цикл ─────────────────────────────────────────────

    def run(self):
        self.log.emit("═" * 40)
        self.log.emit("ПРАВКА ОТЧЁТОВ")
        self.log.emit("═" * 40)

        try:
            new_date = title_date_text(self.date_str)
        except Exception:
            self.info.emit(f"Неверный формат даты: {self.date_str}", "error")
            self.finished_ok.emit(False)
            return

        if not os.path.isdir(self.root_folder):
            self.info.emit("Выбранная папка не найдена", "error")
            self.finished_ok.emit(False)
            return
        if not os.path.isfile(self.scan_path):
            self.info.emit("Файл скана приказа не найден", "error")
            self.finished_ok.emit(False)
            return

        files = self._collect_files()
        if not files:
            self.info.emit("В папках пунктов не найдено ни одного Word-файла", "error")
            self.finished_ok.emit(False)
            return

        self.info.emit(f"Найдено файлов: {len(files)}. Дата на титуле: «{new_date}». "
                       f"Скан: {os.path.basename(self.scan_path)}", "info")

        pythoncom.CoInitialize()
        wd_app = None
        ok_count = warn_count = err_count = 0
        try:
            wd_app = win32com.client.DispatchEx("Word.Application")
            wd_app.Visible = False
            wd_app.DisplayAlerts = 0

            for i, (point, path) in enumerate(files):
                if self._is_cancelled:
                    self.info.emit("Остановлено пользователем", "warn")
                    break
                fname = os.path.basename(path)
                self.log.emit(f"\n── {point} / {fname} ──")
                self.info.emit(f"{point}: правка «{fname}»…", "wait")

                doc = None
                try:
                    doc = self._open_doc(wd_app, path)
                    parts = [
                        self._fix_title_date(doc, new_date),
                        self._fix_certificates(doc),
                        self._fix_table(doc),
                        self._fix_order_scan(doc),
                    ]
                    doc.Save()
                    safe_close_com(doc)
                    doc = None

                    has_warn = any(p.startswith("⚠") for p in parts)
                    summary = ", ".join(parts)
                    self.log.emit(f"   {summary}")
                    if has_warn:
                        warn_count += 1
                        self.info.emit(f"{point} · {fname}: {summary}", "warn")
                    else:
                        ok_count += 1
                        self.info.emit(f"{point} · {fname}: {summary}", "done")
                except Exception as e:
                    err_count += 1
                    self.log.emit(f"   ❌ Ошибка: {e}\n   Путь: {path}")
                    self.info.emit(f"{point} · {fname}: ошибка Word — {e}", "error")
                    safe_close_com(doc)

                self.progress.emit(int((i + 1) / len(files) * 100))

            total_done = ok_count + warn_count
            summary = f"Готово: {total_done} из {len(files)}"
            if warn_count:
                summary += f", с предупреждениями {warn_count}"
            if err_count:
                summary += f", ошибок {err_count}"
            self.info.emit(summary, "error" if err_count else "done")
            self.finished_ok.emit(err_count == 0)
        except Exception as e:
            self.log.emit(f"❌ Критическая ошибка: {e}")
            self.info.emit(f"Критическая ошибка: {e}", "error")
            self.finished_ok.emit(False)
        finally:
            if wd_app:
                try:
                    wd_app.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
