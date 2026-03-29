# -*- coding: utf-8 -*-
"""
Этап 4: Создание отчёта.
Открытие ПНОС_ТТП.xlsm, запуск макроса TransTab, 
затем открытие ПНОС_ТТП.docx (с обработкой окна SQL),
выполнение слияния (MailMerge) и сохранение нового файла Отчёт.docx.
"""

import os
import threading
import time

from PySide6.QtCore import QThread, Signal

import pythoncom
import win32com.client
import win32gui
import win32con

from config import CONFIG
from status_manager import get_stage_status, set_stage_status

def safe_close_com(obj, save=False):
    """Безопасное закрытие COM-объекта (книги или документа)."""
    if obj is None:
        return
    try:
        # Пытаемся разблокировать документ перед закрытием, если это Word
        if hasattr(obj, "MailMerge"):
            obj.MailMerge.MainDocumentType = -1 # wdNotAMergeDocument
        
        if save:
            obj.Save()
        obj.Close(SaveChanges=False)
    except Exception:
        pass


import win32process


class WordDialogKiller:
    """Охотник за окнами: нажимает Enter (Да) на любом диалоговом окне процесса Word."""
    def __init__(self, log_callback, word_pid: int = 0):
        self._stop = threading.Event()
        self._t = None
        self.count = 0
        self.log_callback = log_callback
        self.word_pid = word_pid  # PID чтобы ловить только его диалоги

    def start(self):
        self._t = threading.Thread(target=self._run, daemon=True)
        self._t.start()

    def stop(self):
        self._stop.set()
        if self._t:
            self._t.join(timeout=2)

    def _run(self):
        while not self._stop.is_set():
            try:
                win32gui.EnumWindows(self._cb, None)
            except Exception:
                pass
            self._stop.wait(0.15)  # Проверяем ~7 раз в секунду

    def _cb(self, hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True

            title = (win32gui.GetWindowText(hwnd) or "").lower()
            
            # Ловим:
            # 1) Любое окно процесса Word (если знаем PID)
            # 2) Окна с характерными словами в заголовке
            _, win_pid = win32process.GetWindowThreadProcessId(hwnd)
            matches_word_pid = (self.word_pid > 0 and win_pid == self.word_pid)
            
            # Расширенный список ключевых слов (английские и русские)
            keywords = (
                "sql", "select", "селект", "microsoft word", 
                "выбор таблицы", "select table", "источник данных", "data source",
                "подключение", "connection", "фильтр", "связи", "таблицы", "экспорт",
                "не удалось установить", "ошибка", "microsoft excel", "microsoft word"
            )
            matches_sql_title = any(k in title for k in keywords)

            if matches_word_pid or matches_sql_title:
                title_log = title if title else "[Без заголовка]"
                # Логируем только если это реально похоже на диалог выбора (чтобы не спамить)
                if "выбор" in title or "таблиц" in title or "select" in title or "sql" in title or matches_word_pid:
                    self.log_callback(f"   [DialogKiller] Вижу окно '{title_log}' (PID match: {matches_word_pid}), нажимаем Enter...")
                    # Пробуем несколько способов нажать кнопку
                    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                    
                    # Жесткий метод если PostMessage не сработал
                    win32gui.SetForegroundWindow(hwnd)
                    import win32com.client
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shell.SendKeys("~") # Enter
                    
                    self.count += 1
        except Exception:
            pass
        return True


def find_file(folder: str, exts: tuple) -> str | None:
    """Ищет первый попавшийся файл с заданным расширением в папке."""
    if not folder or not os.path.exists(folder):
        return None
    for f in sorted(os.listdir(folder)):
        if f.startswith("~$"):
            continue
        p = os.path.join(folder, f)
        if os.path.isfile(p) and os.path.splitext(f)[1].lower() in exts:
            return p
    return None


class Stage4Worker(QThread):
    log = Signal(str)             # Сообщение в лог
    progress = Signal(int)        # Прогресс 0-100
    action_update = Signal(str)   # Текст для UI над прогресс-баром
    report_ready = Signal(list)   # Список словарей для модального окна-отчёта
    finished_ok = Signal(bool)    # Завершение (True = успех)
    info = Signal(str, str)       # Дружелюбный статус (сообщение, категория)

    def __init__(self, local_root: str, parent=None):
        super().__init__(parent)
        self.local_root = local_root
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        self.log.emit("═" * 40)
        self.log.emit("ЭТАП 4: Генерация финального отчёта")
        self.log.emit("═" * 40)

        if not os.path.isdir(self.local_root):
            self.log.emit("❌ Корневая папка не найдена.")
            self.finished_ok.emit(False)
            return

        point_folders = []
        for d in os.listdir(self.local_root):
            dpath = os.path.join(self.local_root, d)
            if os.path.isdir(dpath) and d.startswith("п."):
                point_folders.append(dpath)

        if not point_folders:
            self.log.emit("⚠ Не найдено папок пунктов (п.*).")
            self.finished_ok.emit(False)
            return

        self.log.emit(f"📋 Найдено пунктов для обработки: {len(point_folders)}")
        self.progress.emit(5)

        pythoncom.CoInitialize()
        xl_app = None
        wd_app = None
        wd_dk = None

        try:
            self.log.emit("\n⚙ Инициализация приложений Office...")
            xl_app = win32com.client.Dispatch("Excel.Application")
            xl_app.Visible = False
            xl_app.DisplayAlerts = False
            
            wd_app = win32com.client.DispatchEx("Word.Application")
            wd_app.Visible = False
            wd_app.DisplayAlerts = 0 # wdAlertsNone
            
            word_pid = 0
            try:
                word_hwnd = wd_app.Application.Hwnd
                if word_hwnd:
                    _, word_pid = win32process.GetWindowThreadProcessId(word_hwnd)
            except Exception: pass
            
            wd_dk = WordDialogKiller(log_callback=self.log.emit, word_pid=word_pid)
            wd_dk.start()

            self.info.emit("Инструменты для отчетов готовы", "done")
            self.progress.emit(10)
            processed_count = 0
            errors_count = 0
            report_data = []

            for i, p_folder in enumerate(point_folders):
                if self._is_cancelled:
                    self.log.emit("\n❌ Отменено пользователем.")
                    break

                p_name = os.path.basename(p_folder)
                self.info.emit(f"Генерация отчета {p_name}...", "wait")
                self.log.emit(f"\n── {p_name} ──")
                pt_num_str = p_name[2:] if p_name.startswith("п.") else p_name
                
                point_report = {
                    "№ Пункта": pt_num_str,
                    "Макрос Excel": False,
                    "Слияние Word": False,
                    "Результат": "Ожидание"
                }
                
                if get_stage_status(p_folder, "stage4"):
                    self.log.emit(f"   ⏭ Этап 4 уже завершён, пропускаем.")
                    point_report["Результат"] = "⚪ Пропущен"
                    report_data.append(point_report)
                    processed_count += 1
                    continue
                
                work_wb = None
                doc = None

                try:
                    excel_path = find_file(p_folder, (".xls", ".xlsx", ".xlsm"))
                    word_path = find_file(p_folder, (".doc", ".docx"))
                    
                    if not excel_path or not word_path:
                        self.log.emit(f"   ✗ Файлы не найдены.")
                        point_report["Результат"] = "❌ Файлы не найдены"
                        errors_count += 1
                        continue

                    # --- ШАГ 1: Открываем Word (ОБЯЗАТЕЛЬНО ПЕРВЫМ для TransTab) ---
                    self.action_update.emit(f"[{pt_num_str}] Открытие Word...")
                    self.log.emit(f"   [Word] Открытие документа...")
                    doc = wd_app.Documents.Open(word_path, ConfirmConversions=False, Visible=True)
                    
                    # Пытаемся поймать PID еще раз
                    if not wd_dk.word_pid:
                        try:
                            _, pid = win32process.GetWindowThreadProcessId(wd_app.Application.Hwnd)
                            wd_dk.word_pid = pid
                        except Exception: pass

                    # --- ШАГ 2: Открываем Excel и запускаем макрос ---
                    self.action_update.emit(f"[{pt_num_str}] Запуск макроса...")
                    self.log.emit(f"   [Excel] Открытие книги...")
                    work_wb = xl_app.Workbooks.Open(excel_path, UpdateLinks=0)
                    
                    try:
                        self.log.emit(f"   [Excel] Выполнение TransTab...")
                        xl_app.Run(f"'{work_wb.Name}'!modTransTab.TransTab")
                        self.log.emit(f"   ✓ [Excel] Макрос выполнен.")
                        point_report["Макрос Excel"] = True
                    except Exception as e:
                        self.log.emit(f"   ⚠ [Excel] Ошибка макроса: {e}")

                    # --- ШАГ 3: Создаем CleanSource.xlsx ---
                    clean_source_path = os.path.join(p_folder, "CleanSource.xlsx")
                    try:
                        ws_export = None
                        for sh in work_wb.Sheets:
                            if sh.Name.lower() == "экспорт":
                                ws_export = sh
                                break
                        
                        if ws_export:
                            new_wb = xl_app.Workbooks.Add()
                            ws_export.Copy(Before=new_wb.Sheets(1))
                            new_ws = new_wb.Sheets(1)
                            new_ws.Visible = -1 # Visible
                            
                            xl_app.DisplayAlerts = False
                            try:
                                while new_wb.Sheets.Count > 1:
                                    new_wb.Sheets(new_wb.Sheets.Count).Delete()
                            except Exception: pass
                            xl_app.DisplayAlerts = True
                            
                            new_wb.SaveAs(clean_source_path, 51) # xlsx
                            new_wb.Close()
                            self.log.emit(f"   ✓ [Excel] CleanSource создан.")
                        else:
                            clean_source_path = excel_path
                    except Exception as e:
                        self.log.emit(f"   ⚠ [Excel] Ошибка CleanSource: {e}")
                        clean_source_path = excel_path

                    # ЗАКРЫВАЕМ EXCEL перед слиянием
                    if work_wb:
                        work_wb.Save()
                        work_wb.Close()
                        work_wb = None
                    time.sleep(0.5)

                    # --- ШАГ 4: Mail Merge в Word ---
                    self.action_update.emit(f"[{pt_num_str}] Слияние...")
                    try:
                        self.log.emit(f"   [Word] Привязка данных...")
                        doc.MailMerge.MainDocumentType = 0  # wdFormLetters
                        
                        # Даем DialogKiller шанс
                        time.sleep(0.5)
                        
                        doc.MailMerge.OpenDataSource(
                            Name=clean_source_path,
                            SQLStatement="SELECT * FROM [экспорт$]",
                            SubType=1 # SubTypeAccess
                        )
                        
                        # "Заморозка" данных
                        doc.Fields.Update()
                        doc.MailMerge.ViewMailMergeFieldCodes = False
                        doc.Fields.Unlink()
                        doc.MailMerge.MainDocumentType = -1
                        
                        out_name = f"Отчет_{pt_num_str}.docx"
                        out_path = os.path.join(p_folder, out_name)
                        doc.SaveAs2(out_path, 16)
                        
                        self.log.emit(f"   ✓ [Word] Отчет готов: {out_name}")
                        self.info.emit(f"{p_name} — отчет создан", "done")
                        point_report["Слияние Word"] = True
                        point_report["Результат"] = "✅ Успешно"
                        
                        set_stage_status(p_folder, "stage4", True)
                        
                        processed_count += 1
                    except Exception as e:
                        self.log.emit(f"   ❌ [Word] Ошибка: {e}")
                        point_report["Результат"] = "❌ Ошибка Word"

                except Exception as ex:
                    self.log.emit(f"   ❌ Ошибка: {ex}")
                    errors_count += 1
                finally:
                    report_data.append(point_report)
                    if work_wb:
                        try: work_wb.Close(False)
                        except Exception: pass
                    if doc:
                        try: doc.Close(False)
                        except Exception: pass
                    
                    # Чистка временных файлов
                    try:
                        tmp = os.path.join(p_folder, "CleanSource.xlsx")
                        if os.path.exists(tmp): os.remove(tmp)
                    except Exception: pass

                pct = 10 + int((i + 1) / len(point_folders) * 85)
                self.progress.emit(min(pct, 95))

            self.progress.emit(100)
            self.info.emit("Этап 4 успешно завершен!", "done")
            self.log.emit(f"\n✅ Этап 4 завершён! Успешно: {processed_count}, Ошибок: {errors_count}")
            self.report_ready.emit(report_data)
            self.action_update.emit("Готово!")
            self.finished_ok.emit(not self._is_cancelled)

        except Exception as e:
            self.info.emit(f"Ошибка этапа 4: {str(e)[:50]}", "error")
            self.log.emit(f"\n❌ Критическая ошибка: {e}")
            self.finished_ok.emit(False)
        finally:
            if wd_dk: wd_dk.stop()
            if xl_app: 
                try: xl_app.Quit()
                except Exception: pass
            if wd_app:
                try: wd_app.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()

