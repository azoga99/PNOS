# -*- coding: utf-8 -*-
"""
Этап 3: Запуск макросов внутри ПНОС_ТТП.xlsm
Перебор всех пунктов, подтягивание переменных из Word и выполнение макроса Excel.
"""

import os
import shutil
import tempfile
import difflib
import threading
import re
import zipfile
import time
import sys

from PySide6.QtCore import QThread, Signal

# Сторонние библиотеки для COM
import pythoncom
import win32com.client
import win32gui
import win32con

from config import CONFIG


# =====================================================================
# Вспомогательные функции (портированы из старого 2 этап.txt)
# =====================================================================

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

def find_passport(base: str, target: str, cutoff: float) -> str | None:
    """Нечёткий поиск папки паспорта среди подпапок пункта."""
    if not os.path.exists(base):
        return None
    dirs = {}
    for d in os.listdir(base):
        if os.path.isdir(os.path.join(base, d)):
            dirs[d.lower()] = d
    if target in dirs:
        return os.path.join(base, dirs[target])
    
    m = difflib.get_close_matches(target, dirs.keys(), n=1, cutoff=cutoff)
    return os.path.join(base, dirs[m[0]]) if m else None

def safe_close_com(obj, save=False):
    """Безопасное закрытие COM-объекта (книги или документа)."""
    if obj is None:
        return
    try:
        if save:
            obj.Save()
        obj.Close(SaveChanges=False)
    except Exception:
        pass


class DialogKiller:
    """Фоновый поток для автоматического закрытия всплывающих ошибок Excel."""
    def __init__(self, log_callback):
        self._stop = threading.Event()
        self._t = None
        self.count = 0
        self.log_callback = log_callback

    def start(self):
        self._t = threading.Thread(target=self._run, daemon=True)
        self._t.start()

    def stop(self):
        self._stop.set()
        if self._t:
            self._t.join(timeout=2)
        if self.count and self.log_callback:
            self.log_callback(f"[DialogKiller] Автоматически закрыто ошибок: {self.count}")

    def _run(self):
        while not self._stop.is_set():
            try:
                win32gui.EnumWindows(self._cb, None)
            except Exception:
                pass
            self._stop.wait(0.08)

    def _cb(self, hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            t = (win32gui.GetWindowText(hwnd) or "").lower()
            if any(k in t for k in ("excel", "имен", "name", "конфликт")):
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                self.count += 1
        except Exception:
            pass
        return True


# =====================================================================
# QThread Worker
# =====================================================================

class Stage3Worker(QThread):
    """Фоновый поток для Этапа 3 — выполнение макросов в Excel."""

    log = Signal(str)             # Сообщение в лог
    progress = Signal(int)        # Прогресс 0-100
    action_update = Signal(str)   # Текст для UI над прогресс-баром
    report_ready = Signal(list)   # Список словарей для модального окна-отчёта
    table_preview = Signal(list)  # Срез для превью таблицы 3 (list of lists)
    finished_ok = Signal(bool)    # Завершение (True = успех)
    info = Signal(str, str)       # Дружелюбный статус (сообщение, категория)

    def __init__(self, local_root: str, macro_master_path: str = None, parent=None):
        super().__init__(parent)
        self.local_root = local_root
        self.macro_master_path = macro_master_path
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def _clean_pnos_copy(self, master_path: str):
        """Создает временную копию ПНОС без XML тегов definedNames."""
        tmp_dir = tempfile.mkdtemp(prefix="pnos_")
        tmp_path = os.path.join(tmp_dir, os.path.basename(master_path))

        ext = os.path.splitext(master_path)[1].lower()
        if ext not in (".xlsx", ".xlsm"):
            shutil.copy2(master_path, tmp_path)
            return tmp_path, tmp_dir

        try:
            with zipfile.ZipFile(master_path, 'r') as zin, \
                 zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "xl/workbook.xml":
                        txt = data.decode("utf-8")
                        cnt = len(re.findall(r'<definedName ', txt))
                        # Удаляем все теги имён
                        data = re.sub(r'<definedNames>.*?</definedNames>', '', txt, flags=re.DOTALL).encode("utf-8")
                        self.log.emit(f"   [Очистка макроса] Удалено имён: {cnt}")
                    zout.writestr(item, data)
        except Exception as e:
            self.log.emit(f"   [Очистка макроса] Ошибка очистки XML: {e}, используем исходный как есть")
            shutil.copy2(master_path, tmp_path)

        return tmp_path, tmp_dir

    def run(self):
        self.log.emit("═" * 40)
        self.log.emit("ЭТАП 3: Запуск макросов обработки")
        self.log.emit("═" * 40)

        # 1. Проверки
        if not os.path.isdir(self.local_root):
            self.log.emit("❌ Корневая папка не найдена.")
            self.finished_ok.emit(False)
            return

        master_path = self.macro_master_path or CONFIG.get("MACRO_MASTER_WB_PATH", "")
        if not master_path or not os.path.isfile(master_path):
            self.log.emit("❌ Главный файл с макросом не найден.")
            self.log.emit(f"Проверьте путь: {master_path}")
            self.log.emit("Пожалуйста, укажите полный путь в настройках или в config.py.")
            self.finished_ok.emit(False)
            return

        macro_cfg = CONFIG.get("MACRO_CONFIG", {})

        # Ищем все папки пунктов "п.*"
        point_folders = []
        for d in os.listdir(self.local_root):
            dpath = os.path.join(self.local_root, d)
            if os.path.isdir(dpath) and d.startswith("п."):
                point_folders.append(dpath)

        if not point_folders:
            self.log.emit("⚠ Не найдено локальных папок пунктов (п.*).")
            self.finished_ok.emit(False)
            return

        self.log.emit(f"📋 Найдено пунктов для обработки: {len(point_folders)}")
        self.progress.emit(5)

        # 2. Инициализация COM и Диалог Киллера
        pythoncom.CoInitialize()
        dk = DialogKiller(log_callback=self.log.emit)
        
        xl_app = None
        wd_app = None
        pnos_master_wb = None
        pnos_tmp_path = None
        pnos_tmp_dir = None

        try:
            dk.start()
            
            self.log.emit("\n⚙ Инициализация Excel и Word в фоне...")
            xl_app = win32com.client.Dispatch("Excel.Application")
            xl_app.Visible = False
            xl_app.DisplayAlerts = False
            xl_app.AskToUpdateLinks = False
            xl_app.ScreenUpdating = False
            xl_app.EnableEvents = False
            try:
                xl_app.AutomationSecurity = 1
            except Exception:
                pass

            wd_app = win32com.client.Dispatch("Word.Application")
            wd_app.Visible = False
            wd_app.DisplayAlerts = 0

            self.progress.emit(10)

            self.log.emit("\n🔧 Очистка главного файла ПНОС...")
            self.info.emit("Подготовка мастер-файла ПНОС...", "wait")
            pnos_tmp_path, pnos_tmp_dir = self._clean_pnos_copy(master_path)
            
            pnos_master_wb = xl_app.Workbooks.Open(
                pnos_tmp_path, UpdateLinks=0, ReadOnly=True,
                IgnoreReadOnlyRecommended=True, Notify=False)

            try:
                xl_app.Calculation = -4135  # xlManual  (Ускоряет открытие)
            except Exception:
                pass

            self.info.emit("Excel и Word запущены", "done")
            self.log.emit("✅ COM-объекты готовы. Начинаем обработку.")

            # --- Сбор маппинга из Мастер-файла ---
            self.log.emit("\n🔧 Чтение столбцов A и B из листа 'ПЛАН' Мастер-файла...")
            id_mapping = {}
            try:
                master_ws = pnos_master_wb.Sheets("ПЛАН")
                
                # Ищем последнюю заполненную строку в столбце B (-4162 = xlUp)
                last_row = master_ws.Cells(master_ws.Rows.Count, "B").End(-4162).Row
                
                if last_row > 1:
                    # Читаем сразу весь массив для максимальной скорости (от A1 до B_last_row)
                    data_range = master_ws.Range(f"A1:B{last_row}").Value
                    
                    if data_range:
                        # data_range это tuple из tuples: ((A1, B1), (A2, B2), ...)
                        for row_val in data_range:
                            col_a = row_val[0]
                            col_b = row_val[1]
                            
                            if col_b is not None:
                                try:
                                    key = int(float(col_b))
                                except (ValueError, TypeError):
                                    key = str(col_b).strip()
                                
                                id_mapping[key] = col_a
                                
                self.log.emit(f"   ✓ Загружено {len(id_mapping)} связей.")
            except Exception as e:
                self.log.emit(f"   ⚠ Ошибка при чтении листа 'ПЛАН': {e}")
            # ----------------------------------------

            # 3. Цикл обработки пунктов
            processed_count = 0
            errors_count = 0
            report_data = []

            for i, p_folder in enumerate(point_folders):
                if self._is_cancelled:
                    self.log.emit("\n❌ Выполнение отменено пользователем!")
                    break

                p_name = os.path.basename(p_folder)
                self.info.emit(f"Обработка {p_name}...", "wait")
                self.log.emit(f"\n── {p_name} ──")
                
                # Парсим номер пункта из папки "п.1234"
                p_number_str = p_name[2:] if p_name.startswith("п.") else p_name
                try:
                    p_number = int(p_number_str)
                except ValueError:
                    p_number = 2719 # Фолбэк на дефолт если странное название
                
                # --- Переопределение p_number через Мастер-файл ---
                # Ищем по числу:
                mapped_id = id_mapping.get(p_number)
                if mapped_id is None:
                    # Если не нашли, ищем по строке:
                    mapped_id = id_mapping.get(str(p_number))
                
                if mapped_id is not None:
                    self.log.emit(f"   [Маппинг] Найдено соответствие: п.{p_number} -> Столбец А: {mapped_id}")
                    p_number_for_f9 = mapped_id
                else:
                    self.log.emit(f"   [Маппинг] ⚠ п.{p_number} не найден в столбце B листа 'ПЛАН'. Оставляем {p_number}.")
                    p_number_for_f9 = p_number
                # --------------------------------------------------

                self.action_update.emit(f"Извлечение из паспорта для {p_number_str}...")
                
                point_report = {
                    "№ Пункта": p_number_str,
                    "Паспорт": False,
                    "Excel": False,
                    "Результат": "Ожидание"
                }
                
                work_wb = None
                doc = None

                try:
                    # Поиск файла ПНОС_ТТП.xlsm
                    work_path = find_file(p_folder, (".xls", ".xlsx", ".xlsm"))
                    if not work_path:
                        self.log.emit(f"   ✗ Ошибка: Excel-файл (ПНОС_ТТП.xlsm) не найден в папке")
                        errors_count += 1
                        continue

                    # Поиск паспорта
                    pf = find_passport(p_folder, macro_cfg.get("PASSPORT_TARGET", "паспорт"), macro_cfg.get("FUZZY_CUTOFF", 0.7))
                    wf = find_file(pf, (".doc", ".docx"))
                    if not pf or not wf:
                        self.log.emit(f"   ✗ Ошибка: Word-документ Паспорта не найден")
                        point_report["Результат"] = "❌ Нет Паспорта"
                        errors_count += 1
                        report_data.append(point_report)
                        continue
                        
                    point_report["Паспорт"] = True
                    point_report["Excel"] = True # т.к. work_path найден выше

                    # -- WORD: Читаем переменные --
                    # Читаем Word первым, чтобы он не задерживал Excel
                    doc = wd_app.Documents.Open(wf, ReadOnly=True)
                    dv = {}
                    variables = doc.Variables
                    for v_idx in range(1, variables.Count + 1):
                        v = variables.Item(v_idx)
                        dv[v.Name] = v.Value
                    pages = doc.ComputeStatistics(2)
                    doc.Close(SaveChanges=False)
                    doc = None

                    # -- EXCEL: Открытие и Макрос --
                    work_wb = xl_app.Workbooks.Open(
                        work_path, UpdateLinks=0,
                        IgnoreReadOnlyRecommended=True, Notify=False)

                    ws = work_wb.Sheets("паспорт")
                    
                    # Пишем номер пункта (из столбца А Мастер-файла, либо оригинальный)
                    ws.Range("F9").Value = p_number_for_f9
                    
                    # Пишем Word переменные
                    for var_name, cell_addr in macro_cfg.get("DOCVARIABLE_MAP", {}).items():
                        v = dv.get(var_name)
                        if v is not None:
                            ws.Range(cell_addr).Value = v
                        else:
                            self.log.emit(f"   [!] У Word-файла нет переменной «{var_name}»")
                    
                    # Пишем количество страниц
                    page_count_cell = macro_cfg.get("WORD_PAGE_COUNT_CELL", "B48")
                    ws.Range(page_count_cell).Value = pages
                    
                    # -- Обновление Сводных Таблиц (Pivot Tables) --
                    self.action_update.emit(f"Обновление таблиц для {p_number_str}...")
                    try:
                        work_wb.RefreshAll()
                        self.log.emit(f"   [Excel] Сводные таблицы обновлены (RefreshAll).")
                    except Exception as e:
                        self.log.emit(f"   ⚠ Ошибка при RefreshAll: {e}")

                    # Отправка превью одной из обновленных таблиц (например, ТрубыСв)
                    try:
                        preview_ws = work_wb.Sheets("ТрубыСв")
                        # Берем небольшой диапазон, например A4:E8
                        preview_range = preview_ws.Range("A4:E8").Value
                        if preview_range:
                            # Очищаем от None
                            clean_preview = []
                            for row in preview_range:
                                clean_row = [str(cell) if cell is not None else "" for cell in row]
                                # Если строка не пустая, добавляем
                                if any(clean_row):
                                    clean_preview.append(clean_row)
                            self.table_preview.emit(clean_preview)
                    except Exception as e:
                        self.log.emit(f"   ⚠ Не удалось получить превью таблицы 'ТрубыСв': {e}")

                    # Запуск макроса Restore -> Save
                    self.action_update.emit(f"Выполнение макроса Restore для {p_number_str}...")
                    work_wb.Activate()
                    ws.Activate()
                    try:
                        xl_app.Calculation = -4105  # xlAutomatic (Нужно для VLOOKUP перед макросом)
                    except Exception:
                        pass

                    xl_app.EnableEvents = True
                    xl_app.DisplayAlerts = False
                    
                    wb_name = work_wb.Name
                    m_wait = macro_cfg.get("MACRO_WAIT", 1)

                    xl_app.Run(f"'{wb_name}'!Restore")
                    xl_app.Calculate()
                    time.sleep(m_wait)

                    xl_app.Run(f"'{wb_name}'!Save")
                    time.sleep(m_wait)

                    xl_app.EnableEvents = False
                    try:
                        xl_app.Calculation = -4135  # Возвращаем Manual
                    except Exception:
                        pass

                    # Сохранение
                    work_wb.Save()
                    work_wb.Close()
                    work_wb = None

                    self.log.emit(f"   ✓ Успех. Данные перенесены, макрос выполнен.")
                    self.info.emit(f"{p_name} — макрос выполнен", "done")
                    point_report["Результат"] = "✅ Успешно"
                    processed_count += 1

                except Exception as ex:
                    self.log.emit(f"   ❌ Ошибка при обработке {p_name}: {ex}")
                    point_report["Результат"] = "❌ Ошибка"
                    errors_count += 1
                finally:
                    report_data.append(point_report)
                    # На всякий случай подчищаем, если пункт упал с ошибкой
                    safe_close_com(work_wb, save=False)
                    safe_close_com(doc, save=False)

                # Шаг прогресса: 10% .. 95%
                pct = 10 + int((i + 1) / len(point_folders) * 85)
                self.progress.emit(min(pct, 95))

            # 4. Финиш
            self.progress.emit(100)
            self.info.emit("Этап 3 завершен!", "done")
            self.log.emit(f"\n{'═' * 40}")
            self.log.emit(f"✅ Этап 3 завершён! Обработано успешно: {processed_count}, Ошибок: {errors_count}")

            # Сигнализируем успех, если не было сплошных ошибок
            success = not self._is_cancelled and (processed_count > 0 or len(point_folders) == 0)
            self.report_ready.emit(report_data)
            self.action_update.emit("Завершено!")
            self.finished_ok.emit(success)

        except Exception as e:
            self.info.emit(f"Ошибка этапа 3: {str(e)[:50]}", "error")
            self.log.emit(f"\n❌ Критическая ошибка COM: {e}")
            import traceback
            self.log.emit(traceback.format_exc())
            self.finished_ok.emit(False)

        finally:
            dk.stop()
            self.log.emit("🧹 Закрытие фоновых программ Office...")
            
            safe_close_com(pnos_master_wb)
            
            if xl_app:
                try:
                    xl_app.ScreenUpdating = True
                    xl_app.Calculation = -4105
                    xl_app.Quit()
                except Exception:
                    pass
            if wd_app:
                try:
                    wd_app.Quit()
                except Exception:
                    pass
            
            # Удаление временной папки "очищенного" ПНОС
            if pnos_tmp_path and pnos_tmp_dir:
                try:
                    os.remove(pnos_tmp_path)
                    os.rmdir(pnos_tmp_dir)
                except Exception:
                    pass
            
            pythoncom.CoUninitialize()
