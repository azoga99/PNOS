# -*- coding: utf-8 -*-
"""
Этап 4: Создание отчёта.
Открытие ПНОС_ТТП.xlsm, запуск макроса TransTab, 
затем открытие ПНОС_ТТП.docx (с обработкой окна SQL),
выполнение слияния (MailMerge) и сохранение нового файла Отчёт.docx.
"""

import os
import re
import difflib
import math
import threading
import time
from datetime import datetime

from PySide6.QtCore import QThread, Signal

import pythoncom
import win32com.client
import win32gui
import win32con

from config import CONFIG
from status_manager import get_stage_status, set_stage_status, is_point_blocked
from workers.win_dialogs import (
    click_dialog_button, dialog_buttons, describe_dialogs,
)

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
    """
    Охотник за окнами: жмёт «Да»/«ОК» на диалоговых окнах процесса Word.

    Кнопка нажимается сообщением (см. workers.win_dialogs) — без фокуса и без
    вывода окна на передний план. Прежний способ (SetForegroundWindow плюс
    SendKeys) работал только когда ПНОС сам был активен, и слияние на этапе
    «Создание отчёта» стояло, пока пользователь не кликнет по окну программы.
    """
    def __init__(self, log_callback, word_pid: int = 0):
        self._stop = threading.Event()
        self._t = None
        self.count = 0
        self.log_callback = log_callback
        self.word_pid = word_pid  # PID чтобы ловить только его диалоги
        self._seen = set()        # заголовки, о которых уже писали в лог

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

            if not (matches_word_pid or matches_sql_title):
                return True

            # Класс окна не проверяем: в разных версиях Word диалоги бывают не
            # только «#32770». Признак диалога — наличие кнопок «Да»/«ОК»;
            # у главного окна Word таких дочерних кнопок нет.
            кнопки = dialog_buttons(hwnd)
            if not кнопки:
                return True

            # Логируем каждый заголовок один раз — иначе журнал забивается:
            # охотник обходит окна семь раз в секунду.
            новый = title not in self._seen
            self._seen.add(title)
            нажата = click_dialog_button(
                hwnd, log=self.log_callback if новый else None)
            if нажата:
                self.count += 1
                return True

            # Кнопки есть, но утвердительной среди них нет — пробуем прежний
            # способ, вдруг у диалога нестандартная раскладка.
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            if новый:
                подписи = ", ".join(f"«{c}»" for _, c in кнопки)
                self.log_callback(
                    f"   [DialogKiller] диалог «{title or '[без заголовка]'}»: "
                    f"утвердительной кнопки нет (есть: {подписи}), послан Enter")
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


def _clean_report_num(v) -> str:
    """Приводит значение ячейки к целому числу-строке ('2744.0' -> '2744').
    Нечисловое возвращает как есть (обрезав пробелы); пустое -> ''."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    try:
        return str(int(float(s.replace(",", "."))))
    except (ValueError, TypeError):
        return s


def build_report_name(f9, b17, pt_num_str: str) -> str:
    """Строит имя итогового отчёта.
    'ЗЭПБ № ТТП-<F9>.<YY> - ттп № <B17>.docx', напр. 'ЗЭПБ № ТТП-2744.26 - ттп № 205.docx'.
    Если F9 или B17 пусты/нечитаемы — возвращает прежнее имя 'Отчет_<номер>.docx',
    чтобы пайплайн не падал.
    """
    cfg = CONFIG.get("REPORT_NAME", {})
    yy = cfg.get("YEAR_SUFFIX", "26")
    f9c = _clean_report_num(f9)
    b17c = _clean_report_num(b17)
    if f9c and b17c:
        name = f"ЗЭПБ № ТТП-{f9c}.{yy} - ттп № {b17c}.docx"
        # Страховка от недопустимых для Windows символов в имени файла
        for ch in '\\/:*?"<>|':
            name = name.replace(ch, "")
        return name
    return f"Отчет_{pt_num_str}.docx"


def find_report_file(p_folder: str) -> str | None:
    """Находит готовый отчёт в папке пункта.
    Матчит и новое имя ('ЗЭПБ … .docx'), и старое ('Отчет_*.docx') — чтобы работали
    повторные прогоны и уже переименованные вручную файлы. Предпочитает 'ЗЭПБ'.

    Если 'ЗЭПБ'-файлов несколько (например, рядом с основным лежит копия
    '... ПЕЧАТЬ.docx' для печати) — берётся файл БЕЗ 'ПЕЧАТЬ' в имени, и выбор
    детерминирован (сортировка), а не зависит от порядка os.listdir.
    """
    zepb_candidates = []
    otchet_candidates = []
    for f in sorted(os.listdir(p_folder)):
        if f.startswith("~$") or not f.lower().endswith(".docx"):
            continue
        if f.startswith("ЗЭПБ"):
            zepb_candidates.append(f)
        elif f.startswith("Отчет_"):
            otchet_candidates.append(f)

    if zepb_candidates:
        non_print = [f for f in zepb_candidates if "печать" not in f.lower()]
        chosen = sorted(non_print)[0] if non_print else sorted(zepb_candidates)[0]
        return os.path.join(p_folder, chosen)
    if otchet_candidates:
        return os.path.join(p_folder, sorted(otchet_candidates)[0])
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

    MERGE_WATCHDOG_SEC = 20

    def _merge_watchdog(self, word_pid: int) -> threading.Event:
        """
        Сторож долгого COM-вызова: через MERGE_WATCHDOG_SEC пишет в журнал,
        какие диалоги Word висят на экране. Возвращает Event — вызывающая
        сторона взводит его, когда вызов вернулся.

        Нужен для диагностики: диалог с непривычной раскладкой кнопок охотник
        закрыть не сможет, и без такой записи этап выглядит просто зависшим.
        """
        done = threading.Event()

        def _watch():
            if done.wait(self.MERGE_WATCHDOG_SEC):
                return
            окна = describe_dialogs(word_pid)
            if окна:
                self.log.emit("   ⏳ [Word] Привязка данных идёт дольше обычного; "
                              "на экране диалоги: " + "; ".join(окна))
            else:
                self.log.emit("   ⏳ [Word] Привязка данных идёт дольше обычного, "
                              "диалогов не видно — ждём Word.")

        threading.Thread(target=_watch, daemon=True).start()
        return done

    def _find_specification_file(self, p_folder: str) -> str | None:
        """Ищет файл 'Спецификация трубопровода.docx' в папке Первичка."""
        pervichka = None
        for d in os.listdir(p_folder):
            if os.path.isdir(os.path.join(p_folder, d)):
                matches = difflib.get_close_matches("первичка", [d.lower()], n=1, cutoff=0.7)
                if matches:
                    pervichka = os.path.join(p_folder, d)
                    break
        if not pervichka:
            return None

        for f in os.listdir(pervichka):
            if f.startswith("~$"):
                continue
            if "спецификация" in f.lower() and f.lower().endswith((".docx", ".doc")):
                return os.path.join(pervichka, f)
        return None

    def _parse_specification(self, wd_app, spec_path: str) -> dict:
        """
        Парсит таблицу из 'Спецификация трубопровода.docx'.
        Возвращает словарь: {номер_str: {"dy": str, "py": str, "steel": str}}
        """
        result = {}
        spec_doc = None
        try:
            spec_doc = wd_app.Documents.Open(spec_path, ReadOnly=True, Visible=False)
            if spec_doc.Tables.Count == 0:
                self.log.emit(f"   ⚠ [Спецификация] Таблица не найдена в файле")
                return result

            table = spec_doc.Tables(1)
            rows_count = table.Rows.Count
            cols_count = table.Columns.Count

            for r in range(1, rows_count + 1):
                try:
                    # Столбец 0 (индекс 1 в COM) — номер
                    cell_num_text = table.Cell(r, 1).Range.Text.strip().rstrip("\r\x07")
                    # Убираем все нечисловые символы для сравнения
                    num_clean = cell_num_text.strip()
                    if not num_clean:
                        continue

                    # Столбец 1 (индекс 2 в COM) — описание с Dy и Py
                    if cols_count >= 2:
                        desc_text = table.Cell(r, 2).Range.Text.strip().rstrip("\r\x07")
                    else:
                        desc_text = ""

                    # Столбец 5 (индекс 6 в COM) — марка стали
                    steel = ""
                    if cols_count >= 6:
                        steel = table.Cell(r, 6).Range.Text.strip().rstrip("\r\x07")

                    # Извлекаем Dy и Py из описания
                    dy_match = re.search(r'[DdДд][yуYУ]\s*=?\s*(\d+)', desc_text)
                    py_match = re.search(r'[PpРр][yуYУ]\s*=?\s*(\d+)', desc_text)

                    dy_val = dy_match.group(1) if dy_match else ""
                    py_val = py_match.group(1) if py_match else ""

                    result[num_clean] = {
                        "dy": dy_val,
                        "py": py_val,
                        "steel": steel
                    }
                except Exception:
                    continue

            self.log.emit(f"   ✓ [Спецификация] Извлечено записей: {len(result)}")
        except Exception as e:
            self.log.emit(f"   ⚠ [Спецификация] Ошибка чтения: {e}")
        finally:
            if spec_doc:
                try:
                    spec_doc.Close(SaveChanges=False)
                except Exception:
                    pass
        return result

    def _fill_report_table(self, doc, spec_data: dict) -> int:
        """
        Заполняет столбцы 3, 4, 5 в таблице 'Таблица результатов обследования'.
        Возвращает количество заполненных строк.
        """
        filled = 0
        target_table = None

        # Ищем таблицу после текста "Таблица результатов обследования"
        try:
            find_obj = doc.Content.Find
            find_obj.ClearFormatting()
            found = find_obj.Execute(FindText="Таблица результатов обследования")

            if not found:
                self.log.emit(f"   ⚠ [Отчет] Текст 'Таблица результатов обследования' не найден")
                return 0

            # Получаем позицию найденного текста и ищем первую таблицу после неё
            found_range = find_obj.Parent
            found_pos = found_range.End

            for t_idx in range(1, doc.Tables.Count + 1):
                tbl = doc.Tables(t_idx)
                if tbl.Range.Start >= found_pos:
                    target_table = tbl
                    break

            if not target_table:
                self.log.emit(f"   ⚠ [Отчет] Таблица после 'Таблица результатов обследования' не найдена")
                return 0

            rows_count = target_table.Rows.Count
            for r in range(1, rows_count + 1):
                try:
                    # Столбец 1 (индекс 2 в COM) — номер для поиска
                    cell_text = target_table.Cell(r, 2).Range.Text.strip().rstrip("\r\x07")
                    num_key = cell_text.strip()

                    if num_key in spec_data:
                        data = spec_data[num_key]
                        # Столбец 3 (индекс 4 в COM) — Dy
                        if data["dy"]:
                            target_table.Cell(r, 4).Range.Text = data["dy"]
                        # Столбец 4 (индекс 5 в COM) — Py
                        if data["py"]:
                            target_table.Cell(r, 5).Range.Text = data["py"]
                        # Столбец 5 (индекс 6 в COM) — Марка стали
                        if data["steel"]:
                            target_table.Cell(r, 6).Range.Text = data["steel"]
                        filled += 1
                except Exception:
                    continue

        except Exception as e:
            self.log.emit(f"   ⚠ [Отчет] Ошибка заполнения таблицы: {e}")

        return filled

    def _fix_resource_table(self, doc) -> int:
        """
        Исправляет таблицы 'Оценка остаточного ресурса' и 'Таблица Г.3':
        1. Корректирует столбец 1 (S) по числу после 'x' в столбце 0
        2. Пересчитывает скорость коррозии (стб.5) и остаточный ресурс (стб.6)
        Возвращает количество исправленных таблиц.
        """
        fixed_tables = 0
        search_texts = ["Оценка остаточного ресурса", "Таблица Г.3"]

        for search_text in search_texts:
            try:
                # Создаём новый Range для каждого поиска
                search_range = doc.Content.Duplicate
                search_range.Find.ClearFormatting()
                found = search_range.Find.Execute(FindText=search_text)

                if not found:
                    self.log.emit(f"   ℹ [Ресурс] Текст '{search_text}' не найден")
                    continue

                found_pos = search_range.End

                # Ищем первую таблицу после найденного текста
                target_table = None
                for t_idx in range(1, doc.Tables.Count + 1):
                    tbl = doc.Tables(t_idx)
                    if tbl.Range.Start >= found_pos:
                        target_table = tbl
                        break

                if not target_table:
                    self.log.emit(f"   ⚠ [Ресурс] Таблица после '{search_text}' не найдена")
                    continue

                rows_count = target_table.Rows.Count
                rows_fixed = 0

                for r in range(2, rows_count + 1):  # Пропускаем заголовок
                    try:
                        # --- Столбец 0 (COM idx 1): Наименование участка ---
                        name_text = target_table.Cell(r, 1).Range.Text.strip().rstrip("\r\x07")

                        # Извлекаем число после 'x' (латиница/кириллица)
                        x_match = re.search(r'[xхXХ](\d+[,.]\d+|\d+)', name_text)
                        if not x_match:
                            continue

                        correct_s_float = float(x_match.group(1).replace(',', '.'))

                        # --- Столбец 1 (COM idx 2): S — исполнительная толщина ---
                        current_s_text = target_table.Cell(r, 2).Range.Text.strip().rstrip("\r\x07")
                        try:
                            current_s = float(current_s_text.replace(',', '.'))
                        except ValueError:
                            current_s = 0.0

                        if abs(current_s - correct_s_float) > 0.01:
                            new_s_str = f"{correct_s_float:.1f}".replace('.', ',')
                            target_table.Cell(r, 2).Range.Text = new_s_str
                            self.log.emit(f"   [Ресурс] Строка {r}: S исправлено {current_s_text} -> {new_s_str}")
                            current_s = correct_s_float

                        # --- Читаем остальные столбцы ---
                        sf_text = target_table.Cell(r, 3).Range.Text.strip().rstrip("\r\x07")
                        si_text = target_table.Cell(r, 4).Range.Text.strip().rstrip("\r\x07")
                        t_text = target_table.Cell(r, 5).Range.Text.strip().rstrip("\r\x07")

                        sf = float(sf_text.replace(',', '.'))
                        si = float(si_text.replace(',', '.'))
                        t_val = float(t_text.replace(',', '.'))

                        # --- Определяем значение S для расчёта ---
                        # Если исполнительная (после коррекции) совпадает с фактической,
                        # добавляем 0.5 только для расчёта (в таблице не меняем)
                        if abs(current_s - sf) < 0.01:
                            calc_s = current_s + 0.5
                        else:
                            calc_s = current_s

                        # --- Расчёт скорости коррозии ---
                        if t_val > 0:
                            speed = (calc_s - sf) / t_val
                        else:
                            speed = 0.0

                        # --- Запись столбца 5 (скорость коррозии, 3 знака) ---
                        if speed <= 0:
                            target_table.Cell(r, 6).Range.Text = "-"
                            target_table.Cell(r, 7).Range.Text = "-"
                        else:
                            speed_str = f"{speed:.3f}".replace('.', ',')
                            target_table.Cell(r, 6).Range.Text = speed_str

                            # --- Расчёт остаточного ресурса ---
                            resource = (sf - si) / speed
                            if resource <= 0:
                                target_table.Cell(r, 7).Range.Text = "-"
                            else:
                                # Округляем вниз до целого числа
                                resource_int = math.floor(resource)
                                target_table.Cell(r, 7).Range.Text = str(resource_int)

                        rows_fixed += 1
                    except Exception as e:
                        self.log.emit(f"   ⚠ [Ресурс] Ошибка в строке {r}: {e}")
                        continue

                self.log.emit(f"   ✓ [Ресурс] Таблица '{search_text}': обработано {rows_fixed} строк")
                fixed_tables += 1

            except Exception as e:
                self.log.emit(f"   ⚠ [Ресурс] Ошибка обработки '{search_text}': {e}")

        return fixed_tables

    def _fix_attestation_certificate(self, doc) -> bool:
        """
        Заменяет номер свидетельства об аттестации ЛНК в зависимости от даты акта.
        Порог: 20.03.2026.
        До — №ЛНК-057А0459 (до 20 марта 2026 г.), области с "2.1.,"
        После — №ЛНК-057А0516 (до 20 марта 2029 г.), области с "2.1,"
        Возвращает True если замены были выполнены.
        """
        OLD_CERT = "ЛНК-057А0459"
        NEW_CERT = "ЛНК-057А0516"
        THRESHOLD = datetime(2026, 3, 20)

        # Словарь русских месяцев для парсинга даты
        MONTHS_RU = {
            "января": 1, "февраля": 2, "марта": 3, "апреля": 4,
            "мая": 5, "июня": 6, "июля": 7, "августа": 8,
            "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12
        }

        # --- Шаг 1: Найти дату акта ---
        act_date = None
        try:
            # Ищем текст вида "АКТ №... от DD месяц YYYY г." по всем абзацам
            for para in doc.Paragraphs:
                para_text = para.Range.Text
                if "АКТ" not in para_text or " от " not in para_text:
                    continue
                # Извлекаем дату: "от DD месяц YYYY г."
                date_match = re.search(
                    r'от\s+(\d{1,2})\s+([а-яё]+)\s+(\d{4})\s+г\.',
                    para_text, re.IGNORECASE
                )
                if date_match:
                    day = int(date_match.group(1))
                    month_name = date_match.group(2).lower()
                    year = int(date_match.group(3))
                    month = MONTHS_RU.get(month_name, 0)
                    if month > 0:
                        act_date = datetime(year, month, day)
                        self.log.emit(f"   [Свидетельство] Дата акта: {act_date.strftime('%d.%m.%Y')}")
                        break
        except Exception as e:
            self.log.emit(f"   ⚠ [Свидетельство] Ошибка поиска даты: {e}")

        if act_date is None:
            self.log.emit(f"   ⚠ [Свидетельство] Дата акта не найдена, пропускаем замену")
            return False

        is_before = act_date < THRESHOLD
        made_changes = False

        # --- Шаг 2: Замена номера свидетельства и даты действия ---
        try:
            if not is_before:
                # Дата >= 20.03.2026 → меняем номер и дату действия
                # Замена номера по всему документу (покрывает все листы актов)
                find_obj = doc.Content.Find
                find_obj.ClearFormatting()
                find_obj.Replacement.ClearFormatting()
                find_obj.Execute(
                    OLD_CERT,       # FindText
                    False,          # MatchCase
                    False,          # MatchWholeWord
                    False,          # MatchWildcards
                    False,          # MatchSoundsLike
                    False,          # MatchAllWordForms
                    True,           # Forward
                    1,              # Wrap = wdFindContinue
                    False,          # Format
                    NEW_CERT,       # ReplaceWith
                    2               # Replace = wdReplaceAll
                )
                self.log.emit(f"   [Свидетельство] Номер заменён: {OLD_CERT} → {NEW_CERT}")
                made_changes = True

                # Замена даты действия
                find_obj2 = doc.Content.Find
                find_obj2.ClearFormatting()
                find_obj2.Replacement.ClearFormatting()
                find_obj2.Execute(
                    "до 20 марта 2026 г.",
                    False, False, False, False, False,
                    True, 1, False,
                    "до 20 марта 2029 г.",
                    2
                )
                self.log.emit(f"   [Свидетельство] Дата действия: 2026 → 2029")
                made_changes = True

                self.log.emit(f"   ✓ [Свидетельство] Применён новый сертификат (дата акта ≥ 20.03.2026)")
            else:
                self.log.emit(f"   ✓ [Свидетельство] Оставлен старый сертификат (дата акта < 20.03.2026)")

        except Exception as e:
            self.log.emit(f"   ⚠ [Свидетельство] Ошибка замены: {e}")

        return made_changes

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
                blocked, reason = is_point_blocked(p_folder)
                if blocked:
                    self.log.emit(f"\n── {p_name} ──\n   ⏭ Пропущен: {reason}")
                    continue
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

                    # Читаем F9/B17 с листа "паспорт" для имени итогового отчёта
                    report_f9 = report_b17 = None
                    try:
                        rn = CONFIG.get("REPORT_NAME", {})
                        ps = work_wb.Sheets(rn.get("SHEET", "паспорт"))
                        report_f9 = ps.Range(rn.get("TTP_CELL", "F9")).Value
                        report_b17 = ps.Range(rn.get("POINT_CELL", "B17")).Value
                    except Exception as e:
                        self.log.emit(f"   ⚠ Не удалось прочитать F9/B17 для имени отчёта: {e}")

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
                        
                        # Сторож: если привязка встала, в журнале должно быть
                        # видно, какое окно висит, — иначе этап выглядит просто
                        # зависшим (так и было, пока диалог ждал фокуса).
                        watchdog = self._merge_watchdog(wd_dk.word_pid)
                        try:
                            doc.MailMerge.OpenDataSource(
                                Name=clean_source_path,
                                SQLStatement="SELECT * FROM [экспорт$]",
                                SubType=1 # SubTypeAccess
                            )
                        finally:
                            watchdog.set()


                        # "Заморозка" данных
                        doc.Fields.Update()
                        doc.MailMerge.ViewMailMergeFieldCodes = False
                        doc.Fields.Unlink()
                        doc.MailMerge.MainDocumentType = -1
                        
                        out_name = build_report_name(report_f9, report_b17, pt_num_str)
                        out_path = os.path.join(p_folder, out_name)
                        doc.SaveAs2(out_path, 16)

                        # Убираем устаревший файл со старым именем, если остался от прошлого прогона
                        if out_name.startswith("ЗЭПБ"):
                            old_default = os.path.join(p_folder, f"Отчет_{pt_num_str}.docx")
                            if os.path.exists(old_default) and os.path.abspath(old_default) != os.path.abspath(out_path):
                                try:
                                    os.remove(old_default)
                                except Exception:
                                    pass

                        self.log.emit(f"   ✓ [Word] Отчет готов: {out_name}")
                        point_report["Слияние Word"] = True

                        # --- ШАГ 5: Заполнение таблицы из Спецификации ---
                        spec_path = self._find_specification_file(p_folder)
                        if spec_path:
                            self.action_update.emit(f"[{pt_num_str}] Заполнение из спецификации...")
                            self.log.emit(f"   [Спецификация] Найден файл: {os.path.basename(spec_path)}")
                            spec_data = self._parse_specification(wd_app, spec_path)
                            if spec_data:
                                filled_count = self._fill_report_table(doc, spec_data)
                                self.log.emit(f"   ✓ [Отчет] Заполнено строк из спецификации: {filled_count}")
                                if filled_count > 0:
                                    doc.Save()
                            else:
                                self.log.emit(f"   ⚠ [Спецификация] Не удалось извлечь данные")
                        else:
                            self.log.emit(f"   ℹ [Спецификация] Файл не найден в Первичке, пропускаем")

                        # --- ШАГ 6: Перерасчёт таблицы остаточного ресурса ---
                        self.action_update.emit(f"[{pt_num_str}] Перерасчёт ресурса...")
                        fixed_count = self._fix_resource_table(doc)
                        if fixed_count > 0:
                            doc.Save()
                            self.log.emit(f"   ✓ [Ресурс] Пересчитано таблиц: {fixed_count}")

                        # --- ШАГ 7: Замена свидетельства об аттестации ЛНК ---
                        self.action_update.emit(f"[{pt_num_str}] Проверка свидетельства ЛНК...")
                        cert_changed = self._fix_attestation_certificate(doc)
                        if cert_changed:
                            doc.Save()

                        self.info.emit(f"{p_name} — отчет создан", "done")
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

                # --- Строка «Сведения о ремонте» ---
                # COM-объекты закрыты выше, файлы разблокированы: читаем реестр
                # (лист «Ремонт») и вписываем строку в готовый отчёт. Модуль
                # «Ремонт» опционален — его отсутствие не должно ломать этап.
                try:
                    from repair.registry import read_registry
                    from repair.svedenia import build_svedenia
                    from repair.word_writer import write_svedenia
                    xlsm_p = find_file(p_folder, (".xls", ".xlsx", ".xlsm"))
                    records = read_registry(xlsm_p) if xlsm_p else []
                    if records:
                        report_p = find_report_file(p_folder)
                        if report_p and write_svedenia(report_p, build_svedenia(records)):
                            self.log.emit("   ✓ [Ремонт] Строка «Сведения о ремонте» "
                                          "записана в отчёт")
                        elif report_p:
                            self.log.emit("   ⚠ [Ремонт] Строка «Сведения о ремонте» "
                                          "в отчёте не найдена")
                except Exception as e:
                    self.log.emit(f"   ⚠ [Ремонт] Строку «Сведения» записать не удалось: {e}")

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

