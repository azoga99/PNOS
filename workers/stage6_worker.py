# -*- coding: utf-8 -*-
"""
Этап 7: Вставка актов гидроиспытаний (ГИ) в итоговый отчёт.

Что делает для каждого пункта (Отчет_*.docx, уже прошедший этап 6):
  1. Берёт номер участка (папку) из паспорт!B12 — кэш этапа 3 (`gi_number`),
     либо фолбэком читает xlsm через openpyxl.
  2. Скачивает JPG-акты из /ПНОС/<число>/ на Я.Диске (аккаунт шаблонов TOKEN_BASE),
     один раз на каждое уникальное число, в кэш <local_root>/_gi_cache/<число>/.
  3. Вставляет акты по порядку (по хвостовому числу в имени файла), по одному
     на страницу, начиная с метки [ГИ]. Ориентация приводится по EXIF-тегу
     (сканы часто хранятся с "лежащими" пикселями + тегом поворота — если
     не применить его явно, Word вставит картинку как есть, боком).
  4. Полностью пересчитывает номера страниц на листе СОДЕРЖАНИЕ по фактическим
     страницам финального документа (это заодно чинит сдвиг от фото этапа 6).

Если папки актов для числа нет — не ошибка: маркер удаляется, в отчёт пишется
«нет актов ГИ для <число>», содержание всё равно пересчитывается.
"""

import os
import re
import asyncio
import tempfile

import pythoncom
import win32com.client
import win32process
import aiohttp
from PIL import Image, ImageOps
from PySide6.QtCore import QThread, Signal

from config import CONFIG
from status_manager import get_stage_status, set_stage_status, is_point_blocked
from yandex_api import YandexDiskAPI
from workers.stage3_worker import normalize_gi_number, find_file
from workers.stage4_worker import WordDialogKiller, find_report_file

# Word COM константы
WD_ACTIVE_END_PAGE_NUMBER = 3  # wdActiveEndPageNumber
WD_WRAP_FRONT = 3              # wdWrapFront
WD_SHAPE_CENTER = -999995     # значение Left/Top = центрирование (wdShapeCenter)
# Габариты вписывания акта (пункты) — как у схем [ФОТО] на A4-странице. Даёт размер
# ~520×735, совпадающий с эталоном оформления; при wdShapeCenter акт центрируется
# в области печати страницы.
GI_MAX_W, GI_MAX_H = 520.0, 750.0


def gi_sort_key(path: str) -> int:
    """Ключ сортировки актов по хвостовому числу имени: '..._page-0010.jpg' -> 10."""
    name = os.path.basename(path)
    stem = os.path.splitext(name)[0]
    m = re.search(r'(\d+)\s*$', stem)
    return int(m.group(1)) if m else 0


def _toc_key(text: str):
    """Ключ сопоставления строки содержания и заголовка в тексте.
    Устойчив к «грязи» шаблона (двойные номера, лишние слова).
    Возвращает 'ПРИЛ:<буква>' | 'SEC:<номер>:<ПЕРВОЕ_СЛОВО>' | None.
    """
    t = (text or "").strip()
    m = re.match(r'^ПРИЛОЖЕНИЕ\s+([А-ЯЁA-Z])', t, re.IGNORECASE)
    if m:
        return "ПРИЛ:" + m.group(1).upper()
    m = re.match(r'^(\d+(?:\.\d+)*)\b', t)
    if not m:
        return None
    num = m.group(1)
    wm = re.search(r'[A-Za-zА-Яа-яЁё]+', t[m.end():])
    if not wm:
        return None
    return "SEC:" + num + ":" + wm.group(0).upper()


class Stage6Worker(QThread):
    """Этап 7 (UI): Вставка актов ГИ и пересчёт содержания."""
    log = Signal(str)
    progress = Signal(int)
    finished_ok = Signal(bool)
    info = Signal(str, str)
    action_update = Signal(str)
    report_ready = Signal(list)

    def __init__(self, local_root: str, parent=None):
        super().__init__(parent)
        self.local_root = local_root
        self.wd_app = None
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    # ─── Определение номера участка ──────────────────────────────────

    def _get_gi_number(self, p_folder: str) -> str:
        """Берёт номер участка из статуса (этап 3), иначе читает xlsm через openpyxl."""
        num = get_stage_status(p_folder, "gi_number")
        if isinstance(num, str) and num.strip():
            return num.strip()
        return self._read_b12_from_xlsm(p_folder)

    def _read_b12_from_xlsm(self, p_folder: str) -> str:
        """Фолбэк: читает паспорт!B12 из xlsm без COM (openpyxl, кэшированные значения)."""
        try:
            import openpyxl
            xlsm = find_file(p_folder, (".xlsm", ".xlsx", ".xls"))
            if not xlsm:
                return ""
            cell = CONFIG.get("GI_CONFIG", {}).get("PASSPORT_CELL", "B12")
            wb = openpyxl.load_workbook(xlsm, data_only=True, read_only=True)
            val = None
            for name in wb.sheetnames:
                if name.strip().lower() == "паспорт":
                    val = wb[name][cell].value
                    break
            wb.close()
            return normalize_gi_number(val)
        except Exception as e:
            self.log.emit(f"   [ГИ] ⚠ Фолбэк чтения B12 из xlsm не удался: {e}")
            return ""

    # ─── Скачивание актов с Я.Диска (async) ──────────────────────────

    async def _download_all_gi(self, numbers: set) -> dict:
        """numbers: множество чисел. Возвращает dict число -> [пути к jpg] или None."""
        result = {}
        api = YandexDiskAPI(CONFIG["TOKEN_BASE"])
        api.set_log_callback(self.log.emit)
        gi_cfg = CONFIG.get("GI_CONFIG", {})
        base = gi_cfg.get("DISK_BASE", CONFIG.get("DISK_PATH_BASE", "/ПНОС"))
        cache_root = os.path.join(self.local_root, "_gi_cache")

        async with aiohttp.ClientSession() as session:
            for num in numbers:
                if self._is_cancelled:
                    break
                folder_path = await api.find_folder_by_name(session, base, num)
                if not folder_path:
                    self.log.emit(f"   [ГИ] На диске нет папки актов для «{num}».")
                    result[num] = None
                    continue

                items = await api.get_folder_contents(session, folder_path)
                jpgs = [it for it in (items or [])
                        if it.get("type") == "file"
                        and it["name"].lower().endswith((".jpg", ".jpeg", ".png"))]
                if not jpgs:
                    self.log.emit(f"   [ГИ] Папка «{num}» пуста (нет изображений).")
                    result[num] = None
                    continue

                local_dir = os.path.join(cache_root, num)
                os.makedirs(local_dir, exist_ok=True)
                dl_tasks, local_paths = [], []
                for it in jpgs:
                    lp = os.path.join(local_dir, it["name"])
                    local_paths.append(lp)
                    if not os.path.exists(lp):
                        dl_tasks.append(api.download_file(session, it["path"], lp))
                if dl_tasks:
                    await asyncio.gather(*dl_tasks)

                existing = sorted([p for p in local_paths if os.path.exists(p)], key=gi_sort_key)
                result[num] = existing or None
                self.log.emit(f"   [ГИ] «{num}»: изображений готово {len(existing)}.")
        return result

    # ─── Вставка актов в Word ────────────────────────────────────────

    def _insert_gi(self, doc, images) -> int:
        """Вставляет акты по одному на страницу, начиная с метки [ГИ].
        Метод — дублирование страницы метки (как для [ФОТО]). Возвращает число вставленных.
        """
        marker = CONFIG.get("GI_CONFIG", {}).get("MARKER", "[ГИ]")

        find_range = doc.Content
        find_range.Find.Execute(FindText=marker)
        if not find_range.Find.Found:
            self.log.emit(f"   ⚠ Маркер {marker} не найден в отчёте.")
            return 0

        if not images:
            find_range.Text = ""  # актов нет — просто убираем метку
            return 0

        # 1. Диапазон страницы, где стоит метка — эталон для дублирования
        try:
            marker_range = find_range.Duplicate
            marker_range.Select()
            page_range = self.wd_app.Selection.Bookmarks("\\page").Range
            page_range.Copy()

            # 2. Дублируем страницу под каждый дополнительный акт
            insert_pos = page_range.End
            for _ in range(len(images) - 1):
                r_paste = doc.Range(insert_pos, insert_pos)
                r_paste.Paste()
                insert_pos = r_paste.End
        except Exception as e:
            self.log.emit(f"   ❌ Не удалось размножить страницу под акты ГИ: {e}")
            return 0

        # 3. Каждую копию метки заменяем на изображение (по порядку)
        inserted = 0
        for img_path in images:
            fresh = doc.Content
            fresh.Find.Execute(FindText=marker)
            if not fresh.Find.Found:
                self.log.emit("   ⚠ Не найдена область для очередного акта.")
                break
            anchor = fresh.Duplicate
            anchor.Text = " "
            anchor.Collapse(1)

            temp_path = None
            try:
                with Image.open(img_path) as im_raw:
                    orig_format = im_raw.format or "JPEG"
                    # Применяем EXIF Orientation физически: сканы часто содержат
                    # "лежащие" пиксели + тег поворота, который Word игнорирует при
                    # AddPicture (в отличие от проводника/просмотрщиков) — без этого
                    # акт вставляется боком (альбомно) вместо вертикали.
                    im = ImageOps.exif_transpose(im_raw)
                    ow, oh = im.size
                    ratio = min(GI_MAX_W / ow, GI_MAX_H / oh)
                    tw, th = ow * ratio, oh * ratio
                    temp_path = os.path.join(tempfile.gettempdir(), f"gi_{os.path.basename(img_path)}")
                    im.save(temp_path, format=orig_format)

                shape = doc.Shapes.AddPicture(temp_path, False, True, 0, 0, tw, th, anchor)
                shape.WrapFormat.Type = WD_WRAP_FRONT
                shape.LockAspectRatio = True
                # Размер и центрирование — как у схем [ФОТО] (эталон оформления):
                # акт вписывается в рамку A4-страницы и центрируется в области печати.
                shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin
                shape.Left = WD_SHAPE_CENTER
                shape.RelativeVerticalPosition = 0
                shape.Top = WD_SHAPE_CENTER
                inserted += 1
                self.log.emit(f"   ✓ Вставлен акт: {os.path.basename(img_path)}")
            except Exception as e:
                self.log.emit(f"   ❌ Ошибка вставки {os.path.basename(img_path)}: {e}")
            finally:
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass

        # Подчищаем остатки метки, если что-то осталось
        while True:
            cleanup = doc.Content
            if cleanup.Find.Execute(FindText=marker):
                cleanup.Text = " "
            else:
                break
        return inserted

    # ─── Пересчёт содержания ─────────────────────────────────────────

    def _recalc_toc(self, doc) -> int:
        """Полностью пересчитывает номера страниц в СОДЕРЖАНИИ по фактическим страницам."""
        try:
            doc.Repaginate()
        except Exception:
            pass

        toc_entries = []      # [(paragraph, key)]
        key_to_page = {}      # key -> реальная страница
        # Зона содержания открывается СТРОГО по заголовку "СОДЕРЖАНИЕ", а не с начала
        # документа: на титульном листе перед ним попадаются регистрационные номера
        # и даты вида "59-ЗС-...-2026" / "21 апреля 2026 г.", которые ложно похожи на
        # заголовок раздела (начинаются с цифр, дальше есть буквы) и без этого якоря
        # преждевременно "закрывали" зону раньше, чем парсер доходил до реального
        # блока СОДЕРЖАНИЕ — из-за чего пересчёт не находил ни одной строки и молча
        # ничего не менял.
        in_toc_zone = False
        seen_anchor = False

        for para in doc.Paragraphs:
            try:
                raw = para.Range.Text
            except Exception:
                continue
            txt = raw[:-1] if raw.endswith('\r') else raw
            if not seen_anchor:
                if txt.strip().upper() == "СОДЕРЖАНИЕ":
                    seen_anchor = True
                    in_toc_zone = True
                continue
            key = _toc_key(txt)
            if key is None:
                continue
            is_toc_line = re.match(r'^.*\t\d+\s*$', txt, re.S) is not None
            if is_toc_line and in_toc_zone:
                toc_entries.append((para, key))
            elif not is_toc_line:
                # первый настоящий заголовок в теле закрывает зону содержания
                in_toc_zone = False
                if key not in key_to_page:
                    try:
                        key_to_page[key] = int(para.Range.Information(WD_ACTIVE_END_PAGE_NUMBER))
                    except Exception:
                        pass

        if not seen_anchor:
            self.log.emit("   ⚠ Заголовок «СОДЕРЖАНИЕ» не найден — пересчёт страниц пропущен.")

        updated = 0
        for para, key in toc_entries:
            page = key_to_page.get(key)
            if page is not None and self._set_toc_page(para, page):
                updated += 1
        return updated

    def _set_toc_page(self, para, new_page: int) -> bool:
        """Заменяет только хвостовой номер в строке содержания, сохраняя абзац и табстопы."""
        try:
            rng = para.Range
            raw = rng.Text
            has_pilcrow = raw.endswith('\r')
            body_txt = raw[:-1] if has_pilcrow else raw
            m = re.match(r'^(.*\t)(\d+)(\s*)$', body_txt, re.S)
            if not m:
                return False
            if m.group(2) == str(new_page):
                return False  # уже верно
            new_body = m.group(1) + str(new_page) + m.group(3)
            rng2 = para.Range
            if has_pilcrow:
                rng2.End = rng2.End - 1  # не трогаем знак абзаца
            rng2.Text = new_body
            return True
        except Exception as e:
            self.log.emit(f"   ⚠ Не удалось обновить строку содержания: {e}")
            return False

    # ─── Основной цикл ───────────────────────────────────────────────

    def run(self):
        self.log.emit("═" * 40)
        self.log.emit("ЭТАП 7: Вставка актов ГИ")
        self.log.emit("═" * 40)

        if not os.path.isdir(self.local_root):
            self.log.emit("❌ Корневая папка не найдена.")
            self.finished_ok.emit(False)
            return

        point_folders = [os.path.join(self.local_root, d) for d in os.listdir(self.local_root)
                         if os.path.isdir(os.path.join(self.local_root, d)) and d.startswith("п.")]
        if not point_folders:
            self.log.emit("⚠ Папки пунктов не найдены.")
            self.finished_ok.emit(False)
            return

        # 1. Определяем номер участка (B12) для каждого пункта
        self.info.emit("Определение участков актов ГИ...", "wait")
        point_numbers = {p: self._get_gi_number(p) for p in point_folders}
        unique_numbers = {n for n in point_numbers.values() if n}
        self.progress.emit(5)

        # 2. Скачиваем акты (один раз на каждое уникальное число)
        gi_images = {}
        if unique_numbers:
            self.info.emit(f"Скачивание актов ГИ ({len(unique_numbers)} участк.)...", "wait")
            self.log.emit(f"\n📥 Участки для скачивания: {', '.join(sorted(unique_numbers))}")
            try:
                gi_images = asyncio.run(self._download_all_gi(unique_numbers))
            except Exception as e:
                self.log.emit(f"❌ Ошибка скачивания актов ГИ: {e}")
                gi_images = {}
        else:
            self.log.emit("⚠ Ни для одного пункта не определён номер участка (B12).")
        self.progress.emit(20)

        if self._is_cancelled:
            self.finished_ok.emit(False)
            return

        # 3. Word COM: вставка актов + пересчёт содержания
        pythoncom.CoInitialize()
        wd_dk = None
        report_data = []
        try:
            self.wd_app = win32com.client.DispatchEx("Word.Application")
            self.wd_app.Visible = False
            self.wd_app.DisplayAlerts = 0

            word_pid = 0
            try:
                _, word_pid = win32process.GetWindowThreadProcessId(self.wd_app.Application.Hwnd)
            except Exception:
                pass
            wd_dk = WordDialogKiller(log_callback=self.log.emit, word_pid=word_pid)
            wd_dk.start()

            for i, p_folder in enumerate(point_folders):
                if self._is_cancelled:
                    break
                p_name = os.path.basename(p_folder)
                blocked, reason = is_point_blocked(p_folder)
                if blocked:
                    self.log.emit(f"\n── {p_name} ──\n   ⏭ Пропущен: {reason}")
                    continue
                pt_num = p_name[2:] if p_name.startswith("п.") else p_name
                row = {"№ Пункта": pt_num, "Акты ГИ": "—", "Результат": "Ожидание"}

                if get_stage_status(p_folder, "stage6"):
                    self.log.emit(f"\n── {p_name} ──\n   ⏭ Этап 7 уже завершён, пропускаем.")
                    row["Акты ГИ"] = "✔"
                    row["Результат"] = "⚪ Пропущен"
                    report_data.append(row)
                    self.progress.emit(20 + int((i + 1) / len(point_folders) * 80))
                    continue

                self.info.emit(f"Обработка {p_name}...", "wait")
                self.log.emit(f"\n── {p_name} ──")

                # Поиск файла отчёта (новое имя 'ЗЭПБ…' или старое 'Отчет_*')
                report_file = find_report_file(p_folder)
                if not report_file:
                    self.log.emit("   ❌ Файл отчёта не найден. Сначала выполните Этап 6.")
                    row["Результат"] = "❌ Нет отчета"
                    report_data.append(row)
                    self.progress.emit(20 + int((i + 1) / len(point_folders) * 80))
                    continue

                gi_num = point_numbers.get(p_folder, "")
                images = gi_images.get(gi_num) if gi_num else None

                doc = None
                try:
                    doc = self.wd_app.Documents.Open(report_file)

                    inserted = self._insert_gi(doc, images)
                    if images:
                        total = len(images)
                        if inserted == total:
                            row["Акты ГИ"] = str(inserted)
                            self.log.emit(f"   ✅ Вставлено актов ГИ: {inserted}")
                        elif inserted > 0:
                            # Часть актов вставилась, часть — нет (см. лог выше по причине)
                            row["Акты ГИ"] = f"{inserted}/{total}"
                            self.info.emit(f"Вставлена только часть актов ГИ ({inserted}/{total}) для {gi_num}", "warn")
                            self.log.emit(f"   ⚠ Вставлено только {inserted} из {total} актов ГИ.")
                        else:
                            # Акты найдены на диске, но вставить в документ не удалось —
                            # это не то же самое, что "актов нет", поэтому сообщение отдельное
                            row["Акты ГИ"] = f"ошибка вставки ({gi_num})"
                            self.info.emit(f"Не удалось вставить акты ГИ для {gi_num}", "warn")
                            self.log.emit(f"   ❌ Акты ГИ найдены ({total} шт.), но вставить в документ не удалось.")
                    elif gi_num:
                        row["Акты ГИ"] = f"нет папки для {gi_num}"
                        self.info.emit(f"Нет актов ГИ для {gi_num}", "warn")
                        self.log.emit(f"   ⚠ На диске нет папки актов для «{gi_num}».")
                    else:
                        row["Акты ГИ"] = "участок не задан"
                        self.log.emit("   ⚠ Не определён номер участка (B12).")

                    # Пересчёт содержания (всегда — фиксит и сдвиг от фото этапа 6)
                    updated = self._recalc_toc(doc)
                    self.log.emit(f"   ✓ Обновлено записей в СОДЕРЖАНИИ: {updated}")

                    doc.Save()
                    doc.Close()
                    doc = None

                    set_stage_status(p_folder, "stage6", True)
                    row["Результат"] = "✅ Успешно"
                    self.info.emit(f"{p_name} — готово", "done")
                except Exception as ex:
                    self.log.emit(f"   ❌ Ошибка Word: {ex}")
                    row["Результат"] = "❌ Ошибка Word"
                    if doc:
                        try:
                            doc.Close(False)
                        except Exception:
                            pass

                # --- Авто-верификация пункта с ремонтом (самый конец) ---
                # Здесь отчёт полностью готов (задвижки уже удалены на этапе 6),
                # поэтому проверка достоверна. Модуль «Ремонт» опционален.
                try:
                    from repair.detect import find_repair_docs
                    if find_repair_docs(p_folder).has_repair:
                        from repair.verifier import verify_point
                        findings = verify_point(p_folder)
                        errs = [f for f in findings if f.level == "error"]
                        warns = [f for f in findings if f.level == "warn"]
                        set_stage_status(p_folder, "repair_findings",
                                         [(f.level, f.code, f.message) for f in findings])
                        for f in errs + warns:
                            self.log.emit(f"   [Ремонт·проверка {f.code}] {f.message}")
                        if errs:
                            self.info.emit(f"{p_name}: ремонт — проверка нашла "
                                           f"{len(errs)} ошибок, см. журнал", "error")
                        elif warns:
                            self.info.emit(f"{p_name}: ремонт — {len(warns)} "
                                           f"предупреждений", "warn")
                        else:
                            self.info.emit(f"{p_name}: ремонт согласован", "done")
                except Exception as e:
                    self.log.emit(f"   ⚠ [Ремонт] Авто-проверка не выполнена: {e}")

                report_data.append(row)
                self.progress.emit(20 + int((i + 1) / len(point_folders) * 80))

            self.report_ready.emit(report_data)
            self.progress.emit(100)
            self.info.emit("Этап 7 завершён!", "done")
            self.finished_ok.emit(not self._is_cancelled)
        except Exception as e:
            import traceback
            self.log.emit(f"❌ Критическая ошибка: {e}")
            self.log.emit(traceback.format_exc())
            self.finished_ok.emit(False)
        finally:
            if wd_dk:
                wd_dk.stop()
            if self.wd_app:
                try:
                    self.wd_app.Quit()
                except Exception:
                    pass
                self.wd_app = None
            pythoncom.CoUninitialize()
