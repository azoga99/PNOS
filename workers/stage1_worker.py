# -*- coding: utf-8 -*-
"""
Этап 1: Создание структуры на локальном ПК.
- Создание папки ПНОС на рабочем столе
- Пункты, которые сотрудник сам сложил в ПНОС (Паспорт и Первичка на месте),
  берутся как есть — без поиска на Яндекс Диске
- Остальные ищутся и скачиваются с Яндекс Диска (асинхронно, aiohttp)
- Шаблоны ПНОС_ТТП.docx / ПНОС_ТТП.xlsm — с диска шаблонов; копия хранится
  в ПНОС\\_шаблоны и выручает, когда диск недоступен
"""

import os
import re
import shutil
import asyncio
from datetime import datetime

import aiohttp

from PySide6.QtCore import QThread, Signal

from config import CONFIG
from yandex_api import YandexDiskAPI
from excel_service import analyze_excel
from status_manager import (get_stage_status, set_stage_status, get_status_value,
                            reset_stage_status)

# Копия шаблонов внутри ПНОС. Этапы 2–7 берут только папки «п.*», её не тронут
TEMPLATE_CACHE = "_шаблоны"

# «п.3186», «п. 3186», «П.3186», «пункт 3186», «3186», «п.3186 (не корректный)»
_POINT_DIR = re.compile(r"^\s*(?:(?:пункт|п)\.?\s*)?(\d+)(?!\d)", re.IGNORECASE)


def point_number_of(dirname: str) -> str | None:
    """Номер пункта из имени папки или None, если имя на пункт не похоже."""
    m = _POINT_DIR.match(dirname)
    return m.group(1) if m else None


def _has_files(folder: str) -> bool:
    return os.path.isdir(folder) and bool(os.listdir(folder))


def missing_required(point_folder: str) -> list[str]:
    """Обязательные папки (REQUIRED_FOLDERS), которых в пункте нет или они пусты."""
    return [r for r in CONFIG.get("REQUIRED_FOLDERS", ["Паспорт", "Первичка"])
            if not _has_files(os.path.join(point_folder, r))]


def find_local_point(local_root: str, p_num: str) -> str | None:
    """
    Папка пункта в ПНОС под любым привычным именем. Если таких несколько,
    берётся полная (Паспорт и Первичка на месте), среди равных — «п.N».
    Иначе неполная «п.N» заслонила бы собранный вручную «п.N (не корректный)».
    """
    try:
        names = sorted(os.listdir(local_root))
    except OSError:
        return None
    exact = f"п.{p_num}"
    found = [os.path.join(local_root, n) for n in names
             if os.path.isdir(os.path.join(local_root, n)) and point_number_of(n) == p_num]
    if not found:
        return None
    found.sort(key=lambda f: (bool(missing_required(f)), os.path.basename(f) != exact))
    return found[0]


def normalize_point_folder(local_root: str, folder: str, p_num: str) -> tuple[str, str | None]:
    """
    Приводит имя папки к «п.N»: этапы 2–7 берут номер из имени (всё после «п.»),
    а приписка «(не корректный)» блокирует пункт. Возвращает (папка, ошибка).
    """
    target = os.path.join(local_root, f"п.{p_num}")
    if os.path.normcase(os.path.abspath(folder)) == os.path.normcase(os.path.abspath(target)):
        return folder, None
    old = os.path.basename(folder)
    if os.path.exists(target):
        return folder, f"папка «п.{p_num}» уже есть, «{old}» не переименована — уберите лишнюю"
    try:
        os.rename(folder, target)
    except OSError as e:
        return folder, (f"не удалось переименовать «{old}» в «п.{p_num}» "
                        f"(папка открыта в проводнике или в Office?): {e}")
    return target, None


def clear_invalid_mark(point_folder: str) -> bool:
    """
    Снимает пометку «не корректный» из статуса. Черновик шага 3, отметки
    ремонта и пройденные этапы остаются. True — пометка была.
    """
    had = bool(get_stage_status(point_folder, "invalid")) or \
        get_status_value(point_folder, "missing_folders") is not None
    if had:
        reset_stage_status(point_folder, "invalid")
        reset_stage_status(point_folder, "missing_folders")
    return had


class Stage1Worker(QThread):
    """Фоновый поток для Этапа 1 — создание структуры.
    Использует внутренний цикл asyncio для максимальной скорости сети.
    """

    # Сигналы
    log = Signal(str)             # Сообщение в лог
    progress = Signal(int)        # Прогресс 0-100
    finished_ok = Signal(bool)    # Завершение (True = успех)
    report_ready = Signal(dict)   # Данные для модального окна-отчёта
    info = Signal(str, str)       # Дружелюбный статус (сообщение, категория)

    def __init__(self, excel_path: str | None, local_root: str, manual_points: list[str] = None, need_epb: bool = True, max_threads: int = 10, parent=None):
        super().__init__(parent)
        self.excel_path = excel_path
        self.local_root = local_root
        self.manual_points = manual_points
        self.need_epb = need_epb
        # max_threads оставлено для обратной совместимости, хотя aiohttp сам управляет соединениями
        self.max_threads = max_threads
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        """Запуск из QThread. Запускаем asyncio loop."""
        asyncio.run(self.async_run())

    # ─── Шаблоны ────────────────────────────────────────────────────

    async def _prepare_templates(self, session, api_base, online: bool) -> dict:
        """
        Шаблоны в ПНОС\\_шаблоны: при связи — свежие с диска шаблонов (копия
        обновляется), без связи — последняя сохранённая копия. {имя: путь}.
        """
        self.log.emit(f"\n📥 Шаблоны (копия в {TEMPLATE_CACHE})...")
        cache = os.path.join(self.local_root, TEMPLATE_CACHE)
        os.makedirs(cache, exist_ok=True)
        fresh = set()

        if online:
            contents = await api_base.get_folder_contents(session, CONFIG["DISK_PATH_BASE"]) or []
            for item in contents:
                name = item["name"]
                if item["type"] != "file" or name not in CONFIG["TEMPLATE_FILES"]:
                    continue
                dest = os.path.join(cache, name)
                # Сначала во временный файл: оборванная загрузка не портит копию
                part = dest + ".part"
                if await api_base.download_file(session, item["path"], part):
                    try:
                        os.replace(part, dest)
                        fresh.add(name)
                        self.log.emit(f"   ↓ {name}")
                    except OSError as e:
                        self.log.emit(f"   ⚠ {name}: не удалось обновить копию ({e})")
                else:
                    self.log.emit(f"   ✗ Ошибка: {name}")
                if os.path.exists(part):
                    try:
                        os.remove(part)
                    except OSError:
                        pass

        templates = {}
        for name in CONFIG["TEMPLATE_FILES"]:
            path = os.path.join(cache, name)
            if not os.path.isfile(path):
                self.log.emit(f"   ✗ {name}: нет ни на диске шаблонов, ни в сохранённой копии")
                continue
            templates[name] = path
            if name not in fresh:
                stamp = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d.%m.%Y")
                self.log.emit(f"   ✓ {name}: сохранённая копия от {stamp}")
        return templates

    def _copy_templates(self, folder: str, p_num: str, templates: dict) -> list[str]:
        """Досылает недостающие шаблоны (готовые не трогает). Возвращает, каких нет."""
        for tpl_name, tpl_path in templates.items():
            dest = os.path.join(folder, tpl_name)
            if not os.path.exists(dest):
                shutil.copy2(tpl_path, dest)
                self.log.emit(f"   п.{p_num} 📄 {tpl_name} скопирован")
        return [n for n in CONFIG["TEMPLATE_FILES"] if not os.path.exists(os.path.join(folder, n))]

    @staticmethod
    def _folders_detail(folder: str) -> tuple[dict, dict]:
        """Что фактически лежит в папке пункта — для таблицы «Отчёт о прогоне»."""
        targets = {t: _has_files(os.path.join(folder, t)) for t in CONFIG["TARGET_FOLDERS"]}
        optional = {o: _has_files(os.path.join(folder, o)) for o in CONFIG.get("OPTIONAL_FOLDERS", [])}
        return targets, optional

    # ─── Свой комплект ──────────────────────────────────────────────

    def _accept_own_point(self, p_num: str, folder: str, templates: dict) -> dict:
        """Пункт, у которого Паспорт и Первичка уже лежат в ПНОС: только имя, статус и шаблоны."""
        suffix = CONFIG.get("INVALID_SUFFIX", " (не корректный)")
        was_done = (get_stage_status(folder, "stage1") and not get_stage_status(folder, "invalid")
                    and suffix.strip().lower() not in os.path.basename(folder).lower())
        self.log.emit(f"── Пункт п.{p_num} ──")

        old_name = os.path.basename(folder)
        folder, err = normalize_point_folder(self.local_root, folder, p_num)
        if err:
            self.log.emit(f"   ⚠ {err}")
            self.info.emit(f"п.{p_num}: {err}", "warn")
        elif os.path.basename(folder) != old_name:
            self.log.emit(f"   📁 Папка «{old_name}» переименована в «п.{p_num}»")

        if clear_invalid_mark(folder):
            self.log.emit("   ✓ Пометка «не корректный» снята — Паспорт и Первичка на месте")

        missing_tpl = self._copy_templates(folder, p_num, templates)
        set_stage_status(folder, "stage1", True)

        if was_done:
            self.log.emit("   ⏭ Этап 1 уже завершен, пропускаем загрузку.")
            status = "✓ Уже был готов"
        else:
            self.log.emit("   📂 Свой комплект: берётся как есть, с Яндекс Диска не скачивается")
            status = "✓ Свой комплект"
        if missing_tpl:
            status = f"⚠ Нет шаблонов: {', '.join(missing_tpl)}"
            self.info.emit(f"п.{p_num}: нет шаблонов {', '.join(missing_tpl)} — "
                           f"запустите этап 1 ещё раз, когда будет связь с Яндекс.Диском", "warn")
        if err:
            status = f"⚠ {err}"

        folders, optional = self._folders_detail(folder)
        return {"point": p_num, "status": status, "folders": folders, "optional": optional}

    # ─── Пункты с Яндекс Диска ──────────────────────────────────────

    async def _fetch_remote(self, session, api_items, points: list[str], templates: dict, report: dict):
        """Поиск и скачивание пунктов, которых в ПНОС нет или они неполные."""
        search_paths = CONFIG["DISK_PATHS_ITEMS"]
        self.log.emit(f"\n🔍 Поиск в {len(search_paths)} папках на Яндекс Диске...")
        self.info.emit(f"Ищу папки пунктов на Яндекс Диске (в {len(search_paths)} регионах)...", "wait")
        self.log.emit(f"   Искомые пункты: {', '.join(points[:5])}{'...' if len(points) > 5 else ''}")

        # Предзагрузка списков всех папок (ускоряет поиск в N раз)
        extended_search_paths = await api_items.preload_search_paths(session, search_paths)
        self.progress.emit(30)

        completed_checks = 0

        async def check_point(p_num):
            nonlocal completed_checks
            info = {
                "number": p_num,
                "remote_path": None,
                "target_paths": {},
                "optional_paths": {},
                "is_valid": False,
                "found_in": None,
                "missing": []
            }
            if self._is_cancelled:
                return info

            folder_path, found_in = await api_items.find_point_folder(session, extended_search_paths, p_num)
            info["remote_path"] = folder_path
            info["found_in"] = found_in

            if folder_path:
                # Проверка обязательных подпапок
                for target in CONFIG["TARGET_FOLDERS"]:
                    if not self.need_epb and target == "Стар. ЭПБ":
                        continue

                    target_path = await api_items.find_folder_by_name(session, folder_path, target)
                    info["target_paths"][target] = target_path
                    if not target_path:
                        info["missing"].append(target)

                # Проверка необязательных подпапок (Исполнительная, Акт отбраковки)
                for opt_folder in CONFIG.get("OPTIONAL_FOLDERS", []):
                    opt_path = await api_items.find_folder_by_name(session, folder_path, opt_folder)
                    info["optional_paths"][opt_folder] = opt_path
                    if opt_path:
                        self.log.emit(f"   п.{p_num} 📂 Найдена: {opt_folder}")

                info["is_valid"] = True  # Всегда скачиваем, если папка пункта существует

            completed_checks += 1
            pct = 30 + int(completed_checks / len(points) * 15)
            self.progress.emit(min(pct, 45))
            return info

        # Запускаем проверку ВСЕХ пунктов одновременно
        points_info = await asyncio.gather(*[check_point(num) for num in points])
        if self._is_cancelled:
            return

        for info in points_info:
            p_num = info["number"]
            if info["remote_path"]:
                report["found_on_disk"] += 1
                self.log.emit(f"   п.{p_num}: найден в {info['found_in']}")
                report["valid_points"] += 1
                if not info["missing"]:
                    self.log.emit(f"     ✓ Все обязательные папки найдены")
                else:
                    self.log.emit(f"     ⚠ Частично. Не найдены: {', '.join(info['missing'])}")
            else:
                self.log.emit(f"   п.{p_num}: НЕ НАЙДЕН ни в одной из {len(search_paths)} папок")

        self.info.emit("Анализ диска завершен.", "done")
        self.progress.emit(45)

        valid_points = [p for p in points_info if p["is_valid"]]
        invalid_points = [p for p in points_info if not p["is_valid"]]

        self.log.emit(f"\n📋 Итог анализа диска:")
        self.log.emit(f"   Искали на диске: {len(points)}")
        self.log.emit(f"   Найдено на диске: {report['found_on_disk']}")
        self.log.emit(f"   Будет скачано: {len(valid_points)}")
        self.log.emit(f"   Не будет скачано: {len(invalid_points)}")

        if valid_points:
            self.log.emit(f"\n📥 Скачивание {len(valid_points)} пунктов (Асинхронно)...")
            self.info.emit(f"Начинаю скачивание файлов для {len(valid_points)} пунктов...", "wait")

        completed_dl = 0
        suffix = CONFIG.get("INVALID_SUFFIX", " (не корректный)")

        async def download_point(point):
            nonlocal completed_dl
            if self._is_cancelled:
                return None
            p_num = point["number"]

            # Неполная папка могла остаться с прошлого раза под другим именем
            # (в том числе «не корректный») — докачиваем в неё же
            local_folder = os.path.join(self.local_root, f"п.{p_num}")
            existing = find_local_point(self.local_root, p_num)
            if existing:
                local_folder, err = normalize_point_folder(self.local_root, existing, p_num)
                if err:
                    self.log.emit(f"   ⚠ п.{p_num}: {err}")
            os.makedirs(local_folder, exist_ok=True)

            self.log.emit(f"── Начато скачивание п.{p_num} ──")

            # Качаем папки пункта Одновременно!
            dl_tasks = []
            for target, t_path in point["target_paths"].items():
                if t_path:
                    local_target = os.path.join(local_folder, target)
                    if os.path.exists(local_target) and os.listdir(local_target):
                        self.log.emit(f"   п.{p_num} 📂 {target}: Уже скачано, пропускаем")
                    else:
                        self.log.emit(f"   п.{p_num} 📂 {target}:")
                        dl_tasks.append(
                            api_items.download_folder_recursive(session, t_path, local_target, log_callback=self.log.emit)
                        )

            # Скачиваем необязательные папки (если найдены)
            for opt_name, opt_path in point.get("optional_paths", {}).items():
                if opt_path:
                    local_opt = os.path.join(local_folder, opt_name)
                    if os.path.exists(local_opt) and os.listdir(local_opt):
                        self.log.emit(f"   п.{p_num} 📂 {opt_name}: Уже скачано, пропускаем")
                    else:
                        self.log.emit(f"   п.{p_num} 📂 {opt_name} (необяз.):")
                        dl_tasks.append(
                            api_items.download_folder_recursive(session, opt_path, local_opt, log_callback=self.log.emit)
                        )

            if dl_tasks:
                await asyncio.gather(*dl_tasks)

            self._copy_templates(local_folder, p_num, templates)
            set_stage_status(local_folder, "stage1", True)

            # ── Проверка корректности: без обязательных папок пункт
            #    дальше не обрабатывается (переименовываем в «не корректный»)
            missing_req = missing_required(local_folder)
            if missing_req:
                set_stage_status(local_folder, "invalid", True)
                set_stage_status(local_folder, "missing_folders", missing_req)
                invalid_folder = os.path.join(self.local_root, f"п.{p_num}{suffix}")
                if os.path.abspath(local_folder) != os.path.abspath(invalid_folder):
                    if os.path.exists(invalid_folder):
                        # Чужую папку не удаляем: пункт и так заблокирован пометкой в статусе
                        self.log.emit(f"   ⚠ «{os.path.basename(invalid_folder)}» уже есть — "
                                      f"п.{p_num} помечен только в статусе")
                    else:
                        try:
                            os.rename(local_folder, invalid_folder)
                            local_folder = invalid_folder
                        except Exception as e:
                            self.log.emit(f"   ⚠ Не удалось переименовать папку п.{p_num}: {e}")
                self.info.emit(
                    f"п.{p_num}: не будет обработан — нет папок: {', '.join(missing_req)}", "warn")
                self.log.emit(
                    f"   🚫 п.{p_num} помечен НЕ КОРРЕКТНЫМ — отсутствуют: {', '.join(missing_req)}")
            elif clear_invalid_mark(local_folder):
                self.log.emit(f"   ✓ п.{p_num}: пометка «не корректный» снята")

            completed_dl += 1
            pct = 45 + int(completed_dl / len(valid_points) * 50)
            self.progress.emit(min(pct, 95))
            return p_num

        # Запускаем скачивание всех пунктов сразу
        download_results = await asyncio.gather(*[download_point(p) for p in valid_points])
        if self._is_cancelled:
            return

        for i, p_num in enumerate(download_results):
            if p_num:
                point_data = valid_points[i]
                report["created"] += 1
                folders_status = {tgt: bool(pth) for tgt, pth in point_data["target_paths"].items()}
                optional_status = {opt: bool(pth) for opt, pth in point_data.get("optional_paths", {}).items()}
                has_missing = bool(point_data["missing"])

                report["details"].append({
                    "point": p_num,
                    "status": "⚠ Частично" if has_missing else "✓ Успешно",
                    "folders": folders_status,
                    "optional": optional_status
                })

        # Добавляем невалидные в отчёт
        for p in invalid_points:
            report["not_created"] += 1
            reason = "Не найден на диске"
            if p["remote_path"]:
                missing = [tgt for tgt, pth in p["target_paths"].items() if not pth]
                reason = f"Нет папок: {', '.join(missing)}"

            report["details"].append({
                "point": p["number"],
                "status": f"Не создан — {reason}",
            })

    # ─── Основной сценарий ──────────────────────────────────────────

    async def async_run(self):
        report = {
            "excel_total": 0,
            "excel_batch": 0,
            "found_on_disk": 0,
            "valid_points": 0,
            "created": 0,
            "not_created": 0,
            "details": [],  # список словарей с деталями по каждому пункту
        }

        try:
            # ── 1. Создание корневой папки ──────────────────────────────
            self.log.emit("═" * 40)
            self.log.emit("ЭТАП 1: Создание структуры (🔥 ТУРБО-Асинхронный режим 🔥)")
            self.log.emit("═" * 40)

            os.makedirs(self.local_root, exist_ok=True)
            self.info.emit(f"Корневая папка готова: {os.path.basename(self.local_root)}", "done")
            self.log.emit(f"📁 Папка: {self.local_root}")

            # ── 2. Определение пунктов (Excel или Ручной ввод) ───────────────────────────
            if self.manual_points:
                self.info.emit(f"Режим ручного ввода. Пунктов: {len(self.manual_points)}", "info")
                self.log.emit("\n📝 Режим: Ручной ввод пунктов")
                points = self.manual_points
                total_count = len(points)
                report["excel_total"] = total_count
                report["excel_batch"] = total_count
                self.log.emit(f"   Взято пунктов из ручного ввода: {total_count}")
            else:
                self.log.emit(f"\n📊 Чтение Excel: {self.excel_path}")
                self.info.emit("Читаю Excel файл с реестром пунктов...", "wait")
                total_count, points = analyze_excel(self.excel_path)
                report["excel_total"] = total_count
                report["excel_batch"] = len(points)
                self.info.emit(f"Найдено подходящих пунктов: {len(points)}", "done")

                self.log.emit(f"   Всего подходящих строк: {total_count}")
                self.log.emit(f"   Взято для пакета: {len(points)}")

            if not points:
                self.log.emit("⚠ Нет подходящих пунктов для поиска.")
                report["details"].append({
                    "point": "—",
                    "status": "Нет данных (Excel пуст или ручной ввод пуст)",
                })
                self.report_ready.emit(report)
                self.finished_ok.emit(False)
                return

            self.progress.emit(5)

            # ── 3. Свои комплекты в ПНОС — без Яндекс Диска ────────────
            # Комплект, собранный вручную (Паспорт и Первичка на месте), берётся
            # как есть: с диска многое перестало скачиваться, проще положить самому
            own, remote = [], []
            for p_num in points:
                folder = find_local_point(self.local_root, p_num)
                if folder and not missing_required(folder):
                    own.append((p_num, folder))
                else:
                    remote.append(p_num)
            self.log.emit(f"\n📂 Уже лежат в ПНОС (Паспорт и Первичка на месте): {len(own)}")
            self.log.emit(f"   Искать на Яндекс Диске: {len(remote)}")
            if own:
                self.info.emit(f"Свои комплекты в папке ПНОС: {len(own)}, "
                               f"с Яндекс.Диска: {len(remote)}", "info")

            need_templates = bool(remote) or any(
                not os.path.exists(os.path.join(f, n))
                for _, f in own for n in CONFIG["TEMPLATE_FILES"])

            api_items = YandexDiskAPI(CONFIG["TOKEN_ITEMS"])
            api_base = YandexDiskAPI(CONFIG["TOKEN_BASE"])
            api_items.set_log_callback(self.log.emit)
            api_base.set_log_callback(self.log.emit)

            async with aiohttp.ClientSession() as session:
                # ── 4. Связь — только если без диска не обойтись ────────
                items_problem = base_problem = None
                if remote or need_templates:
                    items_problem, base_problem = await asyncio.gather(
                        api_items.check_connection(session),
                        api_base.check_connection(session),
                    )
                    if not items_problem and not base_problem:
                        self.log.emit("   ✓ Связь с Яндекс.Диском есть")

                if remote and items_problem:
                    # Без связи пункт выглядел «не найденным» — называем причину
                    self.info.emit(f"Нет связи с Яндекс.Диском: {items_problem}", "error")
                    self.log.emit(f"\n❌ Нет связи с Яндекс.Диском (диск «материалы пунктов»): {items_problem}")
                    self.log.emit(f"   Не скачаны: {', '.join(remote)}")
                    for p_num in remote:
                        report["details"].append({
                            "point": p_num,
                            "status": "Не скачан — нет связи с Яндекс.Диском",
                        })
                    report["not_created"] += len(remote)
                    remote = []
                if need_templates and base_problem:
                    if base_problem != items_problem:
                        self.info.emit(f"Диск шаблонов недоступен: {base_problem}", "warn")
                    self.log.emit(f"   ⚠ Диск шаблонов недоступен — беру сохранённую копию: {base_problem}")

                templates = {}
                if need_templates:
                    templates = await self._prepare_templates(session, api_base, online=not base_problem)
                self.progress.emit(20)

                # ── 5. Свои комплекты: имя, статус, шаблоны ─────────────
                for p_num, folder in own:
                    if self._is_cancelled:
                        break
                    report["details"].append(self._accept_own_point(p_num, folder, templates))
                    report["created"] += 1
                self.progress.emit(25)

                # ── 6. Остальные — с Яндекс Диска ───────────────────────
                if remote and not self._is_cancelled:
                    await self._fetch_remote(session, api_items, remote, templates, report)

            if self._is_cancelled:
                self.log.emit("❌ Отменено пользователем.")
                self.finished_ok.emit(False)
                return

            ok = report["created"] > 0
            self.progress.emit(100)
            if ok:
                self.info.emit("Этап 1 успешно завершен!", "done")
            else:
                self.info.emit("Этап 1: ни один пункт не подготовлен", "error")
            self.log.emit(f"\n{'═' * 40}")
            self.log.emit(f"{'✅' if ok else '❌'} Этап 1 завершён! Создано: {report['created']}, "
                          f"Не создано: {report['not_created']}")

            self.report_ready.emit(report)
            self.finished_ok.emit(ok)

        except Exception as e:
            self.info.emit(f"Ошибка: {str(e)[:50]}...", "error")
            self.log.emit(f"\n❌ Критическая ошибка: {e}")
            import traceback
            self.log.emit(traceback.format_exc())
            self.report_ready.emit(report)
            self.finished_ok.emit(False)
