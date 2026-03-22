# -*- coding: utf-8 -*-
"""
Этап 1: Создание структуры на локальном ПК.
- Создание папки ПНОС на рабочем столе
- Чтение Excel, фильтрация, первые 20 пунктов
- Поиск папок пунктов на Яндекс Диске (Асинхронно, с aiohttp)
- Скачивание подпапок Паспорт и Первичка (Асинхронно)
- Копирование шаблонов ПНОС_ТТП.docx и ПНОС_ТТП.xlsm
"""

import os
import shutil
import asyncio
import aiohttp

from PySide6.QtCore import QThread, Signal

from config import CONFIG
from yandex_api import YandexDiskAPI
from excel_service import analyze_excel


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

            # ── 3. Поиск папок на Яндекс Диске ─────────────────────────
            api_items = YandexDiskAPI(CONFIG["TOKEN_ITEMS"])
            api_base = YandexDiskAPI(CONFIG["TOKEN_BASE"])
            search_paths = CONFIG["DISK_PATHS_ITEMS"]

            # Подключаем логирование к API для отладки
            api_items.set_log_callback(self.log.emit)
            api_base.set_log_callback(self.log.emit)

            self.log.emit(f"\n🔍 Поиск в {len(search_paths)} папках на Яндекс Диске...")
            self.info.emit(f"Ищу папки пунктов на Яндекс Диске (в {len(search_paths)} регионах)...", "wait")
            self.log.emit(f"   Искомые пункты: {', '.join(points[:5])}{'...' if len(points) > 5 else ''}")

            # Запускаем единую сессию для всех запросов
            async with aiohttp.ClientSession() as session:
                # Предзагрузка списков всех папок (ускоряет поиск в N раз)
                extended_search_paths = await api_items.preload_search_paths(session, search_paths)
                self.progress.emit(15)

                points_info = []
                completed_checks = 0

                async def check_point(p_num):
                    nonlocal completed_checks
                    info = {
                        "number": p_num,
                        "remote_path": None,
                        "target_paths": {},
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

                        info["is_valid"] = True  # Всегда скачиваем, если папка пункта существует
                        
                    completed_checks += 1
                    pct = 15 + int((completed_checks) / len(points) * 20)
                    self.progress.emit(min(pct, 35))
                    return info

                # Запускаем проверку ВСЕХ пунктов одновременно
                check_tasks = [check_point(num) for num in points]
                points_info = await asyncio.gather(*check_tasks)

                if self._is_cancelled:
                    self.log.emit("❌ Отменено пользователем.")
                    self.finished_ok.emit(False)
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
                self.progress.emit(35)

                # ── 4. Модальное окно (данные для отчёта) ───────────────────
                valid_points = [p for p in points_info if p["is_valid"]]
                invalid_points = [p for p in points_info if not p["is_valid"]]

                self.log.emit(f"\n📋 Итог анализа:")
                self.log.emit(f"   В Excel: {total_count} (взято {len(points)})")
                self.log.emit(f"   Найдено на диске: {report['found_on_disk']}")
                self.log.emit(f"   С обяз. папками: {report['valid_points']}")
                self.log.emit(f"   Будет создано: {len(valid_points)}")
                self.log.emit(f"   Не будет создано: {len(invalid_points)}")

                if not valid_points:
                    self.log.emit("⚠ Нет подходящих пунктов для скачивания.")
                    for p in invalid_points:
                        report["details"].append({
                            "point": p["number"],
                            "status": "Не создан — нет обязательных папок или не найден",
                        })
                    report["not_created"] = len(invalid_points)
                    self.report_ready.emit(report)
                    self.finished_ok.emit(False)
                    return

                # ── 5. Скачивание базовых шаблонов (один раз) ──────────────
                self.log.emit(f"\n📥 Скачивание шаблонов с диска шаблонов...")
                self.info.emit("Скачиваю базовые шаблоны (Excel/Word)...", "wait")
                template_files_local = {}

                base_contents = await api_base.get_folder_contents(session, CONFIG["DISK_PATH_BASE"])
                if base_contents:
                    for item in base_contents:
                        if item["type"] == "file" and item["name"] in CONFIG["TEMPLATE_FILES"]:
                            # Скачиваем во временное место (local_root)
                            temp_path = os.path.join(self.local_root, item["name"])
                            if not os.path.exists(temp_path):
                                if await api_base.download_file(session, item["path"], temp_path):
                                    self.log.emit(f"   ↓ {item['name']}")
                                    template_files_local[item["name"]] = temp_path
                                else:
                                    self.log.emit(f"   ✗ Ошибка: {item['name']}")
                            else:
                                template_files_local[item["name"]] = temp_path
                                self.log.emit(f"   ✓ Уже есть: {item['name']}")

                self.progress.emit(40)

                # ── 6. Скачивание пунктов (Параллельно) ──────────────────────────────────
                self.log.emit(f"\n📥 Скачивание {len(valid_points)} пунктов (Асинхронно)...")
                self.info.emit(f"Начинаю скачивание файлов для {len(valid_points)} пунктов...", "wait")

                completed_dl = 0

                async def download_point(point):
                    nonlocal completed_dl
                    if self._is_cancelled:
                        return None
                    p_num = point["number"]
                    local_folder = os.path.join(self.local_root, f"п.{p_num}")
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
                    
                    if dl_tasks:
                        await asyncio.gather(*dl_tasks)

                    # Копируем шаблоны в папку пункта
                    for tpl_name, tpl_path in template_files_local.items():
                        dest = os.path.join(local_folder, tpl_name)
                        if not os.path.exists(dest):
                            shutil.copy2(tpl_path, dest)
                            self.log.emit(f"   п.{p_num} 📄 {tpl_name} скопирован")

                    completed_dl += 1
                    pct = 40 + int(completed_dl / len(valid_points) * 55)
                    self.progress.emit(min(pct, 95))
                    return p_num

                # Запускаем скачивание всех пунктов сразу
                dl_tasks = [download_point(p) for p in valid_points]
                download_results = await asyncio.gather(*dl_tasks)

                if self._is_cancelled:
                    self.log.emit("❌ Отменено пользователем.")
                    self.finished_ok.emit(False)
                    return

                for i, p_num in enumerate(download_results):
                    if p_num:
                        point_data = valid_points[i]
                        report["created"] += 1
                        folders_status = {tgt: bool(pth) for tgt, pth in point_data["target_paths"].items()}
                        has_missing = bool(point_data["missing"])
                        
                        report["details"].append({
                            "point": p_num,
                            "status": "⚠ Частично" if has_missing else "✓ Успешно",
                            "folders": folders_status
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

                self.progress.emit(100)
                self.info.emit("Этап 1 успешно завершен!", "done")
                self.log.emit(f"\n{'═' * 40}")
                self.log.emit(f"✅ Этап 1 завершён! Создано: {report['created']}, "
                             f"Не создано: {report['not_created']}")

                # Очистка корневой папки от скачанных шаблонов
                for tpl_name, tpl_path in template_files_local.items():
                    if os.path.exists(tpl_path):
                        try:
                            os.remove(tpl_path)
                            self.log.emit(f"   🗑 Удалён временный шаблон из корня: {tpl_name}")
                        except Exception as e:
                            pass

                self.report_ready.emit(report)
                self.finished_ok.emit(True)

        except Exception as e:
            self.info.emit(f"Ошибка: {str(e)[:50]}...", "error")
            self.log.emit(f"\n❌ Критическая ошибка: {e}")
            import traceback
            self.log.emit(traceback.format_exc())
            self.report_ready.emit(report)
            self.finished_ok.emit(False)
