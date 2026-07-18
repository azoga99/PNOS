# -*- coding: utf-8 -*-
"""
Страница «Анализ пунктов» — статистическая проверка готовности пунктов.

ScheduleLoadWorker — читает график Excel в фоне (файл широкий, колонки до RE,
чтение может занять несколько секунд) и отдаёт строки для фильтров UI.

AnalyticsWorker — асинхронно сканирует выбранные пункты на Яндекс Диске
БЕЗ скачивания: ищет папку п.NNNN, проверяет наличие и наполненность
подпапок (Паспорт, Первичка, Стар. ЭПБ, Исполнительная, Акт отбраковки).

Режим автономный: не связан с конвейером и .pnos_status.json.
"""

import asyncio
import datetime

import aiohttp

from PySide6.QtCore import QThread, Signal

from config import CONFIG
from yandex_api import YandexDiskAPI
from excel_service import read_analytics_rows

# Вердикты по пункту: код -> подпись
VERDICT_LABELS = {
    "ready": "Готов к обработке",
    "missing_critical": "Не хватает критических папок",
    "empty": "Совсем пустой",
    "not_found": "Не найден на диске",
    "cancelled": "Проверка прервана",
}


def _plural(n: int, one: str, few: str, many: str) -> str:
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


def fmt_folder_cell(folder: dict | None) -> str:
    """Текст ячейки таблицы результатов для одной подпапки."""
    if folder is None:
        return "—"
    if not folder["exists"]:
        return "нет"
    files = folder["files"]
    dirs = folder["items"] - files
    if folder["items"] == 0:
        return "пуста"
    parts = []
    if files:
        parts.append(f"{files} {_plural(files, 'файл', 'файла', 'файлов')}")
    if dirs:
        parts.append(f"{dirs} {_plural(dirs, 'папка', 'папки', 'папок')}")
    return ", ".join(parts)


class ScheduleLoadWorker(QThread):
    """Фоновое чтение графика Excel для страницы «Анализ пунктов»."""
    log = Signal(str)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    rows_ready = Signal(list)   # список словарей из read_analytics_rows

    def __init__(self, excel_path: str, parent=None):
        super().__init__(parent)
        self.excel_path = excel_path

    def run(self):
        try:
            self.info.emit("Читаю график Excel…", "wait")
            rows = read_analytics_rows(self.excel_path)
            self.log.emit(f"[Анализ пунктов] Прочитано строк графика: {len(rows)}")
            self.rows_ready.emit(rows)
            self.info.emit(f"График прочитан: {len(rows)} пунктов", "done")
            self.finished_ok.emit(True)
        except Exception as e:
            self.log.emit(f"[Анализ пунктов] Ошибка чтения графика: {e}")
            self.info.emit(f"Ошибка чтения графика: {e}", "error")
            self.rows_ready.emit([])
            self.finished_ok.emit(False)


class AnalyticsWorker(QThread):
    """Асинхронный скан выбранных пунктов на Я.Диске (только чтение)."""
    log = Signal(str)
    info = Signal(str, str)
    progress = Signal(int)
    finished_ok = Signal(bool)
    result_ready = Signal(list)  # список результатов по пунктам (в порядке выбора)

    def __init__(self, points: list[str], parent=None):
        super().__init__(parent)
        self.points = points
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        asyncio.run(self.async_run())

    async def async_run(self):
        results = []
        try:
            api = YandexDiskAPI(CONFIG["TOKEN_ITEMS"])
            api.set_log_callback(self.log.emit)
            search_paths = CONFIG["DISK_PATHS_ITEMS"]
            folder_names = CONFIG["TARGET_FOLDERS"] + CONFIG["OPTIONAL_FOLDERS"]
            critical = CONFIG["REQUIRED_FOLDERS"]
            total = len(self.points)

            self.info.emit(f"Проверяю {total} пунктов на Яндекс Диске…", "wait")
            self.log.emit(f"[Анализ пунктов] Пункты: {', '.join(self.points)}")

            async with aiohttp.ClientSession() as session:
                extended_paths = await api.preload_search_paths(session, search_paths)
                if not extended_paths:
                    extended_paths = search_paths
                self.progress.emit(10)

                completed = 0

                async def scan_point(p_num: str) -> dict:
                    nonlocal completed
                    res = {
                        "point": p_num,
                        "found": False,
                        "path": None,
                        "folders": {},
                        "verdict": "not_found",
                        "verdict_text": VERDICT_LABELS["not_found"],
                    }
                    if self._is_cancelled:
                        res["verdict"] = "cancelled"
                        res["verdict_text"] = VERDICT_LABELS["cancelled"]
                        return res

                    folder_path, _ = await api.find_point_folder(
                        session, extended_paths, p_num)
                    if folder_path:
                        res["found"] = True
                        res["path"] = folder_path

                        async def check_sub(name):
                            sub = await api.find_folder_by_name(session, folder_path, name)
                            if not sub:
                                return name, {"exists": False, "files": 0, "items": 0}
                            items = await api.get_folder_contents(session, sub) or []
                            files = sum(1 for i in items if i["type"] == "file")
                            return name, {"exists": True, "files": files,
                                          "items": len(items)}

                        pairs = await asyncio.gather(
                            *[check_sub(n) for n in folder_names])
                        res["folders"] = dict(pairs)

                        any_content = any(v["items"] > 0
                                          for v in res["folders"].values())
                        bad_critical = [
                            (n, "нет" if not res["folders"][n]["exists"] else "пуста")
                            for n in critical if res["folders"][n]["items"] == 0
                        ]
                        if not any_content:
                            res["verdict"] = "empty"
                            res["verdict_text"] = VERDICT_LABELS["empty"]
                        elif bad_critical:
                            res["verdict"] = "missing_critical"
                            res["verdict_text"] = ("Не хватает: " + ", ".join(
                                f"{n} — {why}" for n, why in bad_critical))
                        else:
                            res["verdict"] = "ready"
                            res["verdict_text"] = VERDICT_LABELS["ready"]

                    completed += 1
                    self.progress.emit(10 + int(88 * completed / total))
                    self.log.emit(f"[Анализ пунктов] п.{p_num}: {res['verdict_text']}")
                    return res

                results = list(await asyncio.gather(
                    *(scan_point(p) for p in self.points)))

            n_ready = sum(1 for r in results if r["verdict"] == "ready")
            n_missing = sum(1 for r in results if r["verdict"] == "missing_critical")
            n_empty = sum(1 for r in results if r["verdict"] == "empty")
            n_lost = sum(1 for r in results if r["verdict"] == "not_found")
            self.result_ready.emit(results)
            self.progress.emit(100)
            if self._is_cancelled:
                self.info.emit("Анализ прерван пользователем", "warn")
                self.finished_ok.emit(False)
            else:
                self.info.emit(
                    f"Анализ завершён: готовы {n_ready}, не хватает папок "
                    f"{n_missing}, пустых {n_empty}, не найдено {n_lost}", "done")
                self.finished_ok.emit(True)

        except Exception as e:
            self.log.emit(f"[Анализ пунктов] Ошибка: {e}")
            self.info.emit(f"Ошибка анализа: {e}", "error")
            self.result_ready.emit(results)
            self.finished_ok.emit(False)


# ═══════════════════════════════════════════════════════════════════════
# Выгрузка результатов в Word (python-docx, Word не требуется)
# ═══════════════════════════════════════════════════════════════════════

_VERDICT_COLORS = {          # цвет текста вердикта в Word (RGB)
    "ready": (0x10, 0x7A, 0x3D),
    "missing_critical": (0xB4, 0x5F, 0x06),
    "empty": (0xB9, 0x1C, 0x1C),
    "not_found": (0xB9, 0x1C, 0x1C),
    "cancelled": (0x64, 0x74, 0x8B),
}


def export_analysis_word(save_path: str, results: list[dict],
                         specialist: str, equipment: str):
    """Формирует итоговый Word-документ анализа пунктов."""
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT

    folder_names = CONFIG["TARGET_FOLDERS"] + CONFIG["OPTIONAL_FOLDERS"]

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Анализ готовности пунктов")
    run.bold = True
    run.font.size = Pt(16)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    stamp = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    meta_run = meta.add_run(
        f"Специалист: {specialist}   •   Оборудование: {equipment or 'все виды'}"
        f"   •   {stamp}")
    meta_run.font.size = Pt(10)
    meta_run.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    n_ready = sum(1 for r in results if r["verdict"] == "ready")
    n_missing = sum(1 for r in results if r["verdict"] == "missing_critical")
    n_empty = sum(1 for r in results if r["verdict"] == "empty")
    n_lost = sum(1 for r in results if r["verdict"] == "not_found")
    summary = doc.add_paragraph()
    summary.add_run(
        f"Проверено пунктов: {len(results)}. Готовы к обработке: {n_ready}. "
        f"Не хватает критических папок (Паспорт/Первичка): {n_missing}. "
        f"Совсем пустых: {n_empty}. Не найдено на диске: {n_lost}.")

    # ── Сводная таблица ──
    table = doc.add_table(rows=1, cols=2 + len(folder_names))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    head = ["Пункт"] + folder_names + ["Итог"]
    for i, caption in enumerate(head):
        cell = table.rows[0].cells[i]
        cell.text = ""
        r = cell.paragraphs[0].add_run(caption)
        r.bold = True
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    for res in results:
        row = table.add_row().cells
        row[0].text = f"п.{res['point']}"
        for i, name in enumerate(folder_names, start=1):
            row[i].text = (fmt_folder_cell(res["folders"].get(name))
                           if res["found"] else "—")
            row[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        vcell = row[len(folder_names) + 1]
        vcell.text = ""
        vrun = vcell.paragraphs[0].add_run(res["verdict_text"])
        color = _VERDICT_COLORS.get(res["verdict"])
        if color:
            vrun.font.color.rgb = RGBColor(*color)

    # ── Группы пунктов по вердикту ──
    groups = [
        ("ready", "Готовы к обработке"),
        ("missing_critical", "Не хватает критических папок (Паспорт/Первичка)"),
        ("empty", "Совсем пустые"),
        ("not_found", "Не найдены на Яндекс Диске"),
    ]
    for code, caption in groups:
        pts = [r for r in results if r["verdict"] == code]
        if not pts:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        hr = p.add_run(f"{caption} — {len(pts)}")
        hr.bold = True
        color = _VERDICT_COLORS.get(code)
        if color:
            hr.font.color.rgb = RGBColor(*color)
        if code == "missing_critical":
            for r in pts:
                doc.add_paragraph(f"п.{r['point']}: {r['verdict_text']}",
                                  style="List Bullet")
        else:
            doc.add_paragraph(", ".join(f"п.{r['point']}" for r in pts))

    doc.save(save_path)
