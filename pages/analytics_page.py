# -*- coding: utf-8 -*-
"""
Страница «Анализ пунктов» — статистика готовности (вне конвейера).

Примесь к MainWindow: методы работают с тем же self, что и окно. Общее
для всех страниц (_get_local_root, _add_activity, _switch_page, _workers)
живёт в main_window.py.
"""

import os
import time

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QProgressBar,
    QFrame, QLineEdit, QFileDialog, QMessageBox, QStackedWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QListWidget,
    QListWidgetItem,
)

from config import CONFIG
from widgets import dot_icon, LOG_COLORS, EmptyState


# Статусы согласования на странице «Анализ пунктов».
# Коды и подписи дублируют excel_service.ANALYTICS_STATUS_LABELS — сам модуль
# здесь не импортируется, чтобы не тянуть pandas при старте приложения.
AN_STATUS_LABELS = {
    "in_work": "Не оформлен",
    "done": "Оформлен, не отправлен",
    "sent_approval": "Отправлен Шарову",
    "approved": "Согласован Шаровым",
    "sent_final": "Отправлен Жиганьшину",
}
AN_STATUS_COLORS = {
    "in_work": "#94a3b8",
    "done": "#6366f1",
    "sent_approval": "#f59e0b",
    "approved": "#10b981",
    "sent_final": "#0d9488",
}
AN_VERDICT_COLORS = {
    "ready": LOG_COLORS["done"],
    "missing_critical": LOG_COLORS["warn"],
    "empty": LOG_COLORS["error"],
    "not_found": LOG_COLORS["error"],
    "cancelled": LOG_COLORS["info"],
}


class AnalyticsPageMixin:
    """Страница «Анализ пунктов» — статистика готовности (вне конвейера)."""

    # ═══════════════════════════════════════════════════════════════
    # Страница «Анализ пунктов» — статистика готовности (вне конвейера)
    # ═══════════════════════════════════════════════════════════════

    def _build_chip(self, caption: str, value_obj: str = "chip_value"):
        """Компактный чип-сводка. Возвращает (frame, value_label)."""
        box = QFrame()
        box.setObjectName("stat_chip")
        v = QVBoxLayout(box)
        v.setContentsMargins(16, 7, 16, 7)
        v.setSpacing(0)
        val = QLabel("0")
        val.setObjectName(value_obj)
        val.setAlignment(Qt.AlignCenter)
        cap = QLabel(caption)
        cap.setObjectName("chip_caption")
        cap.setAlignment(Qt.AlignCenter)
        v.addWidget(val)
        v.addWidget(cap)
        return box, val

    def _setup_analytics_page(self):
        from widgets import NoWheelComboBox as QComboBox

        self._an_rows = []        # все строки графика (после исключения удалённых)
        self._an_checked = set()  # отмеченные номера пунктов
        self._an_results = []     # результаты последнего анализа
        self._an_updating = False  # защита от itemChanged при перестройке списка
        self._an_loaded_path = None  # путь графика, из которого прочитаны строки

        outer = QVBoxLayout(self.page_analytics)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(10)

        # ── Шапка: заголовок + вкладки (как на «Обзор и отчёт») ──
        head = QHBoxLayout()
        head.addWidget(self._page_title("Анализ пунктов"))
        head.addStretch()
        self.an_btn_tab_select = QPushButton("Выбор пунктов")
        self.an_btn_tab_results = QPushButton("Результаты")
        for b in (self.an_btn_tab_select, self.an_btn_tab_results):
            b.setObjectName("segment")
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
        self.an_btn_tab_select.setChecked(True)
        self.an_btn_tab_select.clicked.connect(lambda: self._an_set_tab(0))
        self.an_btn_tab_results.clicked.connect(lambda: self._an_set_tab(1))
        head.addWidget(self.an_btn_tab_select)
        head.addWidget(self.an_btn_tab_results)
        outer.addLayout(head)

        # ── Строка графика: файл берётся из Настроек, отдельно не выбирается ──
        file_row = QHBoxLayout()
        file_row.setSpacing(10)
        self.an_lbl_file = QLabel("График ещё не прочитан")
        self.an_lbl_file.setObjectName("hint")
        file_row.addWidget(self.an_lbl_file, 1)
        self.an_btn_reload = QPushButton("Перечитать график")
        self.an_btn_reload.setObjectName("ghost_button")
        self.an_btn_reload.setCursor(Qt.PointingHandCursor)
        self.an_btn_reload.clicked.connect(self._an_load_schedule)
        file_row.addWidget(self.an_btn_reload)
        outer.addLayout(file_row)

        self.an_tabs = QStackedWidget()
        outer.addWidget(self.an_tabs, 1)

        # ═══ Вкладка 0: выбор пунктов ═══
        tab_sel = QWidget()
        sv = QVBoxLayout(tab_sel)
        sv.setContentsMargins(0, 0, 0, 0)
        sv.setSpacing(8)

        filters_panel = QFrame()
        filters_panel.setObjectName("settings_panel")
        filters = QHBoxLayout(filters_panel)
        filters.setContentsMargins(14, 10, 14, 10)
        filters.setSpacing(10)
        filters.addWidget(QLabel("Специалист"))
        self.an_cb_person = QComboBox()
        self.an_cb_person.addItems(CONFIG["ANALYTICS"]["SPECIALISTS"])
        self.an_cb_person.currentIndexChanged.connect(self._an_person_changed)
        filters.addWidget(self.an_cb_person)
        filters.addWidget(QLabel("Оборудование"))
        self.an_cb_equipment = QComboBox()
        self.an_cb_equipment.addItem("Все виды")
        self.an_cb_equipment.setEnabled(False)
        self.an_cb_equipment.currentIndexChanged.connect(self._an_refresh_list)
        filters.addWidget(self.an_cb_equipment, 1)
        filters.addWidget(QLabel("Статус"))
        self.an_cb_status = QComboBox()
        self.an_cb_status.addItem("Все статусы", "all")
        for code, label in AN_STATUS_LABELS.items():
            self.an_cb_status.addItem(label, code)
        self.an_cb_status.currentIndexChanged.connect(self._an_refresh_list)
        filters.addWidget(self.an_cb_status, 1)
        self.an_search = QLineEdit()
        self.an_search.setPlaceholderText("Поиск № пункта")
        self.an_search.setFixedWidth(150)
        self.an_search.textChanged.connect(self._an_refresh_list)
        filters.addWidget(self.an_search)
        sv.addWidget(filters_panel)

        self.an_list = QListWidget()
        self.an_list.setObjectName("activity_feed")
        self.an_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.an_list.itemChanged.connect(self._an_item_changed)
        self.an_list_box = EmptyState(
            self.an_list, "График ещё не прочитан",
            "Программа читает сводный график сама при запуске. Если сервер "
            "недоступен, укажите файл графика в «Настройках» и нажмите "
            "«Перечитать график».")
        sv.addWidget(self.an_list_box, 1)

        sel_row = QHBoxLayout()
        sel_row.setSpacing(6)
        for caption, slot in (("Все", self._an_check_all),
                              ("Отметить выделенные", self._an_check_selected),
                              ("Снять", self._an_clear_checks)):
            b = QPushButton(caption)
            b.setObjectName("ghost_button")
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(slot)
            sel_row.addWidget(b)
        self.an_lbl_counter = QLabel(
            "Выбрано: 0 / %d" % CONFIG["ANALYTICS"]["MAX_POINTS"])
        self.an_lbl_counter.setObjectName("hint")
        sel_row.addWidget(self.an_lbl_counter)
        sel_row.addStretch()
        self.an_btn_start = QPushButton("Проверить на Я.Диске")
        self.an_btn_start.setObjectName("card_button")
        self.an_btn_start.setCursor(Qt.PointingHandCursor)
        self.an_btn_start.setEnabled(False)
        self.an_btn_start.clicked.connect(self._an_start_scan)
        sel_row.addWidget(self.an_btn_start)
        sv.addLayout(sel_row)
        self.an_tabs.addWidget(tab_sel)

        # ═══ Вкладка 1: результаты ═══
        tab_res = QWidget()
        rv = QVBoxLayout(tab_res)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(8)

        chips_row = QHBoxLayout()
        chips_row.setSpacing(8)
        chip_total, self.an_chip_total = self._build_chip("ПРОВЕРЕНО")
        chip_ready, self.an_chip_ready = self._build_chip("ГОТОВЫ", "chip_value_good")
        chip_missing, self.an_chip_missing = self._build_chip("НЕ ХВАТАЕТ", "chip_value_warn")
        chip_empty, self.an_chip_empty = self._build_chip("ПУСТЫЕ", "chip_value_bad")
        chip_lost, self.an_chip_lost = self._build_chip("НЕ НАЙДЕНЫ", "chip_value_bad")
        for c in (chip_total, chip_ready, chip_missing, chip_empty, chip_lost):
            chips_row.addWidget(c)
        chips_row.addStretch()
        self.an_btn_export = QPushButton("Выгрузить в Word")
        self.an_btn_export.setObjectName("ghost_button")
        self.an_btn_export.setCursor(Qt.PointingHandCursor)
        self.an_btn_export.setEnabled(False)
        self.an_btn_export.clicked.connect(self._an_export_word)
        chips_row.addWidget(self.an_btn_export)
        rv.addLayout(chips_row)

        self._an_folder_names = CONFIG["TARGET_FOLDERS"] + CONFIG["OPTIONAL_FOLDERS"]
        self.an_table = QTableWidget()
        self.an_table.setColumnCount(2 + len(self._an_folder_names))
        self.an_table.setHorizontalHeaderLabels(
            ["Пункт"] + self._an_folder_names + ["Итог"])
        self.an_table.verticalHeader().setVisible(False)
        self.an_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.an_table.setSelectionMode(QTableWidget.NoSelection)
        header = self.an_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.an_table_box = EmptyState(
            self.an_table, "Результатов пока нет",
            "Отметьте пункты на вкладке «Выбор пунктов» и нажмите «Проверить "
            "на Я.Диске» — здесь появится, каких папок у пунктов не хватает.")
        rv.addWidget(self.an_table_box, 1)
        self.an_tabs.addWidget(tab_res)

        # ── Низ: статус + остановка + прогресс ──
        bottom = QHBoxLayout()
        bottom.setSpacing(10)
        self.an_lbl_status = QLabel(
            "График прочитается сам при открытии страницы.")
        self.an_lbl_status.setObjectName("hint")
        bottom.addWidget(self.an_lbl_status, 1)
        self.an_btn_stop = QPushButton("Остановить")
        self.an_btn_stop.setObjectName("ghost_button")
        self.an_btn_stop.setCursor(Qt.PointingHandCursor)
        self.an_btn_stop.setEnabled(False)
        self.an_btn_stop.clicked.connect(self._an_stop_scan)
        bottom.addWidget(self.an_btn_stop)
        self.an_progress = QProgressBar()
        self.an_progress.setRange(0, 100)
        self.an_progress.setValue(0)
        self.an_progress.setFixedHeight(10)
        self.an_progress.setFixedWidth(220)
        self.an_progress.setTextVisible(False)
        bottom.addWidget(self.an_progress)
        outer.addLayout(bottom)

    def _an_set_tab(self, index: int):
        self.an_tabs.setCurrentIndex(index)
        self.an_btn_tab_select.setChecked(index == 0)
        self.an_btn_tab_results.setChecked(index == 1)

    # ── Загрузка графика (единый мастер-файл, см. _master_excel_path) ──

    def _an_auto_load(self):
        """Первое открытие страницы — график читается сам, без кнопок."""
        if self._an_rows or "an_load" in self._workers:
            return
        self._an_load_schedule()

    def _an_load_schedule(self):
        if "an_load" in self._workers:
            return
        path = self._master_excel_path()
        if not path or not os.path.isfile(path):
            self.an_lbl_file.setText("График не найден")
            if not self._an_rows:
                self.an_list_box.show_hint(
                    "Сводный график не найден",
                    "Укажите файл графика в «Настройках» (блок «План-график»), "
                    "затем нажмите «Перечитать график».")
            self._an_set_status(
                "Не найден мастер-файл сводного графика — укажите его "
                "на странице «Настройки»", "error")
            return
        self.an_lbl_file.setText(
            f"График: {os.path.basename(path)} (мастер-файл из Настроек)")
        self.an_btn_reload.setEnabled(False)
        self._an_loaded_path = path
        if not self._an_rows:
            self.an_list_box.show_hint(
                "Читаю сводный график…",
                "Обычно это несколько секунд. Список пунктов появится здесь.")

        from workers.analytics_worker import ScheduleLoadWorker
        worker = ScheduleLoadWorker(path, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(self._an_set_status)
        worker.rows_ready.connect(self._an_on_rows)
        worker.finished_ok.connect(self._an_on_load_finished)
        worker.start()
        self._workers["an_load"] = worker

    def _an_on_load_finished(self, success: bool):
        self.an_btn_reload.setEnabled(True)
        self._workers.pop("an_load", None)
        if not success and not self._an_rows:
            self.an_list_box.show_hint(
                "Не удалось прочитать график",
                "Причина — в строке состояния внизу страницы. Проверьте файл "
                "и нажмите «Перечитать график».")

    def _an_on_rows(self, rows: list):
        self._an_rows = rows
        self._an_checked.clear()
        self._an_fill_equipment()
        self._an_refresh_list()
        if not rows:
            self.an_list_box.show_hint(
                "В графике нет пунктов",
                "В прочитанном файле не нашлось ни одной строки с пунктом. "
                "Проверьте, тот ли файл указан в «Настройках».")
        if rows:
            if self._an_loaded_path == CONFIG.get("MACRO_MASTER_WB_PATH"):
                src = "с сервера"
            else:
                try:
                    mtime = time.strftime("%d.%m.%Y", time.localtime(
                        os.path.getmtime(self._an_loaded_path)))
                    src = f"из файла (изменён {mtime})"
                except (OSError, TypeError):
                    src = "из файла"
            self._show_toast(
                f"График успешно подтянут {src} — пунктов: {len(rows)}")

    # ── Фильтры и список ──

    def _an_person_rows(self) -> list:
        """Строки графика выбранного специалиста (в порядке графика)."""
        surname = self.an_cb_person.currentText().strip().lower()
        return [r for r in self._an_rows if surname in r["person"].lower()]

    def _an_fill_equipment(self):
        current = self.an_cb_equipment.currentText()
        kinds = sorted({r["equipment"] for r in self._an_person_rows()
                        if r["equipment"]})
        self.an_cb_equipment.blockSignals(True)
        self.an_cb_equipment.clear()
        self.an_cb_equipment.addItem("Все виды")
        self.an_cb_equipment.addItems(kinds)
        idx = self.an_cb_equipment.findText(current)
        if idx > 0:
            self.an_cb_equipment.setCurrentIndex(idx)
        self.an_cb_equipment.setEnabled(bool(kinds))
        self.an_cb_equipment.blockSignals(False)

    def _an_person_changed(self):
        self._an_fill_equipment()
        self._an_refresh_list()

    def _an_refresh_list(self):
        equipment = self.an_cb_equipment.currentText()
        status = self.an_cb_status.currentData()
        search = self.an_search.text().strip().lower()

        self._an_updating = True
        self.an_list.clear()
        for r in self._an_person_rows():
            if equipment != "Все виды" and r["equipment"] != equipment:
                continue
            if status and status != "all" and r["status"] != status:
                continue
            if search and search not in r["point"].lower():
                continue
            item = QListWidgetItem(f"п.{r['point']} — {r['status_label']}")
            item.setIcon(dot_icon(AN_STATUS_COLORS.get(r["status"], "#94a3b8")))
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if r["point"] in self._an_checked
                               else Qt.Unchecked)
            item.setData(Qt.UserRole, r["point"])
            self.an_list.addItem(item)
        self._an_updating = False
        self._an_update_counter()
        # пока график не прочитан, подсказку ставят загрузчик и _an_on_rows
        if self._an_rows:
            if self.an_list.count():
                self.an_list_box.show_content()
            else:
                self.an_list_box.show_hint(
                    "Нет пунктов по выбранным условиям",
                    "Смените специалиста, вид оборудования или статус, "
                    "либо очистите поиск.")

    def _an_item_changed(self, item):
        if self._an_updating:
            return
        point = item.data(Qt.UserRole)
        if item.checkState() == Qt.Checked:
            self._an_checked.add(point)
        else:
            self._an_checked.discard(point)
        self._an_update_counter()

    def _an_update_counter(self):
        max_points = CONFIG["ANALYTICS"]["MAX_POINTS"]
        n = len(self._an_checked)
        self.an_lbl_counter.setText(f"Выбрано: {n} / {max_points}")
        self.an_lbl_counter.setStyleSheet(
            f"color: {LOG_COLORS['error']};" if n > max_points else "")
        self.an_btn_start.setText(
            f"Проверить на Я.Диске ({n})" if n else "Проверить на Я.Диске")
        if "analytics" not in self._workers:
            self.an_btn_start.setEnabled(0 < n <= max_points)

    def _an_check_all(self):
        max_points = CONFIG["ANALYTICS"]["MAX_POINTS"]
        self._an_updating = True
        checked = len(self._an_checked)
        truncated = False
        for i in range(self.an_list.count()):
            item = self.an_list.item(i)
            if item.checkState() == Qt.Checked:
                continue
            if checked >= max_points:
                truncated = True
                break
            item.setCheckState(Qt.Checked)
            self._an_checked.add(item.data(Qt.UserRole))
            checked += 1
        self._an_updating = False
        self._an_update_counter()
        if truncated:
            self._an_set_status(
                f"Отмечены первые {max_points} пунктов — больше за один анализ нельзя",
                "warn")

    def _an_check_selected(self):
        self._an_updating = True
        for item in self.an_list.selectedItems():
            item.setCheckState(Qt.Checked)
            self._an_checked.add(item.data(Qt.UserRole))
        self._an_updating = False
        self._an_update_counter()

    def _an_clear_checks(self):
        self._an_checked.clear()
        self._an_updating = True
        for i in range(self.an_list.count()):
            self.an_list.item(i).setCheckState(Qt.Unchecked)
        self._an_updating = False
        self._an_update_counter()

    # ── Анализ на Я.Диске ──

    def _an_start_scan(self):
        # порядок пунктов — как в графике у выбранного специалиста
        points = [r["point"] for r in self._an_person_rows()
                  if r["point"] in self._an_checked]
        max_points = CONFIG["ANALYTICS"]["MAX_POINTS"]
        if not points:
            QMessageBox.warning(self, "Ошибка",
                                "Отметьте хотя бы один пункт для анализа.")
            return
        if len(points) > max_points:
            QMessageBox.warning(
                self, "Ошибка",
                f"Отмечено {len(points)} пунктов — за один анализ можно "
                f"проверить не более {max_points}.")
            return

        self.an_btn_start.setEnabled(False)
        self.an_btn_reload.setEnabled(False)
        self.an_btn_stop.setEnabled(True)
        self.an_btn_export.setEnabled(False)
        self.an_progress.setValue(0)
        self.an_table.setRowCount(0)
        self._an_results = []
        self.an_table_box.show_hint(
            "Идёт проверка на Яндекс.Диске…",
            "Результаты появятся здесь, когда проверка закончится.")

        from workers.analytics_worker import AnalyticsWorker
        worker = AnalyticsWorker(points, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(self._an_set_status)
        worker.progress.connect(self.an_progress.setValue)
        worker.result_ready.connect(self._an_on_results)
        worker.finished_ok.connect(self._an_on_scan_finished)
        worker.start()
        self._workers["analytics"] = worker

    def _an_stop_scan(self):
        worker = self._workers.get("analytics")
        if worker and worker.isRunning():
            worker.cancel()
            self.an_btn_stop.setEnabled(False)
            self._an_set_status("Останавливаю анализ…", "warn")

    def _an_on_results(self, results: list):
        from workers.analytics_worker import fmt_folder_cell
        self._an_results = results
        critical = set(CONFIG["REQUIRED_FOLDERS"])

        self.an_table.setRowCount(len(results))
        for row, res in enumerate(results):
            self.an_table.setItem(row, 0, QTableWidgetItem(f"п.{res['point']}"))
            for col, name in enumerate(self._an_folder_names, start=1):
                folder = res["folders"].get(name) if res["found"] else None
                cell = QTableWidgetItem(fmt_folder_cell(folder))
                cell.setTextAlignment(Qt.AlignCenter)
                if folder is None:
                    color = LOG_COLORS["info"]
                elif folder["items"] > 0:
                    color = LOG_COLORS["done"]
                elif folder["exists"]:
                    color = LOG_COLORS["warn"]
                else:
                    color = (LOG_COLORS["error"] if name in critical
                             else LOG_COLORS["info"])
                cell.setForeground(QColor(color))
                self.an_table.setItem(row, col, cell)
            verdict = QTableWidgetItem(res["verdict_text"])
            verdict.setForeground(QColor(
                AN_VERDICT_COLORS.get(res["verdict"], LOG_COLORS["info"])))
            self.an_table.setItem(row, len(self._an_folder_names) + 1, verdict)

        self.an_chip_total.setText(str(len(results)))
        self.an_chip_ready.setText(
            str(sum(1 for r in results if r["verdict"] == "ready")))
        self.an_chip_missing.setText(
            str(sum(1 for r in results if r["verdict"] == "missing_critical")))
        self.an_chip_empty.setText(
            str(sum(1 for r in results if r["verdict"] == "empty")))
        self.an_chip_lost.setText(
            str(sum(1 for r in results if r["verdict"] == "not_found")))
        self.an_btn_export.setEnabled(bool(results))
        if results:
            self.an_table_box.show_content()
        self._an_set_tab(1)

    def _an_on_scan_finished(self, success: bool):
        self.an_btn_reload.setEnabled(True)
        self.an_btn_stop.setEnabled(False)
        self._workers.pop("analytics", None)
        self._an_update_counter()  # вернёт кнопку запуска в актуальное состояние
        if not self._an_results:
            self.an_table_box.show_hint(
                "Проверка не дала результатов",
                "Причина — в строке состояния внизу страницы. Можно запустить "
                "проверку ещё раз.")

    # ── Выгрузка в Word ──

    def _an_export_word(self):
        if not self._an_results:
            return
        person = self.an_cb_person.currentText()
        equipment = self.an_cb_equipment.currentText()
        if equipment == "Все виды":
            equipment = ""
        default_dir = os.path.dirname(self._an_loaded_path or "") or ""
        default_name = os.path.join(
            default_dir,
            f"Анализ пунктов {person} {time.strftime('%d.%m.%Y')}.docx")
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить анализ", default_name, "Word (*.docx)")
        if not path:
            return
        try:
            from workers.analytics_worker import export_analysis_word
            export_analysis_word(os.path.normpath(path), self._an_results,
                                 person, equipment)
        except Exception as e:
            QMessageBox.warning(self, "Ошибка",
                                f"Не удалось сохранить документ:\n{e}")
            return
        self._an_set_status(f"Документ сохранён: {os.path.basename(path)}", "done")
        answer = QMessageBox.question(
            self, "Готово", "Документ сохранён. Открыть его?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if answer == QMessageBox.Yes:
            try:
                os.startfile(path)
            except OSError:
                pass

    def _an_set_status(self, message: str, category: str = "info"):
        self.an_lbl_status.setText(message)
        if category in ("warn", "error"):
            self.an_lbl_status.setStyleSheet(f"color: {LOG_COLORS[category]};")
        else:
            self.an_lbl_status.setStyleSheet("")
        if category == "error":
            self._log_error(f"[Анализ пунктов] {message}")
