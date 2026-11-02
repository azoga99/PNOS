# -*- coding: utf-8 -*-
"""
Страница «Правка отчётов» — автономная пакетная правка готовых отчётов.

Примесь к MainWindow: методы работают с тем же self, что и окно. Общее
для всех страниц (_get_local_root, _add_activity, _switch_page, _workers)
живёт в main_window.py.
"""

import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton,
    QProgressBar, QFrame, QLineEdit, QFileDialog, QMessageBox, QListWidget,
    QListWidgetItem,
)

from utils import title_date_text, find_order_scan
from widgets import dot_icon, LOG_COLORS


class FixReportsPageMixin:
    """Страница «Правка отчётов» — автономная пакетная правка готовых отчётов."""

    # ═══════════════════════════════════════════════════════════════
    # Страница «Правка отчётов» — автономная пакетная правка готовых
    # Word-отчётов (вне конвейера, со своим логом)
    # ═══════════════════════════════════════════════════════════════

    def _setup_fixreports_page(self):
        outer = QVBoxLayout(self.page_fixreports)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(10)

        outer.addWidget(self._page_title("Правка отчётов"))
        sub = QLabel(
            "Пакетное внесение правок в уже созданные отчёты. Обрабатываются "
            "Word-файлы из выбранной папки и из корня папок пунктов п.NNNN "
            "(включая копии «ПЕЧАТЬ»), в подпапки программа не заходит.")
        sub.setObjectName("hint")
        sub.setWordWrap(True)
        outer.addWidget(sub)

        # ── Параметры ──
        params = QFrame()
        params.setObjectName("settings_panel")
        g = QGridLayout(params)
        g.setContentsMargins(14, 12, 14, 12)
        g.setHorizontalSpacing(10)
        g.setVerticalSpacing(8)

        g.addWidget(QLabel("Папка с отчётами"), 0, 0)
        self.fix_entry_folder = QLineEdit()
        self.fix_entry_folder.setPlaceholderText(
            "Папка с Word-файлами отчётов и/или папками пунктов п.NNNN")
        g.addWidget(self.fix_entry_folder, 0, 1)
        btn_browse = QPushButton("📂")
        btn_browse.setObjectName("browse_btn")
        btn_browse.setFixedWidth(40)
        btn_browse.setCursor(Qt.PointingHandCursor)
        btn_browse.clicked.connect(self._fix_browse_folder)
        g.addWidget(btn_browse, 0, 2)

        g.addWidget(QLabel("Дата на титуле"), 1, 0)
        date_row = QHBoxLayout()
        date_row.setSpacing(10)
        self.fix_date_edit = self._make_date_edit()
        self.fix_date_edit.setFixedWidth(120)
        self._set_date_edit(self.fix_date_edit, "06.07.2026")
        self.fix_date_edit.textChanged.connect(self._fix_update_date_preview)
        date_row.addWidget(self.fix_date_edit)
        self.fix_lbl_date_preview = QLabel("")
        self.fix_lbl_date_preview.setObjectName("hint")
        date_row.addWidget(self.fix_lbl_date_preview)
        date_row.addStretch()
        g.addLayout(date_row, 1, 1, 1, 2)

        hint_scan = QLabel(
            "Скан нового приказа (JPG или PNG, одна страница) положите в эту же "
            "папку с отчётами — программа найдёт его сама. Если картинок несколько, "
            "будет взята самая свежая.")
        hint_scan.setObjectName("hint")
        hint_scan.setWordWrap(True)
        g.addWidget(hint_scan, 2, 0, 1, 3)
        outer.addWidget(params)

        # ── Что будет изменено ──
        desc = QFrame()
        desc.setObjectName("settings_panel")
        dv = QVBoxLayout(desc)
        dv.setContentsMargins(14, 12, 14, 12)
        dv.setSpacing(4)
        lbl_desc_title = QLabel("В каждом отчёте будет изменено:")
        lbl_desc_title.setObjectName("form_field_label")
        dv.addWidget(lbl_desc_title)
        lbl_desc = QLabel(
            "1. Дата на титульном листе (снизу справа) — на выбранную, словами.\n"
            "2. Квалификационные удостоверения Жиганьшина — новые номера "
            "Т011-00109-77/04922210 (Э7 ЗС) и Т011-00109-77/05137434 (Э7 ТУ).\n"
            "3. Таблица 1 «Перечень документов» — П 255-002-2026 от 02.07.2026 г.\n"
            "4. Скан приказа о назначении группы эксперта (приложение Б/В) — "
            "заменяется новым с сохранением размеров.")
        lbl_desc.setObjectName("hint")
        lbl_desc.setWordWrap(True)
        dv.addWidget(lbl_desc)
        outer.addWidget(desc)

        # ── Управление ──
        ctrl = QHBoxLayout()
        ctrl.setSpacing(10)
        self.fix_btn_start = QPushButton("Начать правку")
        self.fix_btn_start.setObjectName("green_button")
        self.fix_btn_start.setCursor(Qt.PointingHandCursor)
        self.fix_btn_start.clicked.connect(self._start_fix_reports)
        ctrl.addWidget(self.fix_btn_start)
        self.fix_btn_stop = QPushButton("Остановить")
        self.fix_btn_stop.setObjectName("segment")
        self.fix_btn_stop.setCursor(Qt.PointingHandCursor)
        self.fix_btn_stop.setEnabled(False)
        self.fix_btn_stop.clicked.connect(self._stop_fix_reports)
        ctrl.addWidget(self.fix_btn_stop)
        self.fix_progress = QProgressBar()
        self.fix_progress.setRange(0, 100)
        self.fix_progress.setValue(0)
        self.fix_progress.setFixedHeight(10)
        self.fix_progress.setTextVisible(False)
        ctrl.addWidget(self.fix_progress, 1)
        outer.addLayout(ctrl)

        # ── Локальный журнал (не связан с общей лентой) ──
        lbl_log = QLabel("Журнал правки")
        lbl_log.setObjectName("section_label")
        outer.addWidget(lbl_log)
        self.fix_log_list = QListWidget()
        self.fix_log_list.setObjectName("activity_feed")
        self.fix_log_list.setMinimumHeight(140)
        outer.addWidget(self.fix_log_list, 1)

        self._fix_update_date_preview()

    def _fix_browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Папка с отчётами")
        if folder:
            # QFileDialog отдаёт прямые слэши — Word COM такие пути не открывает
            self.fix_entry_folder.setText(os.path.normpath(folder))

    def _fix_update_date_preview(self):
        s = self._date_edit_value(self.fix_date_edit)
        if s and self._date_ok_or_empty(s) and len(s) == 10:
            try:
                self.fix_lbl_date_preview.setText(f"На титуле будет: «{title_date_text(s)}»")
                return
            except Exception:
                pass
        self.fix_lbl_date_preview.setText("Укажите дату в формате ДД.ММ.ГГГГ")

    def _fix_add_log(self, message: str, category: str = "info"):
        item = QListWidgetItem(message)
        item.setIcon(dot_icon(LOG_COLORS.get(category, LOG_COLORS["info"])))
        self.fix_log_list.addItem(item)
        self.fix_log_list.scrollToBottom()
        if category == "error":
            self._log_error(f"[Правка отчётов] {message}")

    def _start_fix_reports(self):
        folder = self.fix_entry_folder.text().strip()
        if not folder or not os.path.isdir(folder):
            QMessageBox.warning(self, "Ошибка", "Укажите существующую папку с отчётами.")
            return
        has_points = any(d.startswith("п.") and os.path.isdir(os.path.join(folder, d))
                         for d in os.listdir(folder))
        has_docx = any(f.lower().endswith(".docx") and not f.startswith("~$")
                       for f in os.listdir(folder))
        if not has_points and not has_docx:
            QMessageBox.warning(
                self, "Ошибка",
                "В выбранной папке нет ни Word-файлов отчётов, "
                "ни папок пунктов (п.NNNN).")
            return

        date_str = self._date_edit_value(self.fix_date_edit)
        if not date_str or len(date_str) != 10 or not self._date_ok_or_empty(date_str):
            QMessageBox.warning(self, "Ошибка", "Укажите дату титула в формате ДД.ММ.ГГГГ.")
            return

        scan_path = find_order_scan(folder)
        if not scan_path:
            QMessageBox.warning(
                self, "Нет скана приказа",
                "В выбранной папке не найден скан приказа (JPG или PNG).\n"
                "Положите файл в корень папки с отчётами и повторите запуск.")
            return

        answer = QMessageBox.question(
            self, "Подтверждение",
            f"Файлы отчётов будут изменены без резервных копий.\n\n"
            f"Дата на титуле: «{title_date_text(date_str)}»\n"
            f"Скан приказа: {os.path.basename(scan_path)}\n\nПродолжить?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            return

        self.fix_btn_start.setEnabled(False)
        self.fix_btn_stop.setEnabled(True)
        self.fix_progress.setValue(0)
        self.fix_log_list.clear()

        from workers.fix_reports_worker import FixReportsWorker
        worker = FixReportsWorker(folder, date_str, scan_path, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(self._fix_add_log)
        worker.progress.connect(self.fix_progress.setValue)
        worker.finished_ok.connect(self._on_fix_reports_finished)
        worker.start()
        self._workers["fixreports"] = worker

    def _stop_fix_reports(self):
        worker = self._workers.get("fixreports")
        if worker and worker.isRunning():
            worker.cancel()
            self.fix_btn_stop.setEnabled(False)
            self._fix_add_log("Остановка… текущий файл будет дописан", "warn")

    def _on_fix_reports_finished(self, success: bool):
        self.fix_btn_start.setEnabled(True)
        self.fix_btn_stop.setEnabled(False)
        self._workers.pop("fixreports", None)
