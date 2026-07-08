# -*- coding: utf-8 -*-
"""
Главное окно приложения ПНОС — PySide6.

Структура: единый конвейер из 7 карточек-шагов на странице «Процесс»,
отдельные страницы — форма ручного контроля (шаг 3), «Обзор» (матрица
статусов), «Итоговый отчёт» (таблицы результатов) и «Настройки».
"""

import os
import sys
import time

from PySide6.QtCore import Qt, QSize, QSettings, QTimer
from PySide6.QtGui import QFont, QIcon, QColor
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QProgressBar, QFrame, QLineEdit,
    QFileDialog, QTextEdit, QMessageBox, QCheckBox,
    QSizePolicy, QApplication, QStackedWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QListWidget,
    QListWidgetItem, QScrollArea,
)

from config import CONFIG
from utils import resource_path
from version import APP_VERSION
from pages.settings_page import SettingsPage
from widgets import StepCard, dot_icon, STATUS_COLORS
# Воркеры этапов импортируются лениво (внутри методов запуска),
# чтобы не тянуть pandas / aiohttp / win32com / python-docx при старте.

import subprocess


# Цвета категорий сообщений единого лога
LOG_COLORS = {
    "info": "#94a3b8",
    "done": "#10b981",
    "wait": "#6366f1",
    "warn": "#f59e0b",
    "error": "#ef4444",
}


class MainWindow(QMainWindow):
    """Главное окно приложения ПНОС."""

    # Индексы страниц в QStackedWidget
    PAGE_PROCESS = 0
    PAGE_MANUAL = 1
    PAGE_OVERVIEW = 2   # объединённая страница «Обзор и отчёт»
    PAGE_SETTINGS = 3

    # UI-шаг -> ключ статуса в .pnos_status.json (учитывает сдвиг воркеров)
    STEP_STATUS_KEYS = {1: "stage1", 2: "stage2", 3: "manual3",
                        4: "stage3", 5: "stage4", 6: "stage5", 7: "stage6"}

    STEP_TITLES = {
        1: "Создание структуры",
        2: "Копирование таблиц",
        3: "Ручной контроль",
        4: "Запуск макросов",
        5: "Создание отчёта",
        6: "Вставка картинок",
        7: "Вставка актов ГИ",
    }

    STEP_DESCS = {
        1: "Скачивание папок пунктов с Яндекс Диска и копирование шаблонов",
        2: "Перенос таблицы №3 из «Акта замеров» на листы «Импорт»",
        3: "Заполнение данных паспорта в форме — без ручного Excel",
        4: "Перенос переменных из паспорта Word и запуск макросов Restore/Save",
        5: "Слияние, спецификация, остаточный ресурс — финальный Word-отчёт",
        6: "Вставка фотографий в отчёт по меткам [ФОТО]",
        7: "Вставка актов гидроиспытаний по метке [ГИ] и пересчёт содержания",
    }

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ПНОС — Пакетный генератор отчётов")
        self.setMinimumSize(1000, 680)
        self.resize(1280, 820)

        self._workers = {}
        self._stage_results = {}     # {ui_step: {"columns": [...], "details": [...]} | dict}
        self._completed_stages = set()
        self._start_time = None
        self._is_stopping = False
        self._step_started = {}      # ui_step -> time.time() старта

        self._settings = QSettings("PNOS", "PNOS")
        self._theme_name = self._settings.value("theme", "light")

        self._session_timer = QTimer(self)
        self._session_timer.setInterval(1000)
        self._session_timer.timeout.connect(self._tick_session_timer)

        self.setWindowIcon(QIcon(resource_path(os.path.join("resources", "icon.ico"))))
        self._setup_ui()

    # ═══════════════════════════════════════════════════════════════
    # Каркас интерфейса
    # ═══════════════════════════════════════════════════════════════

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QHBoxLayout(central)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # ─── Сайдбар ───────────────────────────────────────────────
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(230)
        sb = QVBoxLayout(sidebar)
        sb.setContentsMargins(0, 0, 0, 0)
        sb.setSpacing(0)

        logo_frame = QFrame()
        logo_frame.setObjectName("logo_frame")
        logo_layout = QVBoxLayout(logo_frame)
        logo_layout.setContentsMargins(20, 26, 20, 26)
        app_title = QLabel("ПНОС")
        app_title.setObjectName("sidebar_title")
        logo_layout.addWidget(app_title)
        app_sub = QLabel(f"Версия {APP_VERSION}")
        app_sub.setObjectName("sidebar_subtitle")
        logo_layout.addWidget(app_sub)
        sb.addWidget(logo_frame)

        nav_specs = [
            ("Процесс", self.PAGE_PROCESS),
            ("Ручной контроль", self.PAGE_MANUAL),
            ("Обзор и отчёт", self.PAGE_OVERVIEW),
        ]
        self._nav_buttons = {}
        nav_box = QVBoxLayout()
        nav_box.setContentsMargins(0, 12, 0, 0)
        nav_box.setSpacing(2)
        for caption, page_idx in nav_specs:
            btn = QPushButton(caption)
            btn.setObjectName("nav_button")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda checked, i=page_idx: self._switch_page(i))
            self._nav_buttons[page_idx] = btn
            nav_box.addWidget(btn)
        nav_holder = QWidget()
        nav_holder.setLayout(nav_box)
        sb.addWidget(nav_holder)
        sb.addStretch()

        btn_settings = QPushButton("Настройки")
        btn_settings.setObjectName("nav_button")
        btn_settings.setCheckable(True)
        btn_settings.setCursor(Qt.PointingHandCursor)
        btn_settings.clicked.connect(lambda: self._switch_page(self.PAGE_SETTINGS))
        self._nav_buttons[self.PAGE_SETTINGS] = btn_settings
        sb.addWidget(btn_settings)
        sb.addSpacing(8)

        # Общий прогресс + таймер
        prog_holder = QWidget()
        pv = QVBoxLayout(prog_holder)
        pv.setContentsMargins(20, 8, 20, 22)
        head_row = QHBoxLayout()
        lbl_global = QLabel("Общий прогресс")
        lbl_global.setObjectName("global_caption")
        head_row.addWidget(lbl_global)
        head_row.addStretch()
        self.lbl_session_timer = QLabel("")
        self.lbl_session_timer.setObjectName("session_timer")
        head_row.addWidget(self.lbl_session_timer)
        pv.addLayout(head_row)
        self.global_progress = QProgressBar()
        self.global_progress.setObjectName("global_progress")
        self.global_progress.setRange(0, 100)
        self.global_progress.setValue(0)
        self.global_progress.setFixedHeight(10)
        self.global_progress.setTextVisible(False)
        pv.addWidget(self.global_progress)
        sb.addWidget(prog_holder)

        main_layout.addWidget(sidebar)

        # ─── Контент ───────────────────────────────────────────────
        content_frame = QFrame()
        content_frame.setObjectName("content_frame")
        content_layout = QVBoxLayout(content_frame)
        content_layout.setContentsMargins(24, 20, 24, 20)

        wrapper = QHBoxLayout()
        wrapper.addStretch()
        self.pages_stack = QStackedWidget()
        self.pages_stack.setMaximumWidth(1000)
        wrapper.addWidget(self.pages_stack, 1)
        wrapper.addStretch()
        content_layout.addLayout(wrapper)

        # Страницы (порядок = индексы PAGE_*)
        self.page_process = QWidget()
        self._setup_process_page()
        self.pages_stack.addWidget(self.page_process)

        self.page_manual = QWidget()
        self._setup_manual_page()
        self.pages_stack.addWidget(self.page_manual)

        self.page_overview = QWidget()
        self._setup_overview_page()
        self.pages_stack.addWidget(self.page_overview)

        self.page_settings = SettingsPage()
        self.pages_stack.addWidget(self.page_settings)
        self.page_settings.theme_changed.connect(self.apply_theme)

        main_layout.addWidget(content_frame)
        self._switch_page(self.PAGE_PROCESS)

    def _switch_page(self, index: int):
        self.pages_stack.setCurrentIndex(index)
        for idx, btn in self._nav_buttons.items():
            btn.setChecked(idx == index)
        if index == self.PAGE_OVERVIEW:
            self._refresh_overview()
            self._refresh_report_page()
        elif index == self.PAGE_MANUAL:
            self._load_manual_points()

    def _page_title(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("app_title")
        return lbl

    # ═══════════════════════════════════════════════════════════════
    # Страница «Процесс» — конвейер шагов
    # ═══════════════════════════════════════════════════════════════

    def _setup_process_page(self):
        outer = QVBoxLayout(self.page_process)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        body = QWidget()
        v = QVBoxLayout(body)
        v.setContentsMargins(2, 2, 8, 2)
        v.setSpacing(10)
        scroll.setWidget(body)
        outer.addWidget(scroll)

        v.addWidget(self._page_title("Процесс"))

        # ── Параметры запуска ──
        params = QFrame()
        params.setObjectName("settings_panel")
        g = QGridLayout(params)
        g.setContentsMargins(14, 12, 14, 12)
        g.setHorizontalSpacing(10)
        g.setVerticalSpacing(8)

        lbl_pts = QLabel("Пункты")
        lbl_pts.setAlignment(Qt.AlignTop)
        g.addWidget(lbl_pts, 0, 0)
        self.text_manual = QTextEdit()
        self.text_manual.setObjectName("manual_input")
        self.text_manual.setPlaceholderText(
            "Номера пунктов через пробел, запятую или с новой строки. Пример: 2758, 2759, 2760")
        self.text_manual.setFixedHeight(56)
        g.addWidget(self.text_manual, 0, 1, 1, 2)

        g.addWidget(QLabel("Локальная папка"), 1, 0)
        self.entry_local = QLineEdit()
        desktop = os.path.join(os.path.expanduser("~"), "Desktop", "ПНОС")
        self.entry_local.setText(desktop)
        g.addWidget(self.entry_local, 1, 1)
        btn_browse = QPushButton("📂")
        btn_browse.setObjectName("browse_btn")
        btn_browse.setFixedWidth(40)
        btn_browse.setCursor(Qt.PointingHandCursor)
        btn_browse.clicked.connect(self._browse_local)
        g.addWidget(btn_browse, 1, 2)
        v.addWidget(params)

        # ── Карточки шагов ──
        self.step_cards = {}
        actions = {
            1: self._start_stage1, 2: self._start_stage2,
            3: lambda: self._switch_page(self.PAGE_MANUAL),
            4: self._start_stage3, 5: self._start_stage4,
            6: self._start_stage5, 7: self._start_stage6,
        }
        for step in range(1, 8):
            card = StepCard(
                step, self.STEP_TITLES[step], self.STEP_DESCS[step],
                action_text="Заполнить форму" if step == 3 else "Запустить")
            card.action_clicked.connect(actions[step])
            card.stop_clicked.connect(self._force_stop_active_stage)
            self.step_cards[step] = card
            v.addWidget(card)

        # ── Единый журнал ──
        lbl_log = QLabel("Журнал")
        lbl_log.setObjectName("section_label")
        v.addWidget(lbl_log)
        self.activity_list = QListWidget()
        self.activity_list.setObjectName("activity_feed")
        self.activity_list.setMinimumHeight(140)
        v.addWidget(self.activity_list, 1)

    # ═══════════════════════════════════════════════════════════════
    # Журнал и лог
    # ═══════════════════════════════════════════════════════════════

    def _add_activity(self, step: int, message: str, category: str = "info"):
        """Добавляет строку в единый журнал. step=0 — общесистемное сообщение."""
        prefix = f"Шаг {step} · " if step else ""
        item = QListWidgetItem(f"{prefix}{message}")
        item.setIcon(dot_icon(LOG_COLORS.get(category, LOG_COLORS["info"])))
        self.activity_list.addItem(item)
        self.activity_list.scrollToBottom()
        if category == "error":
            self._log_error(message)

    def _log_error(self, details: str):
        try:
            log_dir = os.path.join(os.getcwd(), "logs")
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            log_file = os.path.join(log_dir, "app_errors.log")
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] {details}\n")
        except Exception:
            pass

    def _append_log(self, text: str):
        print(text)
        if "ошибка" in text.lower() or "error" in text.lower() or "❌" in text:
            self._log_error(text)

    # ═══════════════════════════════════════════════════════════════
    # Вспомогательные
    # ═══════════════════════════════════════════════════════════════

    def _browse_local(self):
        folder = QFileDialog.getExistingDirectory(self, "Выбрать папку")
        if folder:
            self.entry_local.setText(folder)

    def _get_local_root(self) -> str:
        path = self.entry_local.text().strip()
        if not path:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop", "ПНОС")
            self.entry_local.setText(desktop)
            return desktop
        return path

    def _tick_session_timer(self):
        if not self._start_time:
            return
        elapsed = int(time.time() - self._start_time)
        m, s = divmod(elapsed, 60)
        h, m = divmod(m, 60)
        self.lbl_session_timer.setText(
            f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}")

    def _update_global_progress(self):
        self.global_progress.setValue(int(len(self._completed_stages) * (100.0 / 7.0)))

    @staticmethod
    def _fmt_duration(seconds: int) -> str:
        m, s = divmod(int(seconds), 60)
        h, m = divmod(m, 60)
        return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"

    def apply_theme(self, theme_name: str = None):
        from theme import build_stylesheet
        if theme_name:
            self._theme_name = theme_name
            self._settings.setValue("theme", theme_name)
        app = QApplication.instance()
        if app:
            app.setStyleSheet(build_stylesheet(self._theme_name))

    # ═══════════════════════════════════════════════════════════════
    # Управление шагами (карточками)
    # ═══════════════════════════════════════════════════════════════

    def _begin_step(self, step: int):
        """Переводит карточку шага в состояние «выполняется»."""
        card = self.step_cards[step]
        card.set_status("run")
        card.set_progress(0)
        card.set_summary("Выполняется…")
        if step == 4:
            card.set_summary("Выполняется… Не открывайте другие файлы Excel и не перехватывайте фокус мыши")
        self._step_started[step] = time.time()

    def _summarize_details(self, details: list):
        """Считает (готово, ошибки, пропущено) по списку строк-результатов."""
        done = errors = skipped = 0
        for d in details:
            res = str(d.get("Результат", d.get("status", ""))).lower()
            if "пропущен" in res or "⚪" in res:
                skipped += 1
            elif "успешно" in res or "✅" in res or "✔" in res or "частично" in res:
                done += 1
            elif res:
                errors += 1
            else:
                done += 1
        return done, errors, skipped

    def _store_stage_result(self, step: int, columns: list, details: list):
        """Сохраняет результат этапа для страницы «Итоговый отчёт» и карточки."""
        self._stage_results[step] = {"columns": columns, "details": details}
        done, errors, skipped = self._summarize_details(details)
        parts = [f"Готово {done} из {len(details)}"]
        if errors:
            parts.append(f"ошибок {errors}")
        if skipped:
            parts.append(f"пропущено {skipped}")
        started = self._step_started.get(step)
        if started:
            parts.append(self._fmt_duration(time.time() - started))
        self.step_cards[step].set_summary(" · ".join(parts))

    def _on_step_finished(self, step: int, success: bool):
        """Общее завершение шага: карточка, авто-цепочка, обзор."""
        card = self.step_cards[step]
        card.set_status("done" if success else "error")
        if success:
            self._completed_stages.add(step)
            self._add_activity(step, "Этап завершён", "done")
        else:
            self._add_activity(step, "Этап завершён с ошибками", "error")
        self._update_global_progress()

        if success:
            if step >= 7:
                # Конвейер завершён — показываем «Обзор и отчёт»
                self._session_timer.stop()
                self._switch_page(self.PAGE_OVERVIEW)
                return
            next_step = step + 1
            if self.page_settings.chk_auto.isChecked():
                if next_step == 3:
                    # Ручной шаг: подсветить и открыть форму
                    self.step_cards[3].set_status("wait")
                    self.step_cards[3].set_summary("Заполните форму и запишите данные в Excel")
                    self._switch_page(self.PAGE_MANUAL)
                else:
                    QTimer.singleShot(800, self.step_cards[next_step].btn_action.click)

    # ═══════════════════════════════════════════════════════════════
    # Запуск этапов
    # ═══════════════════════════════════════════════════════════════

    def _start_stage1(self):
        local_root = self._get_local_root()

        import re
        raw_points = re.findall(r"\d+", self.text_manual.toPlainText())
        manual_points = []
        for p in raw_points:
            if p not in manual_points:
                manual_points.append(p)
        if not manual_points:
            QMessageBox.warning(self, "Ошибка", "Укажите хотя бы один номер пункта.")
            return

        self._start_time = time.time()
        self._session_timer.start()
        self._stage_results = {}
        self._begin_step(1)
        self._add_activity(1, f"Запуск: {len(manual_points)} пунктов", "wait")

        card = self.step_cards[1]
        from workers.stage1_worker import Stage1Worker
        # «Стар. ЭПБ» ищется всегда (её отсутствие не блокирует пункт),
        # скорость сети всегда максимальная.
        worker = Stage1Worker(
            excel_path=None, local_root=local_root, manual_points=manual_points,
            need_epb=True, max_threads=10, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(1, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(self._on_stage1_report)
        worker.finished_ok.connect(lambda ok: self._on_step_finished(1, ok))
        worker.start()
        self._workers["stage1"] = worker

    def _on_stage1_report(self, report: dict):
        """Результат этапа 1 (dict) — приводим к общему виду."""
        details = []
        for d in report.get("details", []):
            folders = d.get("folders", {})
            optional = d.get("optional", {})
            details.append({
                "№ Пункта": str(d.get("point", "")),
                "Паспорт": folders.get("Паспорт", False),
                "Первичка": folders.get("Первичка", False),
                "Стар. ЭПБ": folders.get("Стар. ЭПБ", False),
                "Исполнит.": optional.get("Исполнительная", False),
                "Акт отбр.": optional.get("Акт отбраковки", False),
                "Результат": d.get("status", ""),
            })
        columns = ["№ Пункта", "Паспорт", "Первичка", "Стар. ЭПБ",
                   "Исполнит.", "Акт отбр.", "Результат"]
        self._store_stage_result(1, columns, details)

    def _start_stage2(self):
        local_root = self._get_local_root()
        self._begin_step(2)
        card = self.step_cards[2]
        from workers.stage2_worker import Stage2Worker
        worker = Stage2Worker(local_root, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(2, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(2, ["№ Пункта", "Акт найден", "Таблица 3 найдена", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(2, ok))
        worker.start()
        self._workers["stage2"] = worker

    def _start_stage3(self):
        local_root = self._get_local_root()
        self._begin_step(4)
        card = self.step_cards[4]
        macro_path = self.page_settings.edit_macro_path.text().strip() or None
        from workers.stage3_worker import Stage3Worker
        worker = Stage3Worker(local_root, macro_master_path=macro_path, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(4, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(4, ["№ Пункта", "Паспорт", "Excel", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(4, ok))
        worker.start()
        self._workers["stage3"] = worker

    def _start_stage4(self):
        local_root = self._get_local_root()
        self._begin_step(5)
        card = self.step_cards[5]
        from workers.stage4_worker import Stage4Worker
        worker = Stage4Worker(local_root, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(5, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(5, ["№ Пункта", "Макрос Excel", "Слияние Word", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(5, ok))
        worker.start()
        self._workers["stage4"] = worker

    def _start_stage5(self):
        local_root = self._get_local_root()
        self._begin_step(6)
        card = self.step_cards[6]
        from workers.stage5_worker import Stage5Worker
        worker = Stage5Worker(local_root, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(6, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(6, ["№ Пункта", "Фото", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(6, ok))
        worker.start()
        self._workers["stage5"] = worker

    def _start_stage6(self):
        local_root = self._get_local_root()
        self._begin_step(7)
        card = self.step_cards[7]
        from workers.stage6_worker import Stage6Worker
        worker = Stage6Worker(local_root, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(7, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(7, ["№ Пункта", "Акты ГИ", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(7, ok))
        worker.start()
        self._workers["stage6"] = worker

    # ═══════════════════════════════════════════════════════════════
    # Принудительная остановка
    # ═══════════════════════════════════════════════════════════════

    def _force_stop_active_stage(self):
        if self._is_stopping:
            return
        self._is_stopping = True
        try:
            self._append_log("\n🛑 ПРИНУДИТЕЛЬНАЯ ОСТАНОВКА ПОЛЬЗОВАТЕЛЕМ...")
            self._add_activity(0, "Принудительная остановка", "warn")

            for worker in list(self._workers.values()):
                if worker:
                    try:
                        worker.disconnect()
                        if worker.isRunning():
                            worker.cancel()
                            worker.quit()
                            for _ in range(15):
                                if worker.wait(100):
                                    break
                                QApplication.processEvents()
                    except Exception as e:
                        print(f"Ошибка при отключении воркера: {e}")
            self._workers = {}

            for proc in ["EXCEL.EXE", "WINWORD.EXE"]:
                try:
                    subprocess.Popen(f"taskkill /F /IM {proc}", shell=True,
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception:
                    pass

            self._session_timer.stop()
            for card in self.step_cards.values():
                if card.status() == "run":
                    card.set_status("idle")
                    card.reset_summary()
            # Кнопка записи на форме шага 3
            if hasattr(self, "manual_btn_write"):
                self.manual_btn_write.setEnabled(True)
                self.manual_btn_write.setText("Записать всё в Excel")

            self._add_activity(0, "Система готова к новому запуску", "done")
        except Exception as e:
            msg = f"Ошибка в процессе остановки: {e}"
            print(msg)
            self._log_error(msg)
        finally:
            self._is_stopping = False

    # ═══════════════════════════════════════════════════════════════
    # Страница «Ручной контроль» (шаг 3)
    # ═══════════════════════════════════════════════════════════════

    def _make_date_edit(self):
        from PySide6.QtWidgets import QDateEdit
        from PySide6.QtCore import QDate
        de = QDateEdit()
        de.setCalendarPopup(True)
        de.setDisplayFormat("dd.MM.yyyy")
        de.setMinimumDate(QDate(1900, 1, 1))
        de.setSpecialValueText("— не выбрана —")
        de.setDate(QDate(1900, 1, 1))
        return de

    def _date_edit_value(self, de) -> str:
        from PySide6.QtCore import QDate
        if de.date() == QDate(1900, 1, 1):
            return ""
        return de.date().toString("dd.MM.yyyy")

    def _set_date_edit(self, de, value: str):
        from PySide6.QtCore import QDate
        s = (value or "").strip()
        d = QDate.fromString(s, "dd.MM.yyyy") if s else QDate()
        de.setDate(d if d.isValid() else QDate(1900, 1, 1))

    def _form_field(self, label_text: str, widget, required: bool = False, note: str = ""):
        block = QWidget()
        v = QVBoxLayout(block)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(4)
        head = QHBoxLayout()
        head.setSpacing(4)
        lbl = QLabel(label_text)
        lbl.setObjectName("form_field_label")
        head.addWidget(lbl)
        if required:
            star = QLabel("*")
            star.setObjectName("form_field_req")
            head.addWidget(star)
        head.addStretch()
        v.addLayout(head)
        v.addWidget(widget)
        if note:
            n = QLabel(note)
            n.setObjectName("field_note")
            n.setWordWrap(True)
            v.addWidget(n)
        return block

    def _setup_manual_page(self):
        from PySide6.QtWidgets import QComboBox
        from PySide6.QtGui import QIntValidator

        self.manual_drafts = {}
        self.manual_current = None
        self._manual_loading = False

        outer = QVBoxLayout(self.page_manual)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(10)

        head_row = QHBoxLayout()
        head_col = QVBoxLayout()
        head_col.setSpacing(2)
        head_col.addWidget(self._page_title("Ручной контроль"))
        sub = QLabel("Заполните данные паспорта для каждого пункта — они будут записаны "
                     "на лист «паспорт» автоматически.")
        sub.setObjectName("card_description")
        sub.setWordWrap(True)
        head_col.addWidget(sub)
        head_row.addLayout(head_col, 1)
        btn_back = QPushButton("К процессу")
        btn_back.setObjectName("ghost_button")
        btn_back.setCursor(Qt.PointingHandCursor)
        btn_back.clicked.connect(lambda: self._switch_page(self.PAGE_PROCESS))
        head_row.addWidget(btn_back, 0, Qt.AlignTop)
        outer.addLayout(head_row)

        split = QHBoxLayout()
        split.setSpacing(12)

        left = QVBoxLayout()
        left.setSpacing(6)
        lbl_pts = QLabel("Пункты")
        lbl_pts.setObjectName("section_label")
        left.addWidget(lbl_pts)
        self.manual_point_list = QListWidget()
        self.manual_point_list.setObjectName("point_list")
        self.manual_point_list.setMinimumWidth(230)
        self.manual_point_list.setMaximumWidth(280)
        self.manual_point_list.currentItemChanged.connect(self._on_manual_point_selected)
        left.addWidget(self.manual_point_list, 1)
        legend = QLabel(
            f'<span style="color:{STATUS_COLORS["done"]}">●</span> заполнен&nbsp;&nbsp;'
            f'<span style="color:{STATUS_COLORS["idle"]}">●</span> не заполнен&nbsp;&nbsp;'
            f'<span style="color:{STATUS_COLORS["wait"]}">●</span> пропущен&nbsp;&nbsp;'
            f'<span style="color:{STATUS_COLORS["error"]}">●</span> не корректный')
        legend.setObjectName("field_note")
        legend.setWordWrap(True)
        left.addWidget(legend)
        split.addLayout(left)

        form_scroll = QScrollArea()
        form_scroll.setWidgetResizable(True)
        form_scroll.setFrameShape(QFrame.NoFrame)
        form_card = QFrame()
        form_card.setObjectName("form_card")
        form = QVBoxLayout(form_card)
        form.setContentsMargins(20, 16, 20, 16)
        form.setSpacing(12)

        self.manual_form_title = QLabel("Выберите пункт слева")
        self.manual_form_title.setObjectName("form_title")
        form.addWidget(self.manual_form_title)

        f = {}
        f["purpose"] = QLineEdit()
        f["purpose"].setPlaceholderText("транспортировки …")
        form.addWidget(self._form_field(
            "Предназначение трубопровода (B14)", f["purpose"], required=True,
            note="Начните с «транспортировки», далее — конкретное вещество "
                 "(напр. «транспортировки рафината растворителя»)."))

        f["specialist"] = QComboBox()
        f["specialist"].addItem("— выберите —", "")
        for name in CONFIG.get("SPECIALISTS", []):
            f["specialist"].addItem(name, name)
        form.addWidget(self._form_field("Специалист (B22)", f["specialist"], required=True))

        f["date"] = self._make_date_edit()
        form.addWidget(self._form_field("Дата (B33)", f["date"], required=True,
                                        note="Формат ДД.ММ.ГГГГ, выбор через календарь."))

        sep = QLabel("Данные предыдущей ЭПБ")
        sep.setObjectName("section_label")
        form.addWidget(sep)

        f["org"] = QComboBox()
        f["org"].addItem("— не выбрано —", "")
        for org in CONFIG.get("EPB_ORGS", []):
            f["org"].addItem(org, org)
        f["org"].currentIndexChanged.connect(self._update_epb_done_label)
        form.addWidget(self._form_field(
            "Организация (B52)", f["org"], required=False,
            note="Необязательно. Если экспертиза ранее не проводилась — оставьте пустым."))

        self.manual_epb_done = QLabel("Проводилась ранее (B51): нет")
        self.manual_epb_done.setObjectName("field_note")
        form.addWidget(self.manual_epb_done)

        f["reg_num"] = QLineEdit()
        f["reg_num"].setPlaceholderText("Рег. номер Ростехнадзора или внутренний")
        form.addWidget(self._form_field("Регистрационный номер (B53)", f["reg_num"], required=False))

        epb_date_row = QWidget()
        edr = QHBoxLayout(epb_date_row)
        edr.setContentsMargins(0, 0, 0, 0)
        edr.setSpacing(6)
        f["epb_date"] = self._make_date_edit()
        edr.addWidget(f["epb_date"], 1)
        btn_clr = QPushButton("Очистить")
        btn_clr.setObjectName("ghost_button")
        btn_clr.setCursor(Qt.PointingHandCursor)
        btn_clr.clicked.connect(lambda: self._set_date_edit(f["epb_date"], ""))
        edr.addWidget(btn_clr)
        form.addWidget(self._form_field("Дата предыдущей ЭПБ (B54)", epb_date_row, required=False))

        f["sheets"] = QLineEdit()
        f["sheets"].setValidator(QIntValidator(0, 100000, self))
        f["sheets"].setPlaceholderText("Число листов заключения")
        form.addWidget(self._form_field(
            "Количество листов (B55)", f["sheets"], required=False,
            note="Лист уведомления Ростехнадзора не учитывается."))

        form.addStretch()
        self.manual_fields = f
        form_scroll.setWidget(form_card)
        split.addWidget(form_scroll, 1)
        outer.addLayout(split, 1)

        # Кнопки действий
        btn_row = QHBoxLayout()
        self.manual_btn_write = QPushButton("Записать всё в Excel")
        self.manual_btn_write.setObjectName("card_button")
        self.manual_btn_write.setMinimumWidth(200)
        self.manual_btn_write.setCursor(Qt.PointingHandCursor)
        self.manual_btn_write.clicked.connect(self._start_manual_stage3)

        btn_skip = QPushButton("Пропустить пункт")
        btn_skip.setObjectName("ghost_button")
        btn_skip.setCursor(Qt.PointingHandCursor)
        btn_skip.clicked.connect(self._skip_current_point)

        btn_excel = QPushButton("Открыть в Excel")
        btn_excel.setObjectName("ghost_button")
        btn_excel.setCursor(Qt.PointingHandCursor)
        btn_excel.clicked.connect(self._open_current_in_excel)

        self.manual_btn_stop = QPushButton("Остановить")
        self.manual_btn_stop.setObjectName("stop_button")
        self.manual_btn_stop.setCursor(Qt.PointingHandCursor)
        self.manual_btn_stop.setEnabled(False)
        self.manual_btn_stop.clicked.connect(self._force_stop_active_stage)

        btn_row.addWidget(self.manual_btn_write)
        btn_row.addWidget(btn_skip)
        btn_row.addWidget(btn_excel)
        btn_row.addStretch()
        btn_row.addWidget(self.manual_btn_stop)
        outer.addLayout(btn_row)

        self.manual_progress = QProgressBar()
        self.manual_progress.setValue(0)
        self.manual_progress.setFixedHeight(8)
        self.manual_progress.setTextVisible(False)
        outer.addWidget(self.manual_progress)

    def _load_manual_points(self):
        points = self._scan_points()
        self.manual_point_list.blockSignals(True)
        self.manual_point_list.clear()
        state_colors = {"invalid": STATUS_COLORS["error"], "skipped": STATUS_COLORS["wait"],
                        "done": STATUS_COLORS["done"], "filled": STATUS_COLORS["done"],
                        "empty": STATUS_COLORS["idle"]}
        for pt in points:
            folder = os.path.join(self._get_local_root(), pt["name"])
            _, state = self._manual_point_state(folder, pt)
            item = QListWidgetItem(pt["name"])
            item.setIcon(dot_icon(state_colors.get(state, STATUS_COLORS["idle"]), 12))
            item.setData(Qt.UserRole, folder)
            item.setData(Qt.UserRole + 1, state)
            if pt["invalid"]:
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
            self.manual_point_list.addItem(item)
        self.manual_point_list.blockSignals(False)

        # Строка итога на карточке шага 3
        filled = sum(1 for i in range(self.manual_point_list.count())
                     if self.manual_point_list.item(i).data(Qt.UserRole + 1) in ("done", "filled"))
        skipped = sum(1 for i in range(self.manual_point_list.count())
                      if self.manual_point_list.item(i).data(Qt.UserRole + 1) == "skipped")
        total = self.manual_point_list.count()
        if total:
            self.step_cards[3].set_summary(
                f"Заполнено {filled} из {total}" + (f" · пропущено {skipped}" if skipped else ""))

        target = None
        for i in range(self.manual_point_list.count()):
            it = self.manual_point_list.item(i)
            if it.flags() & Qt.ItemIsEnabled:
                if target is None:
                    target = i
                if it.data(Qt.UserRole + 1) == "empty":
                    target = i
                    break
        if target is not None:
            self.manual_point_list.setCurrentRow(target)
        else:
            self.manual_current = None
            self._clear_manual_form()

    def _manual_point_state(self, folder, pt):
        if pt["invalid"]:
            return None, "invalid"
        from status_manager import get_status_value
        if pt.get("skipped") or get_status_value(folder, "skipped", False):
            return None, "skipped"
        if folder in self.manual_drafts:
            values = self.manual_drafts[folder]
        else:
            values = get_status_value(folder, "manual3_draft", None)
        if get_status_value(folder, "manual3", False):
            return None, "done"
        if values and self._required_ok(values):
            return None, "filled"
        return None, "empty"

    def _required_ok(self, values: dict) -> bool:
        purpose = (values.get("purpose") or "").strip()
        specialist = (values.get("specialist") or "").strip()
        date = (values.get("date") or "").strip()
        return bool(purpose) and bool(specialist) and bool(date)

    def _on_manual_point_selected(self, current, previous):
        if self._manual_loading:
            return
        if self.manual_current is not None:
            self._save_current_draft()
        if current is None:
            self.manual_current = None
            self._clear_manual_form()
            return
        folder = current.data(Qt.UserRole)
        self.manual_current = folder
        self._load_form_for(folder)

    def _load_form_for(self, folder):
        from status_manager import get_status_value
        self._manual_loading = True
        try:
            self.manual_form_title.setText(f"Пункт {os.path.basename(folder)}")
            if folder in self.manual_drafts:
                values = self.manual_drafts[folder]
            else:
                values = get_status_value(folder, "manual3_draft", None) or {}
            prefix = CONFIG["MANUAL_FORM"].get("PURPOSE_PREFIX", "транспортировки ")
            f = self.manual_fields
            f["purpose"].setText(values.get("purpose") or prefix)
            self._set_combo(f["specialist"], values.get("specialist", ""))
            self._set_date_edit(f["date"], values.get("date", ""))
            self._set_combo(f["org"], values.get("epb_org", ""))
            f["reg_num"].setText(values.get("reg_num", ""))
            self._set_date_edit(f["epb_date"], values.get("epb_date", ""))
            f["sheets"].setText(str(values.get("sheets", "") or ""))
            self._update_epb_done_label()
        finally:
            self._manual_loading = False

    def _clear_manual_form(self):
        self._manual_loading = True
        try:
            self.manual_form_title.setText("Выберите пункт слева")
            f = self.manual_fields
            prefix = CONFIG["MANUAL_FORM"].get("PURPOSE_PREFIX", "транспортировки ")
            f["purpose"].setText(prefix)
            f["specialist"].setCurrentIndex(0)
            self._set_date_edit(f["date"], "")
            f["org"].setCurrentIndex(0)
            f["reg_num"].setText("")
            self._set_date_edit(f["epb_date"], "")
            f["sheets"].setText("")
            self._update_epb_done_label()
        finally:
            self._manual_loading = False

    def _set_combo(self, combo, value: str):
        idx = combo.findData(value)
        combo.setCurrentIndex(idx if idx >= 0 else 0)

    def _collect_form_values(self) -> dict:
        f = self.manual_fields
        return {
            "purpose": f["purpose"].text().strip(),
            "specialist": f["specialist"].currentData() or "",
            "date": self._date_edit_value(f["date"]),
            "epb_org": f["org"].currentData() or "",
            "reg_num": f["reg_num"].text().strip(),
            "epb_date": self._date_edit_value(f["epb_date"]),
            "sheets": f["sheets"].text().strip(),
        }

    def _save_current_draft(self):
        if not self.manual_current:
            return
        values = self._collect_form_values()
        self.manual_drafts[self.manual_current] = values
        try:
            from status_manager import set_stage_status
            set_stage_status(self.manual_current, "manual3_draft", values)
        except Exception:
            pass

    def _update_epb_done_label(self, *args):
        org = self.manual_fields["org"].currentData() or ""
        self.manual_epb_done.setText(
            f"Проводилась ранее (B51): {'да' if org.strip() else 'нет'}")
        combo = self.manual_fields["org"]
        combo.setProperty("attention", not bool(org.strip()))
        combo.style().unpolish(combo)
        combo.style().polish(combo)

    def _skip_current_point(self):
        if not self.manual_current:
            return
        from status_manager import set_stage_status
        folder = self.manual_current
        set_stage_status(folder, "skipped", True)
        set_stage_status(folder, "manual3", False)
        self.manual_drafts.pop(folder, None)
        self.manual_current = None
        self._add_activity(3, f"{os.path.basename(folder)} — пропущен", "warn")
        self._load_manual_points()

    def _open_current_in_excel(self):
        if not self.manual_current:
            return
        from workers.stage2_worker import find_file_exact_ext
        path = find_file_exact_ext(self.manual_current, (".xls", ".xlsx", ".xlsm"))
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось открыть файл:\n{e}")
        else:
            QMessageBox.information(self, "Файл не найден",
                                    "ПНОС_ТТП.xlsm в папке пункта не найден.")

    def _start_manual_stage3(self):
        """Валидация формы всех пунктов и запись данных в Excel."""
        from status_manager import get_status_value

        self._save_current_draft()

        points = self._scan_points()
        entries = []
        not_ready = []
        root = self._get_local_root()

        for pt in points:
            if pt["invalid"]:
                continue
            folder = os.path.join(root, pt["name"])
            if pt.get("skipped") or get_status_value(folder, "skipped", False):
                continue
            values = self.manual_drafts.get(folder) or get_status_value(folder, "manual3_draft", None)
            if not values or not self._required_ok(values):
                not_ready.append(pt["name"])
                continue
            entries.append({"folder": folder, "num": pt["name"].replace("п.", ""), "values": values})

        if not_ready:
            QMessageBox.warning(
                self, "Заполнены не все пункты",
                "Не заполнены обязательные поля (Предназначение, Специалист, Дата) "
                "у пунктов:\n\n" + ", ".join(not_ready) +
                "\n\nЗаполните их или нажмите «Пропустить пункт».")
            self._add_activity(3, f"Не заполнены пункты: {', '.join(not_ready)}", "warn")
            return

        if not entries:
            QMessageBox.information(
                self, "Нет данных для записи",
                "Нет заполненных пунктов для записи. Заполните хотя бы один пункт "
                "или пропустите все.")
            return

        self.manual_btn_write.setText("Запись…")
        self.manual_btn_write.setEnabled(False)
        self.manual_btn_stop.setEnabled(True)
        self.manual_progress.setValue(0)
        self._begin_step(3)

        from workers.manual3_worker import Manual3Worker
        worker = Manual3Worker(entries, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(3, msg, cat))
        worker.progress.connect(self.manual_progress.setValue)
        worker.progress.connect(self.step_cards[3].set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(3, ["№ Пункта", "Записано", "Результат"], d))
        worker.finished_ok.connect(self._on_manual_written)
        worker.start()
        self._workers["manual3"] = worker

    def _on_manual_written(self, success: bool):
        self.manual_btn_write.setText("Записать всё в Excel")
        self.manual_btn_write.setEnabled(True)
        self.manual_btn_stop.setEnabled(False)
        self._load_manual_points()
        if success:
            self._switch_page(self.PAGE_PROCESS)
        self._on_step_finished(3, success)

    # ═══════════════════════════════════════════════════════════════
    # Страница «Обзор»
    # ═══════════════════════════════════════════════════════════════

    def _section_header(self, title: str, subtitle: str, button: QPushButton = None):
        """Заголовок секции объединённой страницы: название + пояснение + кнопка."""
        row = QHBoxLayout()
        col = QVBoxLayout()
        col.setSpacing(2)
        lbl = QLabel(title)
        lbl.setObjectName("form_title")
        col.addWidget(lbl)
        sub = QLabel(subtitle)
        sub.setObjectName("card_description")
        sub.setWordWrap(True)
        col.addWidget(sub)
        row.addLayout(col, 1)
        if button is not None:
            row.addWidget(button, 0, Qt.AlignTop)
        return row

    def _setup_overview_page(self):
        """Объединённая страница: сверху «Обзор», снизу «Итоговый отчёт»."""
        outer = QVBoxLayout(self.page_overview)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(2, 2, 8, 2)
        layout.setSpacing(10)
        scroll.setWidget(body)
        outer.addWidget(scroll)

        layout.addWidget(self._page_title("Обзор и отчёт"))

        # ── Секция 1: Обзор — где мы находимся ──
        btn_refresh = QPushButton("Обновить")
        btn_refresh.setObjectName("ghost_button")
        btn_refresh.setCursor(Qt.PointingHandCursor)
        btn_refresh.clicked.connect(self._refresh_overview)
        layout.addLayout(self._section_header(
            "Обзор — где мы находимся",
            "Статусы пунктов по шагам. Читаются с диска и сохраняются между запусками программы.",
            btn_refresh))

        self.lbl_overview_summary = QLabel("")
        self.lbl_overview_summary.setObjectName("card_description")
        layout.addWidget(self.lbl_overview_summary)

        self.overview_table = QTableWidget()
        headers = ["Пункт"] + [str(i) for i in range(1, 8)]
        self.overview_table.setColumnCount(len(headers))
        self.overview_table.setHorizontalHeaderLabels(headers)
        self.overview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.overview_table.setSelectionMode(QTableWidget.NoSelection)
        self.overview_table.verticalHeader().setVisible(False)
        hh = self.overview_table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.Stretch)
        for c in range(1, len(headers)):
            hh.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        layout.addWidget(self.overview_table)

        lbl_problems = QLabel("Проблемные пункты")
        lbl_problems.setObjectName("section_label")
        layout.addWidget(lbl_problems)
        self.overview_problems = QListWidget()
        self.overview_problems.setObjectName("activity_feed")
        self.overview_problems.setMinimumHeight(84)
        self.overview_problems.setMaximumHeight(140)
        layout.addWidget(self.overview_problems)

        # ── Разделитель секций ──
        divider = QFrame()
        divider.setObjectName("h_divider")
        layout.addSpacing(6)
        layout.addWidget(divider)
        layout.addSpacing(6)

        # ── Секция 2: Итоговый отчёт — как прошёл прогон ──
        btn_open = QPushButton("📂 Папка отчётов")
        btn_open.setObjectName("ghost_button")
        btn_open.setCursor(Qt.PointingHandCursor)
        btn_open.clicked.connect(self._open_reports_folder)
        layout.addLayout(self._section_header(
            "Итоговый отчёт — как прошёл прогон",
            "Детали текущего запуска программы по каждому шагу. Очищается при перезапуске.",
            btn_open))

        self.lbl_report_summary = QLabel("Отчёт появится после выполнения этапов.")
        self.lbl_report_summary.setObjectName("card_description")
        layout.addWidget(self.lbl_report_summary)

        self._report_sections = QVBoxLayout()
        self._report_sections.setSpacing(10)
        layout.addLayout(self._report_sections)
        layout.addStretch()

    def _scan_points(self):
        """Сканирует локальную папку: [{name, invalid, missing, skipped, steps{}}]."""
        import json
        root = self._get_local_root()
        result = []
        if not root or not os.path.isdir(root):
            return result
        for name in sorted(os.listdir(root)):
            folder = os.path.join(root, name)
            if not os.path.isdir(folder):
                continue
            low = name.lower()
            if not (low.startswith("п.") or low.startswith("п ") or low.startswith("пункт")):
                continue
            status_path = os.path.join(folder, ".pnos_status.json")
            data = {}
            if os.path.exists(status_path):
                try:
                    with open(status_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception:
                    data = {}
            steps = {ui: bool(data.get(key, False))
                     for ui, key in self.STEP_STATUS_KEYS.items()}
            result.append({
                "name": name,
                "invalid": bool(data.get("invalid", False)) or ("не корректн" in low),
                "missing": data.get("missing_folders", []),
                "skipped": bool(data.get("skipped", False)),
                "steps": steps,
            })
        return result

    def _refresh_overview(self):
        points = self._scan_points()
        self.overview_table.setRowCount(len(points))
        self.overview_problems.clear()

        c_done = QColor(STATUS_COLORS["done"])
        c_wait = QColor(STATUS_COLORS["idle"])
        c_skip = QColor(STATUS_COLORS["wait"])
        c_bad = QColor(STATUS_COLORS["error"])

        done_all = 0
        for r, pt in enumerate(points):
            self.overview_table.setItem(r, 0, QTableWidgetItem(pt["name"]))
            invalid, skipped = pt["invalid"], pt["skipped"]
            all_done = True
            for ui_step in range(1, 8):
                if invalid:
                    glyph, color = "✕", c_bad
                    all_done = False
                elif skipped and ui_step >= 3:
                    glyph, color = "›", c_skip
                elif pt["steps"].get(ui_step):
                    glyph, color = "✓", c_done
                else:
                    glyph, color = "—", c_wait
                    all_done = False
                cell = QTableWidgetItem(glyph)
                cell.setTextAlignment(Qt.AlignCenter)
                cell.setForeground(color)
                self.overview_table.setItem(r, ui_step, cell)
            if all_done:
                done_all += 1

            if invalid:
                reason = ", ".join(pt["missing"]) if pt["missing"] else "нет обязательных папок"
                it = QListWidgetItem(f"{pt['name']} — не будет обработан ({reason})")
                it.setIcon(dot_icon(STATUS_COLORS["error"]))
                self.overview_problems.addItem(it)
            elif skipped:
                it = QListWidgetItem(f"{pt['name']} — пропущен пользователем")
                it.setIcon(dot_icon(STATUS_COLORS["wait"]))
                self.overview_problems.addItem(it)

        self.lbl_overview_summary.setText(
            f"Всего пунктов: {len(points)} · Полностью готовы: {done_all}")
        if self.overview_problems.count() == 0:
            it = QListWidgetItem("Проблемных пунктов нет")
            it.setIcon(dot_icon(STATUS_COLORS["done"]))
            self.overview_problems.addItem(it)

        # Высота матрицы под содержимое (страница прокручивается целиком)
        h = self.overview_table.horizontalHeader().height() + 2
        for r in range(self.overview_table.rowCount()):
            h += self.overview_table.rowHeight(r)
        self.overview_table.setFixedHeight(max(min(h + 4, 420), 64))

    # ═══════════════════════════════════════════════════════════════
    # Страница «Итоговый отчёт»
    # ═══════════════════════════════════════════════════════════════

    def _open_reports_folder(self):
        root = self._get_local_root()
        if root and os.path.isdir(root):
            try:
                os.startfile(root)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось открыть папку:\n{e}")
        else:
            QMessageBox.information(self, "Папка не найдена", "Локальная папка отчётов не найдена.")

    def _refresh_report_page(self):
        """Пересобирает секции отчёта из накопленных результатов этапов."""
        # Очистка старых секций
        while self._report_sections.count():
            item = self._report_sections.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        if not self._stage_results:
            self.lbl_report_summary.setText("Отчёт появится после выполнения этапов.")
            return

        # Сводка
        total_errors = 0
        for step, res in self._stage_results.items():
            _, errors, _ = self._summarize_details(res.get("details", []))
            total_errors += errors
        elapsed = self._fmt_duration(time.time() - self._start_time) if self._start_time else "—"
        skipped_pts = sum(1 for pt in self._scan_points() if pt["skipped"])
        invalid_pts = sum(1 for pt in self._scan_points() if pt["invalid"])
        summary = f"Время: {elapsed} · Ошибок: {total_errors}"
        if skipped_pts:
            summary += f" · Пропущено пунктов: {skipped_pts}"
        if invalid_pts:
            summary += f" · Не корректных: {invalid_pts}"
        self.lbl_report_summary.setText(summary)

        # Секции по шагам
        for step in sorted(self._stage_results.keys()):
            res = self._stage_results[step]
            columns, details = res.get("columns", []), res.get("details", [])
            if not details:
                continue

            box = QFrame()
            box.setObjectName("form_card")
            bv = QVBoxLayout(box)
            bv.setContentsMargins(14, 12, 14, 12)
            bv.setSpacing(8)

            title = QLabel(f"Шаг {step}: {self.STEP_TITLES.get(step, '')}")
            title.setObjectName("form_field_label")
            bv.addWidget(title)

            table = QTableWidget()
            table.setColumnCount(len(columns))
            table.setHorizontalHeaderLabels(columns)
            table.setRowCount(len(details))
            table.setEditTriggers(QTableWidget.NoEditTriggers)
            table.setSelectionMode(QTableWidget.NoSelection)
            table.verticalHeader().setVisible(False)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.setAlternatingRowColors(True)

            c_ok = QColor(STATUS_COLORS["done"])
            c_bad = QColor(STATUS_COLORS["error"])
            for r, d in enumerate(details):
                for c, col in enumerate(columns):
                    val = d.get(col, "")
                    if isinstance(val, bool):
                        item = QTableWidgetItem("✓" if val else "—")
                        item.setForeground(c_ok if val else QColor(STATUS_COLORS["idle"]))
                    else:
                        text = str(val)
                        item = QTableWidgetItem(text)
                        if col == "Результат":
                            low = text.lower()
                            good = "успешно" in low or "✅" in low or "частично" in low or "пропущен" in low
                            item.setForeground(c_ok if good else c_bad)
                    item.setTextAlignment(Qt.AlignCenter)
                    table.setItem(r, c, item)

            # Высота таблицы под содержимое (без внутренней прокрутки)
            h = table.horizontalHeader().height() + 2
            for r in range(table.rowCount()):
                h += table.rowHeight(r)
            table.setFixedHeight(min(h + 4, 420))
            bv.addWidget(table)
            self._report_sections.addWidget(box)
