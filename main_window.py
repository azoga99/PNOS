# -*- coding: utf-8 -*-
"""
Главное окно приложения ПНОС — PySide6.

Структура: единый конвейер из 7 карточек-шагов на странице «Процесс»,
отдельные страницы — форма ручного контроля (шаг 3), «Обзор и отчёт»
(матрица статусов + таблицы результатов), автономные «Правка отчётов»
и «Анализ пунктов», а также «Настройки».
"""

import os
import sys
import time

from PySide6.QtCore import Qt, QSize, QSettings, QTimer, QThread, Signal
from PySide6.QtGui import QFont, QIcon, QColor
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QProgressBar, QFrame, QLineEdit,
    QFileDialog, QTextEdit, QMessageBox, QCheckBox, QDialog,
    QSizePolicy, QApplication, QStackedWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QListWidget,
    QListWidgetItem, QScrollArea,
)

from config import CONFIG
from utils import resource_path, title_date_text, find_order_scan
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


class MasterCheckWorker(QThread):
    """Фоновая проверка доступности сводного графика при старте.

    Сетевой путь (S:\\…) при недоступном сервере может отвечать десятки
    секунд — проверять в UI-потоке нельзя.
    """
    result = Signal(bool, str)  # (доступен, проверенный путь)

    def __init__(self, path: str, parent=None):
        super().__init__(parent)
        self.path = path

    def run(self):
        ok = bool(self.path) and os.path.isfile(self.path)
        self.result.emit(ok, self.path)


class MasterScheduleDialog(QDialog):
    """Окно после загрузки программы: график не подтянулся с сервера.

    Показывается при КАЖДОМ запуске с недоступным сервером — человек всегда
    осознанно подтверждает, с какой копией графика работает. Прошлый ручной
    выбор подставляется как подсказка, рядом видна дата изменения файла.
    """

    def __init__(self, missing_path: str, prefill: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Сводный график")
        self.setModal(True)
        self.setMinimumWidth(540)

        v = QVBoxLayout(self)
        v.setContentsMargins(22, 20, 22, 18)
        v.setSpacing(10)

        lbl_head = QLabel("График не подтянулся — сервер недоступен")
        lbl_head.setObjectName("form_field_label")
        v.addWidget(lbl_head)

        lbl_text = QLabel(
            f"Не удалось открыть сводный график:\n"
            f"{missing_path or '(путь не задан)'}\n\n"
            "Выберите файл графика вручную — он будет использоваться во всех "
            "разделах программы до её закрытия. Проверьте по дате изменения, "
            "что это актуальная копия графика.")
        lbl_text.setObjectName("hint")
        lbl_text.setWordWrap(True)
        v.addWidget(lbl_text)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.edit_path = QLineEdit()
        self.edit_path.setPlaceholderText("Файл графика (.xlsm / .xlsx)")
        self.edit_path.textChanged.connect(self._validate)
        row.addWidget(self.edit_path, 1)
        btn_browse = QPushButton("📂")
        btn_browse.setObjectName("browse_btn")
        btn_browse.setFixedWidth(40)
        btn_browse.setCursor(Qt.PointingHandCursor)
        btn_browse.clicked.connect(self._browse)
        row.addWidget(btn_browse)
        v.addLayout(row)

        self.lbl_file_date = QLabel("")
        self.lbl_file_date.setObjectName("hint")
        v.addWidget(self.lbl_file_date)

        btns = QHBoxLayout()
        btns.addStretch()
        btn_later = QPushButton("Позже")
        btn_later.setObjectName("segment")
        btn_later.setCursor(Qt.PointingHandCursor)
        btn_later.clicked.connect(self.reject)
        btns.addWidget(btn_later)
        self.btn_ok = QPushButton("Использовать этот график")
        self.btn_ok.setObjectName("green_button")
        self.btn_ok.setCursor(Qt.PointingHandCursor)
        self.btn_ok.setEnabled(False)
        self.btn_ok.clicked.connect(self.accept)
        btns.addWidget(self.btn_ok)
        v.addLayout(btns)

        if prefill:
            self.edit_path.setText(prefill)  # textChanged сам вызовет _validate

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Сводный график", "", "Excel (*.xlsm *.xlsx)")
        if path:
            self.edit_path.setText(os.path.normpath(path))

    def _validate(self):
        path = self.edit_path.text().strip()
        ok = os.path.isfile(path)
        self.btn_ok.setEnabled(ok)
        if ok:
            mtime = time.strftime("%d.%m.%Y %H:%M",
                                  time.localtime(os.path.getmtime(path)))
            self.lbl_file_date.setText(f"Файл изменён: {mtime}")
        elif path:
            self.lbl_file_date.setText("Файл не найден")
        else:
            self.lbl_file_date.setText("")

    def selected_path(self) -> str:
        return self.edit_path.text().strip()


class MainWindow(QMainWindow):
    """Главное окно приложения ПНОС."""

    # Индексы страниц в QStackedWidget
    PAGE_PROCESS = 0
    PAGE_MANUAL = 1
    PAGE_OVERVIEW = 2   # объединённая страница «Обзор и отчёт»
    PAGE_FIXREPORTS = 3  # автономная «Правка отчётов» (вне конвейера)
    PAGE_ANALYTICS = 4   # автономный «Анализ пунктов» (вне конвейера)
    PAGE_SETTINGS = 5

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

        # Проверка доступности сводного графика — после показа окна.
        # Если серверный путь недоступен, программа сразу предложит выбрать
        # файл вручную; выбор запоминается и используется во всех разделах.
        QTimer.singleShot(800, self._check_master_schedule)

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
            ("Правка отчётов", self.PAGE_FIXREPORTS),
            ("Анализ пунктов", self.PAGE_ANALYTICS),
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

        self.page_fixreports = QWidget()
        self._setup_fixreports_page()
        self.pages_stack.addWidget(self.page_fixreports)

        self.page_analytics = QWidget()
        self._setup_analytics_page()
        self.pages_stack.addWidget(self.page_analytics)

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
        elif index == self.PAGE_ANALYTICS:
            self._an_auto_load()

    def _page_title(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("app_title")
        return lbl

    # ═══════════════════════════════════════════════════════════════
    # Сводный график — единый файл для шага 4 и «Анализа пунктов»
    # ═══════════════════════════════════════════════════════════════

    def _master_excel_path(self) -> str:
        """Путь к сводному графику: из Настроек (запоминается), иначе config.py."""
        override = self.page_settings.edit_macro_path.text().strip()
        return override or CONFIG.get("MACRO_MASTER_WB_PATH", "")

    def _check_master_schedule(self):
        worker = MasterCheckWorker(self._master_excel_path(), parent=self)
        worker.result.connect(self._on_master_check)
        worker.start()
        self._workers["master_check"] = worker

    def _on_master_check(self, ok: bool, path: str):
        self._workers.pop("master_check", None)
        if ok:
            # график доступен — сразу подтягиваем его в фоне;
            # по завершении чтения появится тост (см. _an_on_rows)
            self._an_load_schedule()
            return
        dlg = MasterScheduleDialog(
            path, prefill=self._settings.value("master_excel", ""), parent=self)
        if dlg.exec() == QDialog.Accepted:
            # поле в Настройках действует до закрытия программы;
            # textChanged заодно сохранит путь как подсказку на будущее
            self.page_settings.edit_macro_path.setText(dlg.selected_path())
            self._an_load_schedule()

    def _show_toast(self, message: str, duration_ms: int = 4000):
        """Всплывающее уведомление поверх окна; исчезает само."""
        if getattr(self, "_toast", None) is None:
            self._toast = QLabel("", self)
            self._toast.setObjectName("toast")
            self._toast.setAlignment(Qt.AlignCenter)
            self._toast.hide()
            self._toast_timer = QTimer(self)
            self._toast_timer.setSingleShot(True)
            self._toast_timer.timeout.connect(self._toast.hide)
        self._toast.setText(message)
        self._toast.adjustSize()
        self._toast.move(max(0, (self.width() - self._toast.width()) // 2), 16)
        self._toast.show()
        self._toast.raise_()
        self._toast_timer.start(duration_ms)

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
            # QFileDialog отдаёт прямые слэши — Word COM такие пути не открывает
            self.entry_local.setText(os.path.normpath(folder))

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
                # Конвейер завершён — показываем вкладку «Отчёт о прогоне»
                self._session_timer.stop()
                self._switch_page(self.PAGE_OVERVIEW)
                self._set_overview_tab(1)
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
            # Кнопки страницы «Правка отчётов»
            if hasattr(self, "fix_btn_start"):
                self.fix_btn_start.setEnabled(True)
                self.fix_btn_stop.setEnabled(False)

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
        """Поле даты с жёсткой маской дд.мм.гггг (без календаря)."""
        de = QLineEdit()
        de.setInputMask("00.00.0000;_")
        de.setCursorPosition(0)
        return de

    def _date_edit_value(self, de) -> str:
        """Строка дд.мм.гггг или '' если поле пустое (одни подчёркивания)."""
        s = de.text().strip()
        if not any(c.isdigit() for c in s):
            return ""
        return s

    def _set_date_edit(self, de, value: str):
        de.setText((value or "").strip())
        de.setCursorPosition(0)

    @staticmethod
    def _date_ok_or_empty(s: str) -> bool:
        """Пустая строка — ок; иначе должна быть существующая дата дд.мм.гггг."""
        s = (s or "").strip()
        if not s:
            return True
        from PySide6.QtCore import QDate
        d = QDate.fromString(s, "dd.MM.yyyy")
        return d.isValid() and 1900 <= d.year() <= 2100

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
        from widgets import NoWheelComboBox as QComboBox
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

        import user_data
        f["specialist"] = QComboBox()
        f["specialist"].addItem("— выберите —", "")
        for name in user_data.get_list("specialists"):
            f["specialist"].addItem(name, name)
        form.addWidget(self._form_field("Специалист (B22)", f["specialist"], required=True))

        f["date"] = self._make_date_edit()
        form.addWidget(self._form_field("Дата (B33)", f["date"], required=True,
                                        note="Введите дату цифрами в формате дд.мм.гггг."))

        sep = QLabel("Данные предыдущей ЭПБ")
        sep.setObjectName("section_label")
        form.addWidget(sep)

        f["org"] = QComboBox()
        f["org"].addItem("— не выбрано —", "")
        for org in user_data.get_list("epb_orgs"):
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

        f["epb_date"] = self._make_date_edit()
        form.addWidget(self._form_field(
            "Дата предыдущей ЭПБ (B54)", f["epb_date"], required=False,
            note="Формат дд.мм.гггг. Если даты нет — оставьте поле пустым."))

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

    def _refresh_form_choices(self):
        """Пересобирает выпадающие списки специалистов/организаций из user_data
        (могли измениться в настройках), сохраняя текущий выбор."""
        import user_data
        for key, list_name in (("specialist", "specialists"), ("org", "epb_orgs")):
            combo = self.manual_fields[key]
            current = combo.currentData() or ""
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("— выберите —" if key == "specialist" else "— не выбрано —", "")
            for val in user_data.get_list(list_name):
                combo.addItem(val, val)
            idx = combo.findData(current)
            combo.setCurrentIndex(idx if idx >= 0 else 0)
            combo.blockSignals(False)

    def _load_manual_points(self):
        self._refresh_form_choices()
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
        # B33 обязательна и должна быть существующей датой;
        # B54 необязательна, но если введена — тоже должна быть корректной.
        if not (purpose and specialist and date):
            return False
        if not self._date_ok_or_empty(date):
            return False
        if not self._date_ok_or_empty(values.get("epb_date", "")):
            return False
        return True

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
                "или даты введены некорректно у пунктов:\n\n" + ", ".join(not_ready) +
                "\n\nПроверьте поля или нажмите «Пропустить пункт».")
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

    def _setup_overview_page(self):
        """Страница «Обзор и отчёт» с вкладками: Обзор | Отчёт о прогоне."""
        outer = QVBoxLayout(self.page_overview)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(12)

        # ── Шапка: заголовок + переключатель вкладок ──
        head = QHBoxLayout()
        head.addWidget(self._page_title("Обзор и отчёт"))
        head.addStretch()
        self.btn_tab_overview = QPushButton("Обзор")
        self.btn_tab_report = QPushButton("Отчёт о прогоне")
        for b in (self.btn_tab_overview, self.btn_tab_report):
            b.setObjectName("segment")
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
        self.btn_tab_overview.setChecked(True)
        self.btn_tab_overview.clicked.connect(lambda: self._set_overview_tab(0))
        self.btn_tab_report.clicked.connect(lambda: self._set_overview_tab(1))
        head.addWidget(self.btn_tab_overview)
        head.addWidget(self.btn_tab_report)
        outer.addLayout(head)

        self.overview_tabs = QStackedWidget()
        outer.addWidget(self.overview_tabs, 1)

        # ═══ Вкладка 0: Обзор ═══
        tab_ov = QWidget()
        ov = QVBoxLayout(tab_ov)
        ov.setContentsMargins(0, 0, 0, 0)
        ov.setSpacing(10)

        chips_row = QHBoxLayout()
        chips_row.setSpacing(8)
        chip_total, self.chip_total = self._build_chip("ПУНКТОВ")
        chip_done, self.chip_done = self._build_chip("ГОТОВЫ", "chip_value_good")
        chip_bad, self.chip_bad = self._build_chip("НЕ КОРРЕКТНЫХ", "chip_value_bad")
        chip_skip, self.chip_skip = self._build_chip("ПРОПУЩЕНО", "chip_value_warn")
        for c in (chip_total, chip_done, chip_bad, chip_skip):
            chips_row.addWidget(c)
        chips_row.addStretch()

        btn_refresh = QPushButton("Обновить")
        btn_refresh.setObjectName("ghost_button")
        btn_refresh.setCursor(Qt.PointingHandCursor)
        btn_refresh.clicked.connect(self._refresh_overview)
        btn_reset = QPushButton("Сбросить статус…")
        btn_reset.setObjectName("ghost_button")
        btn_reset.setCursor(Qt.PointingHandCursor)
        btn_reset.clicked.connect(self._reset_point_dialog)
        chips_row.addWidget(btn_refresh, 0, Qt.AlignBottom)
        chips_row.addWidget(btn_reset, 0, Qt.AlignBottom)
        ov.addLayout(chips_row)

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
        ov.addWidget(self.overview_table, 1)

        self.overview_tabs.addWidget(tab_ov)

        # ═══ Вкладка 1: Отчёт о прогоне ═══
        tab_rep = QWidget()
        rp = QVBoxLayout(tab_rep)
        rp.setContentsMargins(0, 0, 0, 0)
        rp.setSpacing(10)

        rep_head = QHBoxLayout()
        self.lbl_report_summary = QLabel("Отчёт появится после выполнения этапов.")
        self.lbl_report_summary.setObjectName("card_description")
        rep_head.addWidget(self.lbl_report_summary)
        rep_head.addStretch()
        btn_open = QPushButton("📂 Папка отчётов")
        btn_open.setObjectName("ghost_button")
        btn_open.setCursor(Qt.PointingHandCursor)
        btn_open.clicked.connect(self._open_reports_folder)
        rep_head.addWidget(btn_open)
        rp.addLayout(rep_head)

        rep_scroll = QScrollArea()
        rep_scroll.setWidgetResizable(True)
        rep_scroll.setFrameShape(QFrame.NoFrame)
        rep_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        rep_body = QWidget()
        rep_v = QVBoxLayout(rep_body)
        rep_v.setContentsMargins(0, 0, 6, 0)
        rep_v.setSpacing(10)
        self._report_sections = QVBoxLayout()
        self._report_sections.setSpacing(10)
        rep_v.addLayout(self._report_sections)
        rep_v.addStretch()
        rep_scroll.setWidget(rep_body)
        rp.addWidget(rep_scroll, 1)

        self.overview_tabs.addWidget(tab_rep)

    def _set_overview_tab(self, idx: int):
        """Переключает вкладку страницы «Обзор и отчёт»."""
        self.overview_tabs.setCurrentIndex(idx)
        self.btn_tab_overview.setChecked(idx == 0)
        self.btn_tab_report.setChecked(idx == 1)

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

        c_done = QColor(STATUS_COLORS["done"])
        c_wait = QColor(STATUS_COLORS["idle"])
        c_skip = QColor(STATUS_COLORS["wait"])
        c_bad = QColor(STATUS_COLORS["error"])
        bg_bad = QColor(239, 68, 68, 26)     # полупрозрачная подсветка строки
        bg_skip = QColor(245, 158, 11, 20)

        done_all = invalid_n = skipped_n = 0
        for r, pt in enumerate(points):
            invalid, skipped = pt["invalid"], pt["skipped"]

            # Имя пункта с причиной прямо в строке
            name = pt["name"]
            if invalid:
                reason = ", ".join(pt["missing"]) if pt["missing"] else "нет обязательных папок"
                base = name.split(" (")[0]  # без суффикса «(не корректный)»
                name_item = QTableWidgetItem(f"{base} — {reason}")
                name_item.setForeground(c_bad)
                invalid_n += 1
            elif skipped:
                name_item = QTableWidgetItem(f"{name} — пропущен")
                name_item.setForeground(QColor("#b45309"))
                skipped_n += 1
            else:
                name_item = QTableWidgetItem(name)
            self.overview_table.setItem(r, 0, name_item)

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

            # Подсветка всей строки для проблемных/пропущенных
            if invalid or skipped:
                bg = bg_bad if invalid else bg_skip
                for c in range(self.overview_table.columnCount()):
                    it = self.overview_table.item(r, c)
                    if it:
                        it.setBackground(bg)

            if not invalid and not skipped and all_done:
                done_all += 1

        # Чипы-сводка
        self.chip_total.setText(str(len(points)))
        self.chip_done.setText(str(done_all))
        self.chip_bad.setText(str(invalid_n))
        self.chip_skip.setText(str(skipped_n))

    def _reset_point_dialog(self):
        """Диалог: выбрать пункт и что сбросить (весь пункт / конкретный шаг)."""
        from PySide6.QtWidgets import QDialog, QDialogButtonBox
        from widgets import NoWheelComboBox as QComboBox
        points = [pt["name"] for pt in self._scan_points()]
        if not points:
            QMessageBox.information(self, "Нет пунктов", "В локальной папке нет пунктов.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Сброс статуса")
        v = QVBoxLayout(dlg)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)

        v.addWidget(QLabel("Пункт:"))
        cb_point = QComboBox()
        cb_point.addItem("— все пункты —")
        cb_point.addItems(points)
        v.addWidget(cb_point)

        v.addWidget(QLabel("Что сбросить:"))
        cb_scope = QComboBox()
        cb_scope.addItem("Весь пункт (обработать заново с шага 1)", "all")
        for step in range(1, 8):
            cb_scope.addItem(f"Только шаг {step}: {self.STEP_TITLES[step]}",
                             self.STEP_STATUS_KEYS[step])
        v.addWidget(cb_scope)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Сбросить")
        buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        v.addWidget(buttons)

        if dlg.exec() != QDialog.Accepted:
            return

        from status_manager import reset_point_status, reset_stage_status
        scope = cb_scope.currentData()
        root = self._get_local_root()

        # Список целей: конкретный пункт или все
        if cb_point.currentIndex() == 0:
            if QMessageBox.question(
                    self, "Сброс всех пунктов",
                    f"Сбросить {'все статусы' if scope == 'all' else 'этот шаг'} "
                    f"у всех пунктов ({len(points)})?\n\nСкачанные файлы не удаляются.",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
                return
            targets = points
        else:
            targets = [cb_point.currentText()]

        ok_n = 0
        for name in targets:
            folder = os.path.join(root, name)
            if scope == "all":
                ok_n += 1 if reset_point_status(folder) else 0
            else:
                ok_n += 1 if reset_stage_status(folder, scope) else 0
        what = "статус полностью" if scope == "all" else f"шаг «{scope}»"
        self._add_activity(0, f"Сброшен {what}: {ok_n} из {len(targets)} пунктов", "done")
        self._refresh_overview()


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
            box.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
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

    # ═══════════════════════════════════════════════════════════════
    # Страница «Анализ пунктов» — статистика готовности (вне конвейера)
    # ═══════════════════════════════════════════════════════════════

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
        sv.addWidget(self.an_list, 1)

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
        self.an_btn_start.setObjectName("green_button")
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
        self.an_btn_export.setObjectName("segment")
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
        rv.addWidget(self.an_table, 1)
        self.an_tabs.addWidget(tab_res)

        # ── Низ: статус + остановка + прогресс ──
        bottom = QHBoxLayout()
        bottom.setSpacing(10)
        self.an_lbl_status = QLabel(
            "График прочитается сам при открытии страницы.")
        self.an_lbl_status.setObjectName("hint")
        bottom.addWidget(self.an_lbl_status, 1)
        self.an_btn_stop = QPushButton("Остановить")
        self.an_btn_stop.setObjectName("segment")
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
            self._an_set_status(
                "Не найден мастер-файл сводного графика — укажите его "
                "на странице «Настройки»", "error")
            return
        self.an_lbl_file.setText(
            f"График: {os.path.basename(path)} (мастер-файл из Настроек)")
        self.an_btn_reload.setEnabled(False)
        self._an_loaded_path = path

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

    def _an_on_rows(self, rows: list):
        self._an_rows = rows
        self._an_checked.clear()
        self._an_fill_equipment()
        self._an_refresh_list()
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
        self._an_set_tab(1)

    def _an_on_scan_finished(self, success: bool):
        self.an_btn_reload.setEnabled(True)
        self.an_btn_stop.setEnabled(False)
        self._workers.pop("analytics", None)
        self._an_update_counter()  # вернёт кнопку запуска в актуальное состояние

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
