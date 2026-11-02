# -*- coding: utf-8 -*-
"""
Главное окно приложения ПНОС — PySide6.

Здесь каркас окна (боковое меню, стек страниц, тосты), сводный график и
общие для всех страниц методы. Сами страницы — классы-примеси в pages/*_page.py:
«Процесс» (карточки шагов и запуск этапов), «Ручной контроль» (шаг 3),
«Обзор и отчёт», автономные «Правка отчётов» и «Анализ пунктов»;
«Настройки» — отдельный виджет pages/settings_page.py.
"""

import os
import time

from PySide6.QtCore import Qt, QSettings, QTimer, QThread, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QProgressBar, QFrame, QLineEdit, QFileDialog, QDialog, QApplication,
    QStackedWidget, QListWidgetItem,
)

from config import CONFIG
from utils import resource_path
from version import APP_VERSION
from pages.settings_page import SettingsPage
from widgets import dot_icon, LOG_COLORS
# Страницы — примеси к MainWindow (см. pages/*_page.py)
from pages.process_page import ProcessPageMixin
from pages.manual_page import ManualPageMixin
from pages.overview_page import OverviewPageMixin
from pages.fixreports_page import FixReportsPageMixin
from pages.analytics_page import AnalyticsPageMixin
# Воркеры этапов импортируются лениво (внутри методов запуска),
# чтобы не тянуть pandas / aiohttp / win32com / python-docx при старте.


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


class MainWindow(ProcessPageMixin, ManualPageMixin, OverviewPageMixin,
                 FixReportsPageMixin, AnalyticsPageMixin, QMainWindow):
    """Главное окно приложения ПНОС."""

    # Индексы страниц в QStackedWidget
    PAGE_PROCESS = 0
    PAGE_MANUAL = 1
    PAGE_OVERVIEW = 2   # объединённая страница «Обзор и отчёт»
    PAGE_FIXREPORTS = 3  # автономная «Правка отчётов» (вне конвейера)
    PAGE_ANALYTICS = 4   # автономный «Анализ пунктов» (вне конвейера)
    PAGE_SETTINGS = 5
    # Отдельной страницы «Ремонт» нет: модуль живёт внутри «Ручного контроля»
    # (кнопка «Внести данные по ремонту» у пункта с ремонтом).

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
