# -*- coding: utf-8 -*-
"""
Главное окно приложения ПНОС — PySide6.

Здесь каркас окна (узкое боковое меню NavRail, стек страниц, тосты),
сводный график и общие для всех страниц методы. Сами страницы — классы-примеси
в pages/*_page.py: «Пункты» (рабочее место: запуск шагов, «До конца»),
«Отчёт о прогоне» (вкладка рабочего места), «Ручной контроль» (шаг 3),
автономные «Правка отчётов» и «Анализ пунктов»; «Настройки» — отдельный
виджет pages/settings_page.py.
"""

import os
import time

from PySide6.QtCore import Qt, QSettings, QTimer, QThread, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFrame, QLineEdit, QFileDialog, QDialog, QApplication, QStackedWidget,
)

from config import CONFIG
from utils import resource_path
from version import APP_VERSION
from pages.settings_page import SettingsPage
from widgets import NavRail
# Страницы — примеси к MainWindow (см. pages/*_page.py)
from pages.process_page import ProcessPageMixin
from pages.manual_page import ManualPageMixin
from pages.report_page import RunReportMixin
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
        btn_later.setObjectName("ghost_button")
        btn_later.setCursor(Qt.PointingHandCursor)
        btn_later.clicked.connect(self.reject)
        btns.addWidget(btn_later)
        self.btn_ok = QPushButton("Использовать этот график")
        self.btn_ok.setObjectName("card_button")
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


class MainWindow(ProcessPageMixin, ManualPageMixin, RunReportMixin,
                 FixReportsPageMixin, AnalyticsPageMixin, QMainWindow):
    """Главное окно приложения ПНОС."""

    # Индексы страниц в QStackedWidget
    PAGE_WORKSPACE = 0   # «Пункты»: таблица пунктов × шагов и «Отчёт о прогоне»
    PAGE_MANUAL = 1
    PAGE_FIXREPORTS = 2  # автономная «Правка отчётов» (вне конвейера)
    PAGE_ANALYTICS = 3   # автономный «Анализ пунктов» (вне конвейера)
    PAGE_SETTINGS = 4
    # Отдельной страницы «Ремонт» нет: модуль живёт внутри «Ручного контроля»
    # (кнопка «Внести данные по ремонту» у пункта с ремонтом).

    # Боковое меню — единственное место, где перечислены его пункты:
    # (ключ, подпись, значок resources/icons/<значок>.svg, страница, внизу ли).
    NAV_ITEMS = [
        ("points", "Пункты", "points", PAGE_WORKSPACE, False),
        ("form", "Форма", "form", PAGE_MANUAL, False),
        ("fix", "Правка", "fix", PAGE_FIXREPORTS, False),
        ("analytics", "Анализ", "analytics", PAGE_ANALYTICS, False),
        ("settings", "Настройки", "settings", PAGE_SETTINGS, True),
    ]
    # Ширина страниц-форм на широком мониторе (таблицы занимают всю ширину)
    READABLE_WIDTH = 1000

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

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ПНОС — Пакетный генератор отчётов")
        self.setMinimumSize(1000, 680)
        self.resize(1280, 820)

        self._workers = {}
        self._stage_results = {}     # {ui_step: {"columns": [...], "details": [...]} | dict}
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

        # ─── Боковое меню ──────────────────────────────────────────
        self.nav_rail = NavRail(
            resource_path(os.path.join("resources", "icon_mark.png")),
            f"ПНОС, версия {APP_VERSION}")
        self._nav_keys = {}   # индекс страницы -> ключ пункта меню
        nav_pages = {}
        for key, caption, icon_name, page_idx, bottom in self.NAV_ITEMS:
            if bottom:
                self.nav_rail.add_bottom_item(key, caption, icon_name)
            else:
                self.nav_rail.add_item(key, caption, icon_name)
            self._nav_keys[page_idx] = key
            nav_pages[key] = page_idx
        self.nav_rail.item_clicked.connect(
            lambda key: self._switch_page(nav_pages[key]))
        self.nav_rail.recolor(self._palette())
        main_layout.addWidget(self.nav_rail)

        # ─── Контент ───────────────────────────────────────────────
        content_frame = QFrame()
        content_frame.setObjectName("content_frame")
        content_layout = QVBoxLayout(content_frame)
        content_layout.setContentsMargins(24, 14, 24, 20)
        content_layout.setSpacing(10)

        self.pages_stack = QStackedWidget()
        content_layout.addWidget(self.pages_stack, 1)

        # Настройки создаются первыми: из них остальные страницы берут папку
        # пунктов (_get_local_root), а в стек они встают последними.
        self.page_settings = SettingsPage()
        self.page_settings.theme_changed.connect(self.apply_theme)
        self.page_settings.local_root_changed.connect(self._on_local_root_changed)

        # Страницы (порядок = индексы PAGE_*). Страницы-формы держим читаемой
        # ширины по центру; страницы с таблицами занимают всю ширину.
        self.page_process = QWidget()
        self._setup_process_page()
        self.pages_stack.addWidget(self.page_process)

        self.page_manual = QWidget()
        self._setup_manual_page()
        self.pages_stack.addWidget(self.page_manual)

        self.page_fixreports = QWidget()
        self._setup_fixreports_page()
        self.pages_stack.addWidget(self._readable(self.page_fixreports))

        self.page_analytics = QWidget()
        self._setup_analytics_page()
        self.pages_stack.addWidget(self.page_analytics)

        self.pages_stack.addWidget(self._readable(self.page_settings))

        main_layout.addWidget(content_frame, 1)
        self._switch_page(self.PAGE_WORKSPACE)

    def _readable(self, page: QWidget) -> QWidget:
        """Обёртка: страница не шире READABLE_WIDTH, по центру."""
        page.setMaximumWidth(self.READABLE_WIDTH)
        holder = QWidget()
        h = QHBoxLayout(holder)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        h.addStretch()
        h.addWidget(page, 1)
        h.addStretch()
        return holder

    def _palette(self) -> dict:
        from theme import get_palette
        return get_palette(self._theme_name)

    def _update_nav_badges(self):
        """
        Счётчик на «Форме»: сколько пунктов в папке ПНОС ждут ручного шага 3 —
        таблицы скопированы (stage2), форма не записана (manual3), пункт не
        «не корректный» и не пропущен.
        """
        from status_manager import get_stage_status, is_point_blocked
        root = self._get_local_root()
        waiting = 0
        try:
            names = os.listdir(root) if os.path.isdir(root) else []
        except OSError:
            names = []
        for name in names:
            low = name.lower()
            if not (low.startswith("п.") or low.startswith("п ")
                    or low.startswith("пункт")):
                continue
            folder = os.path.join(root, name)
            if not os.path.isdir(folder):
                continue
            if (get_stage_status(folder, "stage2")
                    and not get_stage_status(folder, "manual3")
                    and not is_point_blocked(folder)[0]):
                waiting += 1
        self.nav_rail.set_badge("form", waiting)

    def _switch_page(self, index: int):
        self.pages_stack.setCurrentIndex(index)
        self.nav_rail.set_current(self._nav_keys.get(index, ""))
        self._update_nav_badges()
        if index == self.PAGE_WORKSPACE:
            self._refresh_workspace()
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
        self.workspace.add_log(step, message, category)
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
        self.page_settings.browse_local_root()

    def _get_local_root(self) -> str:
        """Папка пунктов — поле в Настройках (запоминается между запусками)."""
        return self.page_settings.local_root()

    def _on_local_root_changed(self, path: str):
        self._add_activity(0, f"Папка пунктов: {path}", "info")
        self._refresh_workspace()
        self._update_nav_badges()

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

    def _tick_session_timer(self):
        self._update_run_status()

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
        self.nav_rail.recolor(self._palette())
