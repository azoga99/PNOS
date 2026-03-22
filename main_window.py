# -*- coding: utf-8 -*-
"""
Главное окно приложения ПНОС — PySide6.
5 карточек-этапов, панель настроек, лог.
"""

import os
import sys
import time

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont, QIcon, QPainter, QColor
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QProgressBar, QFrame, QLineEdit,
    QFileDialog, QTextEdit, QMessageBox, QDialog, QCheckBox,
    QSizePolicy, QSpacerItem, QApplication, QStackedWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QListWidget,
    QListWidgetItem
)

from config import CONFIG
from workers.stage1_worker import Stage1Worker
from workers.stage2_worker import Stage2Worker
from workers.stage3_worker import Stage3Worker
from workers.stage4_worker import Stage4Worker
from workers.stage5_worker import Stage5Worker

import subprocess


# ═══════════════════════════════════════════════════════════════
# Главное окно
# ═══════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    """Главное окно приложения ПНОС."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ПНОС — Пакетный генератор отчётов")
        self.setMinimumSize(1100, 750)
        self.resize(1200, 800)

        self._workers = {}
        self._stage_results = {} # Храним итоги для общего отчета
        self._start_time = None
        self._is_stopping = False
        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        # ─── Основной двухколоночный Layout (Sidebar + Content) ────
        main_layout = QHBoxLayout(central)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # ─── САЙДБАР (Слева) ───────────────────────────────────────
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(260)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)

        # Логотип/Заголовок в сайдбаре
        logo_frame = QFrame()
        logo_frame.setObjectName("logo_frame")
        logo_layout = QVBoxLayout(logo_frame)
        logo_layout.setContentsMargins(20, 30, 20, 30)
        
        app_title = QLabel("ПНОС")
        app_title.setObjectName("sidebar_title")
        logo_layout.addWidget(app_title)

        app_sub = QLabel("Генератор отчётов")
        app_sub.setObjectName("sidebar_subtitle")
        logo_layout.addWidget(app_sub)
        
        sidebar_layout.addWidget(logo_frame)

        # Навигационные кнопки (Шаги)
        self.nav_buttons = []
        stages = [
            (1, "Создание структуры"),
            (2, "Копирование таблиц"),
            (3, "Ручной контроль"),
            (4, "Запуск макросов"),
            (5, "Создание отчёта"),
            (6, "Вставка картинок"),
        ]

        # Контейнер для кнопок, чтобы они прижимались наверх
        nav_container = QWidget()
        nav_layout = QVBoxLayout(nav_container)
        nav_layout.setContentsMargins(0, 10, 0, 0)
        nav_layout.setSpacing(5)

        for num, title in stages:
            btn = QPushButton(f"Шаг {num}: {title}")
            btn.setObjectName("nav_button")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            
            # Сохраняем индекс страницы для переключения
            idx = num - 1
            btn.clicked.connect(lambda checked, i=idx: self._switch_page(i))
            
            self.nav_buttons.append(btn)
            nav_layout.addWidget(btn)

        sidebar_layout.addWidget(nav_container)
        sidebar_layout.addStretch()
        
        # ─── Кнопка настроек внизу ─────────────────────────────────
        self.btn_settings = QPushButton("⚙️ Настройки")
        self.btn_settings.setObjectName("nav_button")
        self.btn_settings.setCheckable(True)
        self.btn_settings.setCursor(Qt.PointingHandCursor)
        self.btn_settings.clicked.connect(lambda: self._switch_page(7))
        sidebar_layout.addWidget(self.btn_settings)
        sidebar_layout.addSpacing(10)

        # ─── Глобальный прогресс ───────────────────────────────────
        progress_container = QWidget()
        progress_layout = QVBoxLayout(progress_container)
        progress_layout.setContentsMargins(20, 10, 20, 30)
        
        lbl_global = QLabel("Общий прогресс:")
        lbl_global.setStyleSheet("color: #94a3b8; font-weight: bold; font-size: 11px;")
        progress_layout.addWidget(lbl_global)
        
        self.global_progress = QProgressBar()
        self.global_progress.setValue(0)
        self.global_progress.setMaximum(100)
        self.global_progress.setFixedHeight(12)
        self.global_progress.setTextVisible(False)
        self.global_progress.setStyleSheet("""
            QProgressBar { border: none; background-color: #334155; border-radius: 6px; }
            QProgressBar::chunk { background-color: #2dd4bf; border-radius: 6px; }
        """)
        progress_layout.addWidget(self.global_progress)
        
        sidebar_layout.addWidget(progress_container)

        main_layout.addWidget(sidebar)

        # ─── КОНТЕНТ МЕНЮ (Справа) ─────────────────────────────────
        content_frame = QFrame()
        content_frame.setObjectName("content_frame")
        content_layout = QVBoxLayout(content_frame)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(15)

        # StackedWidget для страниц (Этапов) (центрированный с макс-шириной)
        stack_wrapper = QHBoxLayout()
        stack_wrapper.addStretch()
        
        self.pages_stack = QStackedWidget()
        self.pages_stack.setMaximumWidth(850)
        stack_wrapper.addWidget(self.pages_stack, 1)
        
        stack_wrapper.addStretch()
        
        content_layout.addLayout(stack_wrapper)

        # ─── СТРАНИЦА 1: Настройки и Запуск ───────────────────────
        self.page1 = QWidget()
        self._setup_page1()
        self.pages_stack.addWidget(self.page1)

        # ─── СТРАНИЦЫ 2-5: Заглушки ───────────────────────────────
        self.stub_pages = []
        
        stage_meta = {
            2: {"title": "Шаг 2: Копирование таблиц", "desc": "Извлечение 3-й таблицы из 'Акта замеров' и перенос её на листы 'Импорт' без форматирования.", "warn": ""},
            3: {"title": "Шаг 3: Ручной контроль", "desc": "Проверьте файлы перед запуском макросов. Этот этап выполняется вручную.", "warn": "Инструкции будут добавлены позже."},
            4: {"title": "Шаг 4: Запуск макросов", "desc": "Перенос переменных из Паспорта (Word) в Excel и автоматический запуск макросов Restore и Save.", "warn": "⚠️ Во время работы этого этапа не открывайте другие файлы Excel и не перехватывайте фокус мыши!"},
            5: {"title": "Шаг 5: Создание отчёта", "desc": "Генерация финального отчета Word.", "warn": ""},
            6: {"title": "Шаг 6: Вставка картинок", "desc": "Пост-обработка: вставка скачанных фотографий в отчёт (в разработке).", "warn": ""}
        }
        
        for stage_num in range(2, 7):
            page = QWidget()
            layout = QVBoxLayout(page)
            layout.setContentsMargins(0, 0, 0, 0)
            
            meta = stage_meta.get(stage_num, {})
            lbl = QLabel(meta.get("title", f"Настройки для Шага {stage_num}"))
            lbl.setObjectName("app_title")
            layout.addWidget(lbl)
            
            desc = QLabel(meta.get("desc", ""))
            desc.setObjectName("card_description")
            desc.setWordWrap(True)
            layout.addWidget(desc)
            
            if meta.get("warn"):
                warn = QLabel(meta.get("warn"))
                warn.setStyleSheet("color: #fbbf24; font-weight: bold; margin-top: 10px;")
                warn.setWordWrap(True)
                layout.addWidget(warn)
            
            layout.addStretch()

            btn_layout = QHBoxLayout()
            btn_run = QPushButton(f"Запустить {meta.get('title', f'Шаг {stage_num}').split(':')[0]}")
            btn_run.setObjectName("card_button")
            btn_run.setFixedWidth(200)
            btn_run.setCursor(Qt.PointingHandCursor)
            
            btn_stop = QPushButton("⏹ Остановить")
            btn_stop.setObjectName("stop_button")
            btn_stop.setFixedWidth(120)
            btn_stop.setCursor(Qt.PointingHandCursor)
            btn_stop.setEnabled(False)
            btn_stop.clicked.connect(self._force_stop_active_stage)

            btn_skip = QPushButton("Пропустить")
            btn_skip.setObjectName("link_button")
            btn_skip.setCursor(Qt.PointingHandCursor)

            # Привязываем к заглушкам
            if stage_num == 2: btn_run.clicked.connect(self._start_stage2)
            if stage_num == 3: btn_run.clicked.connect(self._start_manual_stage3)
            if stage_num == 4: btn_run.clicked.connect(self._start_stage3)
            if stage_num == 5: btn_run.clicked.connect(self._start_stage4)
            if stage_num == 6: btn_run.clicked.connect(self._start_stage5)
            
            btn_skip.clicked.connect(lambda checked, idx=stage_num: self._skip_stage(idx))
            
            btn_layout.addStretch()
            btn_layout.addWidget(btn_run)
            btn_layout.addWidget(btn_stop)
            btn_layout.addWidget(btn_skip)
            btn_layout.addStretch()
            layout.addLayout(btn_layout)

            # Лента активности вместо одного лейбла
            act_list = QListWidget()
            act_list.setObjectName("activity_feed")
            act_list.setFixedHeight(100)
            layout.addWidget(act_list)

            # Таблица предпросмотра (только для 2 и 4 этапов)
            preview_table = None
            if stage_num in (2, 4):
                preview_table = QTableWidget()
                preview_table.setColumnCount(5) # примерно 5 колонок для превью
                preview_table.horizontalHeader().setVisible(False)
                preview_table.verticalHeader().setVisible(False)
                preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
                preview_table.setSelectionMode(QTableWidget.NoSelection)
                preview_table.setFixedHeight(120)
                preview_table.setStyleSheet("""
                    QTableWidget {
                        background-color: #1e293b;
                        gridline-color: #334155;
                        color: #94a3b8;
                        font-size: 10px;
                        border: 1px solid #334155;
                        border-radius: 4px;
                    }
                """)
                preview_table.hide() # скрыта по умолчанию
                layout.addWidget(preview_table)

            # Прогресс бар
            progress = QProgressBar()
            progress.setValue(0)
            progress.setFixedHeight(10)
            progress.setTextVisible(False)
            layout.addWidget(progress)

            # Интегрированная таблица отчета
            report_table = QTableWidget()
            report_table.setEditTriggers(QTableWidget.NoEditTriggers)
            report_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            report_table.setAlternatingRowColors(True)
            report_table.setFixedHeight(250)
            report_table.hide()
            layout.addWidget(report_table)
            
            self.stub_pages.append({
                "page": page,
                "btn_run": btn_run,
                "btn_stop": btn_stop,
                "progress": progress,
                "activity_list": act_list,
                "btn_skip": btn_skip,
                "preview_table": preview_table,
                "report_table": report_table
            })
            self.pages_stack.addWidget(page)

        # ─── СТРАНИЦА: ГЛОБАЛЬНЫЙ ОТЧЕТ ───────────────────────────
        self.page_summary = QWidget()
        self._setup_summary_page()
        self.pages_stack.addWidget(self.page_summary)

        # ─── СТРАНИЦА: НАСТРОЙКИ (Кнопка внизу) ───────────────────
        self.page_settings = QWidget()
        self._setup_settings_page()
        self.pages_stack.addWidget(self.page_settings)

        main_layout.addWidget(content_frame)

        # Выбираем первую страницу по умолчанию
        self._switch_page(0)
        # Отключаем кнопки 2-5
        for i in range(1, 5):
            self.nav_buttons[i].setEnabled(False)

    def _switch_page(self, index: int):
        """Переключает активную страницу и подсвечивает кнопку в сайдбаре."""
        self.pages_stack.setCurrentIndex(index)
        for i, btn in enumerate(self.nav_buttons):
            btn.setChecked(i == index)
        self.btn_settings.setChecked(index == 7)

    def _browse_macro_path(self):
        """Выбор файла макроса в настройках."""
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл макроса", "", "Excel (*.xlsm)")
        if path:
            self.edit_macro_path.setText(os.path.normpath(path))

    def _setup_page1(self):
        """Создает UI для первого этапа (Настройки и парсинг)."""
        main_layout = QVBoxLayout(self.page1)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(20)
        
        title = QLabel("Шаг 1: Создание структуры папок")
        title.setObjectName("app_title")
        main_layout.addWidget(title)
        
        desc = QLabel("На этом этапе программа скачает обязательные папки с Яндекс Диска для каждого пункта и скопирует шаблоны отчета.")
        desc.setObjectName("card_description")
        desc.setWordWrap(True)
        main_layout.addWidget(desc)

        # Панель настроек
        settings = QFrame()
        settings.setObjectName("settings_panel")
        s_layout = QGridLayout(settings)
        s_layout.setSpacing(10)

        # ─── Переключатель режимов ─────────────────────────────────
        mode_layout = QHBoxLayout()
        self.btn_mode_excel = QPushButton("Из Excel")
        self.btn_mode_excel.setCursor(Qt.PointingHandCursor)
        self.btn_mode_excel.setCheckable(True)
        self.btn_mode_excel.setChecked(True)
        self.btn_mode_excel.setProperty("active", True)
        self.btn_mode_excel.setFixedWidth(120)

        self.btn_mode_manual = QPushButton("Вручную")
        self.btn_mode_manual.setCursor(Qt.PointingHandCursor)
        self.btn_mode_manual.setCheckable(True)
        self.btn_mode_manual.setProperty("active", False)
        self.btn_mode_manual.setFixedWidth(120)

        mode_layout.addWidget(self.btn_mode_excel)
        mode_layout.addWidget(self.btn_mode_manual)
        mode_layout.addStretch()
        
        # Заворачиваем в QWidget для Grid
        mode_widget = QWidget()
        mode_widget.setLayout(mode_layout)
        mode_layout.setContentsMargins(0, 0, 0, 0)
        s_layout.addWidget(mode_widget, 0, 0, 1, 3)

        self.btn_mode_excel.clicked.connect(lambda: self._set_mode("excel"))
        self.btn_mode_manual.clicked.connect(lambda: self._set_mode("manual"))
        self._mode = "excel"

        # ─── Stack для смены полей ввода ───────────────────────────
        from PySide6.QtWidgets import QStackedWidget
        self.input_stack = QStackedWidget()
        s_layout.addWidget(self.input_stack, 1, 0, 1, 3)

        # -- Страница 0: Excel --
        page_excel = QWidget()
        layout_excel = QHBoxLayout(page_excel)
        layout_excel.setContentsMargins(0, 0, 0, 0)

        lbl_excel = QLabel("Файл Excel:")
        layout_excel.addWidget(lbl_excel)

        self.entry_excel = QLineEdit()
        self.entry_excel.setPlaceholderText("Выберите Excel файл с реестром пунктов...")
        default_path = CONFIG.get("DEFAULT_EXCEL_PATH", "").strip()
        if default_path:
            self.entry_excel.setText(default_path)
        layout_excel.addWidget(self.entry_excel)

        btn_browse_excel = QPushButton("📂")
        btn_browse_excel.setObjectName("browse_btn")
        btn_browse_excel.setFixedWidth(40)
        btn_browse_excel.setCursor(Qt.PointingHandCursor)
        btn_browse_excel.clicked.connect(self._browse_excel)
        layout_excel.addWidget(btn_browse_excel)
        
        self.input_stack.addWidget(page_excel)

        # -- Страница 1: Ручной ввод --
        page_manual = QWidget()
        layout_manual = QHBoxLayout(page_manual)
        layout_manual.setContentsMargins(0, 0, 0, 0)
        
        lbl_manual = QLabel("Пункты:")
        lbl_manual.setAlignment(Qt.AlignTop)
        layout_manual.addWidget(lbl_manual)
        
        self.text_manual = QTextEdit()
        self.text_manual.setPlaceholderText("Введите номера пунктов через пробел, запятую или с новой строки...\nПример: 2758, 2759, 2760")
        self.text_manual.setMaximumHeight(80)
        layout_manual.addWidget(self.text_manual)
        
        self.input_stack.addWidget(page_manual)

        # Путь к локальной папке
        lbl_local = QLabel("Локальная папка:")
        s_layout.addWidget(lbl_local, 2, 0)

        self.entry_local = QLineEdit()
        desktop = os.path.join(os.path.expanduser("~"), "Desktop", "ПНОС")
        self.entry_local.setText(desktop)
        self.entry_local.setPlaceholderText("Папка на рабочем столе для отчётов...")
        s_layout.addWidget(self.entry_local, 2, 1)

        btn_browse_local = QPushButton("📂")
        btn_browse_local.setObjectName("browse_btn")
        btn_browse_local.setFixedWidth(40)
        btn_browse_local.setCursor(Qt.PointingHandCursor)
        btn_browse_local.clicked.connect(self._browse_local)
        s_layout.addWidget(btn_browse_local, 2, 2)

        main_layout.addWidget(settings)
        
        main_layout.addStretch()
        
        # Общий прогресс-бар для Этапа 1
        self.stage1_progress = QProgressBar()
        self.stage1_progress.setValue(0)
        self.stage1_progress.setFixedHeight(12)
        self.stage1_progress.setTextVisible(False)
        main_layout.addWidget(self.stage1_progress)
        
        # Кнопки управления
        ctrl_layout = QHBoxLayout()
        self.btn_start_stage1 = QPushButton("Запустить создание структуры")
        self.btn_start_stage1.setObjectName("card_button")
        self.btn_start_stage1.setFixedWidth(260)
        self.btn_start_stage1.setCursor(Qt.PointingHandCursor)
        self.btn_start_stage1.clicked.connect(self._start_stage1)
        
        self.btn_stop_stage1 = QPushButton("⏹ Остановить")
        self.btn_stop_stage1.setObjectName("stop_button")
        self.btn_stop_stage1.setFixedWidth(150)
        self.btn_stop_stage1.setCursor(Qt.PointingHandCursor)
        self.btn_stop_stage1.setEnabled(False)
        self.btn_stop_stage1.clicked.connect(self._force_stop_active_stage)
        
        ctrl_layout.addStretch()
        ctrl_layout.addWidget(self.btn_start_stage1)
        ctrl_layout.addWidget(self.btn_stop_stage1)
        ctrl_layout.addStretch()
        main_layout.addLayout(ctrl_layout)

        # Лента активности для Этапа 1
        lbl_act = QLabel("История действий:")
        lbl_act.setStyleSheet("color: #64748b; font-weight: bold; font-size: 11px; margin-top: 10px;")
        main_layout.addWidget(lbl_act)
        
        self.stage1_activity = QListWidget()
        self.stage1_activity.setObjectName("activity_feed")
        self.stage1_activity.setFixedHeight(120)
        main_layout.addWidget(self.stage1_activity)

        # Интегрированный отчет для Этапа 1
        self.stage1_report_table = QTableWidget()
        self.stage1_report_table.setColumnCount(5)
        self.stage1_report_table.setHorizontalHeaderLabels(["№ Пункта", "Паспорт", "Первичка", "Стар. ЭПБ", "Результат"])
        self.stage1_report_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.stage1_report_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.stage1_report_table.setFixedHeight(180)
        self.stage1_report_table.hide()
        main_layout.addWidget(self.stage1_report_table)

    def _setup_settings_page(self):
        """Создает UI для страницы 'Настройки программы'."""
        from PySide6.QtWidgets import QGroupBox
        
        layout = QVBoxLayout(self.page_settings)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        title = QLabel("⚙️ Настройки программы")
        title.setObjectName("app_title")
        layout.addWidget(title)
        
        scroll_area = QWidget()
        scroll_layout = QVBoxLayout(scroll_area)
        scroll_layout.setSpacing(25)

        # Блок 1: Основные параметры
        group_main = QGroupBox("Автоматизация и поиск")
        group_main.setStyleSheet("QGroupBox { font-weight: bold; color: #1e293b; border: 1px solid #cbd5e1; border-radius: 8px; margin-top: 15px; padding: 15px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }")
        g1_layout = QVBoxLayout(group_main)
        
        self.chk_epb = QCheckBox("Искать и скачивать папку 'Стар. ЭПБ'")
        self.chk_epb.setChecked(True)
        self.chk_epb.setStyleSheet("font-size: 14px; font-weight: bold; color: #1e293b;")
        g1_layout.addWidget(self.chk_epb)

        lbl_epb_desc = QLabel("   Если галочка снята, программа не будет проверять наличие старых экспертиз.")
        lbl_epb_desc.setStyleSheet("color: #475569; font-size: 12px;")
        g1_layout.addWidget(lbl_epb_desc)
        
        self.chk_auto = QCheckBox("Автоматический режим (остановка только на ручном шаге 3)")
        self.chk_auto.setChecked(False)
        self.chk_auto.setStyleSheet("font-size: 14px; font-weight: bold; color: #1e293b; margin-top: 10px;")
        g1_layout.addWidget(self.chk_auto)
        
        scroll_layout.addWidget(group_main)

        # Блок 2: Пути к файлам (НОВОЕ)
        group_paths = QGroupBox("Пути к файлам")
        group_paths.setStyleSheet("QGroupBox { font-weight: bold; color: #1e293b; border: 1px solid #cbd5e1; border-radius: 8px; margin-top: 15px; padding: 15px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }")
        g_paths_layout = QVBoxLayout(group_paths)

        lbl_macro = QLabel("Мастер-файл с макросом (для 3 этапа):")
        lbl_macro.setStyleSheet("color: #1e293b; font-weight: bold; font-size: 13px;")
        g_paths_layout.addWidget(lbl_macro)

        macro_box = QHBoxLayout()
        self.edit_macro_path = QLineEdit()
        self.edit_macro_path.setPlaceholderText("Оставьте пустым для использования пути из config.py")
        self.edit_macro_path.setStyleSheet("color: #1e293b;")
        btn_browse_macro = QPushButton("📂 Выбрать")
        btn_browse_macro.setFixedWidth(100)
        btn_browse_macro.clicked.connect(self._browse_macro_path)
        macro_box.addWidget(self.edit_macro_path)
        macro_box.addWidget(btn_browse_macro)
        g_paths_layout.addLayout(macro_box)
        
        lbl_macro_hint = QLabel("Укажите путь к «ПНОС сводный график Перечни ЭПБ на 2026г.xlsm»")
        lbl_macro_hint.setStyleSheet("color: #64748b; font-size: 11px;")
        g_paths_layout.addWidget(lbl_macro_hint)
        scroll_layout.addWidget(group_paths)

        # Блок 3: Скорость работы
        group_speed = QGroupBox("Производительность")
        group_speed.setStyleSheet("QGroupBox { font-weight: bold; color: #1e293b; border: 1px solid #cbd5e1; border-radius: 8px; margin-top: 15px; padding: 15px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }")
        g2_layout = QVBoxLayout(group_speed)
        
        lbl_speed = QLabel("Скорость проверки Яндекс Диска:")
        lbl_speed.setStyleSheet("color: #1e293b; font-size: 13px;")
        g2_layout.addWidget(lbl_speed)

        self.btn_speed_fast = QPushButton("🚀 Максимальная")
        self.btn_speed_fast.setCheckable(True)
        self.btn_speed_fast.setChecked(True)
        self.btn_speed_fast.setCursor(Qt.PointingHandCursor)
        self.btn_speed_fast.setObjectName("nav_button")

        self.btn_speed_safe = QPushButton("🐢 Безопасная")
        self.btn_speed_safe.setCheckable(True)
        self.btn_speed_safe.setCursor(Qt.PointingHandCursor)
        self.btn_speed_safe.setObjectName("nav_button")

        speed_h = QHBoxLayout()
        speed_h.addWidget(self.btn_speed_fast)
        speed_h.addWidget(self.btn_speed_safe)
        g2_layout.addLayout(speed_h)
        
        self.btn_speed_fast.clicked.connect(lambda: self.btn_speed_safe.setChecked(False) or self.btn_speed_fast.setChecked(True))
        self.btn_speed_safe.clicked.connect(lambda: self.btn_speed_fast.setChecked(False) or self.btn_speed_safe.setChecked(True))
        scroll_layout.addWidget(group_speed)

        layout.addWidget(scroll_area)
        layout.addStretch()

    def _setup_summary_page(self):
        """Создает страницу финального общего отчета."""
        layout = QVBoxLayout(self.page_summary)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(30)

        title = QLabel("🏁 Работа завершена!")
        title.setObjectName("app_title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        self.summary_card = QFrame()
        self.summary_card.setObjectName("settings_panel")
        self.summary_card.setStyleSheet("background-color: #1e293b; border-radius: 12px; padding: 30px;")
        s_layout = QVBoxLayout(self.summary_card)
        
        self.lbl_summary_stats = QLabel("Общая статистика будет здесь...")
        self.lbl_summary_stats.setStyleSheet("font-size: 16px; color: #f8fafc;")
        self.lbl_summary_stats.setWordWrap(True)
        self.lbl_summary_stats.setAlignment(Qt.AlignCenter)
        s_layout.addWidget(self.lbl_summary_stats)
        
        layout.addWidget(self.summary_card)

        btn_finish = QPushButton("Вернуться к первому шагу")
        btn_finish.setObjectName("card_button")
        btn_finish.setFixedWidth(300)
        btn_finish.clicked.connect(lambda: self._switch_page(0))
        
        layout.addWidget(btn_finish, alignment=Qt.AlignCenter)
        layout.addStretch()

    def _add_activity(self, stage_idx: int, message: str, category: str = "info"):
        """
        Добавляет запись в ленту активности конкретного этапа.
        category: 'info', 'done', 'wait', 'warn', 'error'
        """
        icons = {
            "info": "ℹ️",
            "done": "✅",
            "wait": "🔄",
            "warn": "⚠️",
            "error": "❌"
        }
        icon = icons.get(category, "•")
        item_text = f"{icon} {message}"
        
        # Индекс 0 - это этап 1, 1-4 - это этапы 2-5
        if stage_idx == 0:
            lw = self.stage1_activity
        else:
            lw = self.stub_pages[stage_idx-1]["activity_list"]
            
        item = QListWidgetItem(item_text)
        lw.addItem(item)
        lw.scrollToBottom()
        
        # Если это ошибка, дублируем в технический лог
        if category == "error":
            self._log_error(message)

    def _log_error(self, details: str):
        """Пишет подробную техническую ошибку в файл."""
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
        """Печать в консоль и в лог-файл если похоже на ошибку."""
        print(text)
        if "ошибка" in text.lower() or "error" in text.lower() or "❌" in text:
            self._log_error(text)

    # ─── Вспомогательные методы ────────────────────────────────────

    def _set_mode(self, mode: str):
        self._mode = mode
        if mode == "excel":
            self.btn_mode_excel.setChecked(True)
            self.btn_mode_excel.setProperty("active", True)
            self.btn_mode_manual.setChecked(False)
            self.btn_mode_manual.setProperty("active", False)
            self.input_stack.setCurrentIndex(0)
        else:
            self.btn_mode_excel.setChecked(False)
            self.btn_mode_excel.setProperty("active", False)
            self.btn_mode_manual.setChecked(True)
            self.btn_mode_manual.setProperty("active", True)
            self.input_stack.setCurrentIndex(1)
            
        # Обновляем стили
        self.btn_mode_excel.style().unpolish(self.btn_mode_excel)
        self.btn_mode_excel.style().polish(self.btn_mode_excel)
        self.btn_mode_manual.style().unpolish(self.btn_mode_manual)
        self.btn_mode_manual.style().polish(self.btn_mode_manual)

    def _browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выбрать Excel файл", "",
            "Excel Files (*.xlsx *.xlsm);;All Files (*)"
        )
        if path:
            self.entry_excel.setText(path)

    def _browse_local(self):
        folder = QFileDialog.getExistingDirectory(self, "Выбрать папку")
        if folder:
            self.entry_local.setText(folder)


    def _get_excel_path(self) -> str | None:
        path = self.entry_excel.text().strip()
        if not path:
            QMessageBox.warning(self, "Ошибка", "Укажите путь к файлу Excel.")
            return None
        if not os.path.exists(path):
            QMessageBox.warning(self, "Ошибка", f"Файл не найден:\n{path}")
            return None
        return path

    def _get_local_root(self) -> str:
        path = self.entry_local.text().strip()
        if not path:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop", "ПНОС")
            self.entry_local.setText(desktop)
            return desktop
        return path

    # ─── Запуск Этапа 1 ───────────────────────────────────────────

    def _start_stage1(self):
        local_root = self._get_local_root()

        excel_path = None
        manual_points = None

        if self._mode == "excel":
            excel_path = self._get_excel_path()
            if not excel_path:
                return
        else:
            import re
            text = self.text_manual.toPlainText()
            # Извлекаем все числа
            raw_points = re.findall(r'\d+', text)
            # Убираем дубликаты, сохраняя порядок
            manual_points = []
            for p in raw_points:
                if p not in manual_points:
                    manual_points.append(p)
            
            if not manual_points:
                QMessageBox.warning(self, "Ошибка", "Укажите хотя бы один номер пункта.")
                return

        self._start_time = time.time()
        self._stage_results = {}
        self.btn_start_stage1.setEnabled(False)
        self.btn_start_stage1.setText("Выполняется...")
        self.btn_stop_stage1.setEnabled(True)  # Включаем СТОП
        self.stage1_progress.setValue(0)
        self.stage1_report_table.hide()

        need_epb = self.chk_epb.isChecked()
        max_threads = 10 if self.btn_speed_fast.isChecked() else 3

        worker = Stage1Worker(
            excel_path=excel_path, 
            local_root=local_root, 
            manual_points=manual_points, 
            need_epb=need_epb,
            max_threads=max_threads,
            parent=self
        )
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(0, msg, cat))
        worker.progress.connect(self.stage1_progress.setValue)
        worker.report_ready.connect(self._update_stage1_table)
        worker.finished_ok.connect(lambda ok: self._on_stage1_finished(ok))
        worker.start()
        self._workers["stage1"] = worker

    def _update_stage1_table(self, report: dict):
        """Интегрированное обновление таблицы для 1 этапа."""
        self._stage_results[1] = report
        details = report.get("details", [])
        if not details:
            return
            
        self.stage1_report_table.setRowCount(len(details))
        for row_idx, d in enumerate(details):
            point = str(d.get("point", ""))
            status_text = d.get("status", "")
            folders = d.get("folders", {})
            
            # Колонка 0: Пункт
            self.stage1_report_table.setItem(row_idx, 0, QTableWidgetItem(point))
            
            # Колонки 1,2,3 - Папки
            col_map = {"Паспорт": 1, "Первичка": 2, "Стар. ЭПБ": 3}
            for folder_name, col_idx in col_map.items():
                f_status = folders.get(folder_name, False)
                item = QTableWidgetItem("✅" if f_status else "❌")
                item.setTextAlignment(Qt.AlignCenter)
                self.stage1_report_table.setItem(row_idx, col_idx, item)
                
            # Колонка 4 - Общий результат
            res_item = QTableWidgetItem(status_text)
            if "успешно" in status_text.lower():
                res_item.setForeground(QColor("#2d6a4f"))
            else:
                res_item.setForeground(QColor("#e63946"))
            self.stage1_report_table.setItem(row_idx, 4, res_item)
            
        self.stage1_report_table.show()

    def _on_stage1_finished(self, success: bool):
        self.btn_stop_stage1.setEnabled(False) # Выключаем СТОП
        if success:
            self.btn_start_stage1.setText("✓ Выполнено (Догрузить ещё)")
            self.btn_start_stage1.setEnabled(True)
            # Разблокируем Этап 2 в сайдбаре
            if len(self.nav_buttons) > 1:
                self.nav_buttons[1].setEnabled(True)
                self._switch_page(1)
                
            if getattr(self, "chk_auto", None) and self.chk_auto.isChecked():
                self.stub_pages[0]["btn_run"].click()
        else:
            self.btn_start_stage1.setText("Повторить")
            self.btn_start_stage1.setEnabled(True)
        self._update_global_progress()

    def _skip_stage(self, stage_num: int):
        """Пропускает этап и включает следующий."""
        self._on_stub_finished(stage_num - 1, True)

    def _start_stage2(self):
        local_root = self._get_local_root()
        p = self.stub_pages[0]
        btn, prog, act_list, stop, report_table = p["btn_run"], p["progress"], p["activity_list"], p["btn_stop"], p["report_table"]
        preview_table = p["preview_table"]
        
        btn.setText("Выполняется...")
        btn.setEnabled(False)
        stop.setEnabled(True)
        prog.setValue(0)
        report_table.hide()
        
        if preview_table:
            preview_table.clear()
            preview_table.hide()
            
        worker = Stage2Worker(local_root, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(1, msg, cat))
        worker.progress.connect(prog.setValue)
        worker.table_preview.connect(lambda d: self._update_table_preview(d, 0))
        worker.report_ready.connect(lambda d: self._show_integrated_report(2, ["№ Пункта", "Акт найден", "Таблица 3 найдена", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_stub_finished(1, ok))
        worker.start()
        self._workers["stage2"] = worker

    def _start_manual_stage3(self):
        """Шаг 3: Ручной контроль (без воркера)."""
        p = self.stub_pages[1]
        btn, stop = p["btn_run"], p["btn_stop"]
        
        btn.setText("Выполнено (Продолжить)")
        btn.setEnabled(True)
        stop.setEnabled(False)
        
        # На этом этапе просто ждем нажатия на кнопку "Запустить", 
        # которая для этого этапа переименована в подтверждение.
        # Но привязка в _setup_ui уже есть к этому методу.
        # Чтобы кнопка сработала как "финиш" этапа:
        self._on_stub_finished(2, True)

    def _show_integrated_report(self, stage_idx: int, columns: list, details: list):
        """Отображает отчет прямо в интерфейсе шага. stage_idx: 2-6"""
        self._stage_results[stage_idx] = details
        # stub_pages[0] это Шаг 2, поэтому вычитаем 2
        table = self.stub_pages[stage_idx-2]["report_table"]
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.setRowCount(len(details))
        
        for row_idx, d in enumerate(details):
            for col_idx, col_name in enumerate(columns):
                val = d.get(col_name, "")
                if isinstance(val, bool):
                    val = "✅" if val else "❌"
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignCenter)
                if col_name == "Результат" and "успешно" in str(val).lower():
                    item.setForeground(QColor("#2d6a4f"))
                elif col_name == "Результат":
                    item.setForeground(QColor("#e63946"))
                table.setItem(row_idx, col_idx, item)
        table.show()

    def _update_table_preview(self, data: list, stage_idx: int = 0):
        """Обновляет таблицу-превью (0 для этапа 2, 1 для этапа 3)."""
        p = self.stub_pages[stage_idx]
        preview_table = p["preview_table"]
        if not preview_table or not data:
            return
            
        rows = len(data)
        cols = max(len(r) for r in data) if rows > 0 else 0
        
        preview_table.setRowCount(rows)
        preview_table.setColumnCount(cols)
        
        for r_idx, row in enumerate(data):
            for c_idx, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignCenter)
                preview_table.setItem(r_idx, c_idx, item)
                
        preview_table.show()

    def _start_stage3(self):
        local_root = self._get_local_root()
        p = self.stub_pages[2] # Шаг 4: Запуск макросов
        btn, prog, act_list, stop, report_table = p["btn_run"], p["progress"], p["activity_list"], p["btn_stop"], p["report_table"]
        preview_table = p["preview_table"]
        
        btn.setText("Выполняется...")
        btn.setEnabled(False)
        stop.setEnabled(True)
        prog.setValue(0)
        report_table.hide()
        
        if preview_table:
            preview_table.clear()
            preview_table.hide()
            
        macro_path = self.edit_macro_path.text().strip() or None
        worker = Stage3Worker(local_root, macro_master_path=macro_path, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(3, msg, cat))
        worker.progress.connect(prog.setValue)
        worker.table_preview.connect(lambda d: self._update_table_preview(d, 1))
        worker.report_ready.connect(lambda d: self._show_integrated_report(4, ["№ Пункта", "Паспорт", "Excel", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_stub_finished(3, ok))
        worker.start()
        self._workers["stage3"] = worker

    def _start_stage4(self):
        local_root = self._get_local_root()
        p = self.stub_pages[3] # Шаг 5: Создание отчета
        btn, prog, act_list, stop, report_table = p["btn_run"], p["progress"], p["activity_list"], p["btn_stop"], p["report_table"]
        
        btn.setText("Выполняется...")
        btn.setEnabled(False)
        stop.setEnabled(True)
        prog.setValue(0)
        report_table.hide()
        
        worker = Stage4Worker(local_root, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(4, msg, cat))
        worker.progress.connect(prog.setValue)
        worker.report_ready.connect(lambda d: self._show_integrated_report(5, ["№ Пункта", "Макрос Excel", "Слияние Word", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_stub_finished(4, ok))
        worker.start()
        self._workers["stage4"] = worker

    def _start_stage5(self):
        local_root = self._get_local_root()
        p = self.stub_pages[4] # Шаг 6: Вставка картинок
        btn, prog, act_list, stop, report_table = p["btn_run"], p["progress"], p["activity_list"], p["btn_stop"], p["report_table"]
        btn.setText("Выполняется...")
        btn.setEnabled(False)
        stop.setEnabled(True)
        prog.setValue(0)
        report_table.hide()
        
        worker = Stage5Worker(local_root, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(5, msg, cat))
        worker.progress.connect(prog.setValue)
        worker.report_ready.connect(lambda d: self._show_integrated_report(6, ["№ Пункта", "Фото", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_stub_finished(5, ok))
        worker.start()
        self._workers["stage5"] = worker

    def _on_stub_finished(self, step_idx: int, success: bool):
        p = self.stub_pages[step_idx - 1]
        btn, stop = p["btn_run"], p["btn_stop"]
        stop.setEnabled(False)
        if success:
            btn.setText("✓ Выполнено (Повторить)")
            btn.setEnabled(True)
            self._add_activity(step_idx, "Этап успешно завершен!", "done")
            # Активировать следующий этап если есть
            next_actual_stage = step_idx + 1
            if next_actual_stage < len(self.nav_buttons):
                self.nav_buttons[next_actual_stage].setEnabled(True)
                self._switch_page(next_actual_stage)
                
                # Если авторежим и это не переход к ручному шагу 3 (индекс 2)
                if getattr(self, "chk_auto", None) and self.chk_auto.isChecked():
                    if next_actual_stage != 2:
                        from PySide6.QtCore import QTimer
                        QTimer.singleShot(1000, self.stub_pages[next_actual_stage - 1]["btn_run"].click)
            else:
                # Все этапы завершены -> Перейти к глобальному отчету
                self._calculate_global_summary()
                self._switch_page(6)
        else:
            btn.setText("Повторить")
            btn.setEnabled(True)
        self._update_global_progress()

    def _update_global_progress(self):
        completed = 0
        s1 = self.btn_start_stage1.text()
        if "Завершено" in s1 or "Готово" in s1 or "Выполнено" in s1:
            completed += 1
        for p in self.stub_pages:
            t = p["btn_run"].text()
            if "Готово" in t or "Выполнено" in t:
                completed += 1
        
        self.global_progress.setValue(int(completed * (100 / 6)))

    def _force_stop_active_stage(self):
        """Принудительная остановка текущего процесса и очистка COM."""
        if self._is_stopping:
            return
        self._is_stopping = True
        
        try:
            self._append_log("\n🛑 ПРИНУДИТЕЛЬНАЯ ОСТАНОВКА ПОЛЬЗОВАТЕЛЕМ...")
            
            # 1. Останавливаем воркеры
            workers_to_stop = list(self._workers.values())
            for worker in workers_to_stop:
                if worker:
                    try:
                        # Сначала отключаем все сигналы, чтобы воркер "осиротел" и не дергал UI
                        worker.disconnect() 
                        
                        if worker.isRunning():
                            worker.cancel() # Флаг для мягкой остановки
                            worker.quit()   # Запрос на выход из event loop
                            # Даем немного времени
                            for _ in range(15): 
                                if worker.wait(100):
                                    break
                                QApplication.processEvents()
                    except Exception as e:
                        print(f"Ошибка при отключении воркера: {e}")
            
            self._workers = {}

            # 2. Убиваем процессы через taskkill
            for proc in ["EXCEL.EXE", "WINWORD.EXE"]:
                try:
                    subprocess.Popen(f'taskkill /F /IM {proc}', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception: pass
            
            self._append_log("🧹 Принудительная очистка офисных процессов.")

            # 3. Сбрасываем UI
            self.btn_start_stage1.setEnabled(True)
            self.btn_start_stage1.setText("Запустить заново")
            self.btn_stop_stage1.setEnabled(False)
            
            if hasattr(self, 'stub_pages'):
                for p in self.stub_pages:
                    try:
                        p["btn_run"].setEnabled(True)
                        p["btn_run"].setText("Запустить заново")
                        p["btn_stop"].setEnabled(False)
                    except Exception: 
                        continue
            
            self._append_log("✅ Система готова к новому запуску.")
            
        except Exception as e:
            msg = f"⚠️ Ошибка в процессе остановки: {e}"
            print(msg)
            self._log_error(msg)
        finally:
            self._is_stopping = False

    def _calculate_global_summary(self):
        """Подсчитывает итоги по всем этапам."""
        total_time = 0
        if self._start_time:
            total_time = int(time.time() - self._start_time)
        
        minutes = total_time // 60
        seconds = total_time % 60
        
        success_count = 0
        error_count = 0
        errors_list = []

        # Считаем по 1 этапу
        st1 = self._stage_results.get(1, {})
        success_count += st1.get("created", 0)
        error_count += st1.get("not_created", 0)
        
        # Считаем по остальным (там списки словарей)
        for stage_idx in [2, 4, 5, 6]:
            results_val = self._stage_results.get(stage_idx)
            if not results_val or not isinstance(results_val, list):
                continue
            for res in results_val:
                res_str = str(res.get("Результат", "")).lower()
                if "успешно" in res_str or "✅" in res_str:
                    # Мы не плюсуем success_count здесь, так как пункты те же самые, 
                    # успех на 5 этапе означает успех всей цепочки.
                    pass
                else:
                    err_msg = f"Шаг {stage_idx}, п.{res.get('№ Пункта')}: {res.get('Результат')}"
                    errors_list.append(err_msg)

        summary_text = (
            f"⏱ <b>Время выполнения:</b> {minutes} мин. {seconds} сек.<br><br>"
            f"✅ <b>Пунктов обработано:</b> {success_count}<br>"
            f"❌ <b>Ошибок зафиксировано:</b> {len(errors_list)}<br><br>"
        )
        
        if errors_list:
            summary_text += "<b>Детали ошибок:</b><br>"
            for err in errors_list[:10]: # Ограничим вывод
                summary_text += f"• {err}<br>"
            if len(errors_list) > 10:
                summary_text += f"<i>...и еще {len(errors_list)-10} ошибок.</i>"
        else:
            summary_text += "✨ Все этапы пройдены идеально!"

        self.lbl_summary_stats.setText(summary_text)
