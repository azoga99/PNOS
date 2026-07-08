# -*- coding: utf-8 -*-
"""
Страница «Настройки программы».
Вынесена из main_window.py для улучшения структуры проекта.
"""

import os

from PySide6.QtCore import Qt, Signal, QSettings
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QProgressBar, QLineEdit, QFileDialog, QCheckBox, QFrame,
    QGroupBox, QScrollArea, QMessageBox, QApplication
)

from version import APP_VERSION
# updater импортируется лениво (внутри методов), чтобы не тянуть requests при старте.


class SettingsPage(QWidget):
    """Виджет страницы настроек программы."""

    theme_changed = Signal(str)  # 'light' / 'dark'

    def __init__(self, parent=None):
        super().__init__(parent)
        self._update_worker = None
        self._new_exe_path = ""
        self._settings = QSettings("PNOS", "PNOS")
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        title = QLabel("Настройки")
        title.setObjectName("app_title")
        layout.addWidget(title)

        # Контейнер для скролла
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setContentsMargins(0, 0, 10, 0)
        scroll_layout.setSpacing(25)

        # ── Блок 1: Автоматизация ─────────────────────────────────
        group_main = QGroupBox("Автоматизация и поиск")
        g1_layout = QVBoxLayout(group_main)

        self.chk_epb = QCheckBox("Искать и скачивать папку «Стар. ЭПБ»")
        self.chk_epb.setChecked(True)
        g1_layout.addWidget(self.chk_epb)

        lbl_epb_desc = QLabel("Если галочка снята, программа не будет проверять наличие старых экспертиз.")
        lbl_epb_desc.setObjectName("hint")
        lbl_epb_desc.setWordWrap(True)
        g1_layout.addWidget(lbl_epb_desc)

        self.chk_auto = QCheckBox("Автоматический режим (остановка только на ручном шаге 3)")
        self.chk_auto.setChecked(True)
        g1_layout.addWidget(self.chk_auto)

        scroll_layout.addWidget(group_main)

        # ── Блок 2: Оформление ───────────────────────────────────
        group_theme = QGroupBox("Оформление")
        g_theme_layout = QVBoxLayout(group_theme)

        lbl_theme = QLabel("Тема интерфейса:")
        g_theme_layout.addWidget(lbl_theme)

        self.btn_theme_light = QPushButton("Светлая")
        self.btn_theme_light.setCheckable(True)
        self.btn_theme_light.setCursor(Qt.PointingHandCursor)
        self.btn_theme_light.setObjectName("segment")

        self.btn_theme_dark = QPushButton("Тёмная")
        self.btn_theme_dark.setCheckable(True)
        self.btn_theme_dark.setCursor(Qt.PointingHandCursor)
        self.btn_theme_dark.setObjectName("segment")

        cur_theme = self._settings.value("theme", "light")
        self.btn_theme_light.setChecked(cur_theme != "dark")
        self.btn_theme_dark.setChecked(cur_theme == "dark")

        theme_h = QHBoxLayout()
        theme_h.addWidget(self.btn_theme_light)
        theme_h.addWidget(self.btn_theme_dark)
        theme_h.addStretch()
        g_theme_layout.addLayout(theme_h)

        self.btn_theme_light.clicked.connect(lambda: self._select_theme("light"))
        self.btn_theme_dark.clicked.connect(lambda: self._select_theme("dark"))

        scroll_layout.addWidget(group_theme)

        # ── Блок 3: Пути к файлам ────────────────────────────────
        group_paths = QGroupBox("Пути к файлам")
        g_paths_layout = QVBoxLayout(group_paths)

        lbl_macro = QLabel("Мастер-файл с макросом (для шага 4):")
        g_paths_layout.addWidget(lbl_macro)

        macro_box = QHBoxLayout()
        self.edit_macro_path = QLineEdit()
        self.edit_macro_path.setPlaceholderText("Оставьте пустым для использования пути из config.py")
        btn_browse_macro = QPushButton("📂 Выбрать")
        btn_browse_macro.setFixedWidth(110)
        btn_browse_macro.setCursor(Qt.PointingHandCursor)
        btn_browse_macro.clicked.connect(self._browse_macro_path)
        macro_box.addWidget(self.edit_macro_path)
        macro_box.addWidget(btn_browse_macro)
        g_paths_layout.addLayout(macro_box)

        lbl_macro_hint = QLabel("Укажите путь к «ПНОС сводный график Перечни ЭПБ на 2026г.xlsm»")
        lbl_macro_hint.setObjectName("hint")
        lbl_macro_hint.setWordWrap(True)
        g_paths_layout.addWidget(lbl_macro_hint)
        scroll_layout.addWidget(group_paths)

        # ── Блок 4: Скорость работы ──────────────────────────────
        group_speed = QGroupBox("Производительность")
        g2_layout = QVBoxLayout(group_speed)

        lbl_speed = QLabel("Скорость проверки Яндекс Диска:")
        g2_layout.addWidget(lbl_speed)

        self.btn_speed_fast = QPushButton("Максимальная")
        self.btn_speed_fast.setCheckable(True)
        self.btn_speed_fast.setChecked(True)
        self.btn_speed_fast.setCursor(Qt.PointingHandCursor)
        self.btn_speed_fast.setObjectName("segment")

        self.btn_speed_safe = QPushButton("Безопасная")
        self.btn_speed_safe.setCheckable(True)
        self.btn_speed_safe.setCursor(Qt.PointingHandCursor)
        self.btn_speed_safe.setObjectName("segment")

        speed_h = QHBoxLayout()
        speed_h.addWidget(self.btn_speed_fast)
        speed_h.addWidget(self.btn_speed_safe)
        speed_h.addStretch()
        g2_layout.addLayout(speed_h)

        self.btn_speed_fast.clicked.connect(lambda: self.btn_speed_safe.setChecked(False) or self.btn_speed_fast.setChecked(True))
        self.btn_speed_safe.clicked.connect(lambda: self.btn_speed_fast.setChecked(False) or self.btn_speed_safe.setChecked(True))
        scroll_layout.addWidget(group_speed)

        # ── Блок 5: Обновление программы ─────────────────────────
        group_update = QGroupBox("Обновление программы")
        g_upd_layout = QVBoxLayout(group_update)

        lbl_cur_ver = QLabel(f"Текущая версия: {APP_VERSION}")
        lbl_cur_ver.setObjectName("form_field_label")
        g_upd_layout.addWidget(lbl_cur_ver)

        lbl_upd_desc = QLabel("Проверить наличие новой версии и скачать обновление с GitHub.")
        lbl_upd_desc.setObjectName("hint")
        lbl_upd_desc.setWordWrap(True)
        g_upd_layout.addWidget(lbl_upd_desc)

        self.btn_check_update = QPushButton("Проверить обновления")
        self.btn_check_update.setObjectName("blue_button")
        self.btn_check_update.setCursor(Qt.PointingHandCursor)
        self.btn_check_update.clicked.connect(self._check_for_updates)
        g_upd_layout.addWidget(self.btn_check_update)

        self.btn_install_update = QPushButton("Установить и перезапустить")
        self.btn_install_update.setObjectName("green_button")
        self.btn_install_update.setCursor(Qt.PointingHandCursor)
        self.btn_install_update.clicked.connect(self._install_update)
        self.btn_install_update.setVisible(False)
        g_upd_layout.addWidget(self.btn_install_update)

        self.update_progress = QProgressBar()
        self.update_progress.setValue(0)
        self.update_progress.setMaximum(100)
        self.update_progress.setFixedHeight(14)
        self.update_progress.setVisible(False)
        g_upd_layout.addWidget(self.update_progress)

        self.lbl_update_status = QLabel("")
        self.lbl_update_status.setObjectName("hint")
        self.lbl_update_status.setWordWrap(True)
        g_upd_layout.addWidget(self.lbl_update_status)

        scroll_layout.addWidget(group_update)
        scroll_layout.addStretch()

        # Оборачиваем в QScrollArea
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setFrameShape(QFrame.NoFrame)
        layout.addWidget(scroll_area)

    # ── Обработчики ──────────────────────────────────────────────

    def _select_theme(self, name: str):
        """Переключает тему и уведомляет главное окно."""
        self.btn_theme_light.setChecked(name == "light")
        self.btn_theme_dark.setChecked(name == "dark")
        self._settings.setValue("theme", name)
        self.theme_changed.emit(name)

    def _browse_macro_path(self):
        """Выбор файла макроса."""
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл макроса", "", "Excel (*.xlsm)")
        if path:
            self.edit_macro_path.setText(os.path.normpath(path))

    def _check_for_updates(self):
        """Запуск проверки и скачивания обновлений с GitHub Releases."""
        self.btn_check_update.setEnabled(False)
        self.btn_check_update.setText("Проверка…")
        self.btn_install_update.setVisible(False)
        self.update_progress.setVisible(True)
        self.update_progress.setValue(0)
        self.lbl_update_status.setText("Подключение к GitHub...")

        from updater import UpdateWorker
        self._update_worker = UpdateWorker()
        self._update_worker.status.connect(self._on_update_status)
        self._update_worker.download_progress.connect(self.update_progress.setValue)
        self._update_worker.finished_ok.connect(self._on_update_finished)
        self._update_worker.start()

    def _on_update_status(self, text: str):
        self.lbl_update_status.setText(text)

    def _on_update_finished(self, success: bool, message: str):
        self.btn_check_update.setEnabled(True)
        self.btn_check_update.setText("Проверить обновления")
        self.lbl_update_status.setText(message)

        if success and self._update_worker and self._update_worker.new_exe_path:
            self._new_exe_path = self._update_worker.new_exe_path
            self.btn_install_update.setVisible(True)
        elif success:
            self.update_progress.setValue(100)

    def _install_update(self):
        if not self._new_exe_path:
            return
        reply = QMessageBox.question(
            self, "Подтверждение обновления",
            "Программа будет закрыта и перезапущена с новой версией.\nПродолжить?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )
        if reply == QMessageBox.Yes:
            from updater import apply_update_and_restart
            apply_update_and_restart(self._new_exe_path)
