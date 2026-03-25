# -*- coding: utf-8 -*-
"""
Страница «Настройки программы».
Вынесена из main_window.py для улучшения структуры проекта.
"""

import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QProgressBar, QLineEdit, QFileDialog, QCheckBox, QFrame,
    QGroupBox, QScrollArea, QMessageBox, QApplication
)

from version import APP_VERSION
from updater import UpdateWorker, apply_update_and_restart


class SettingsPage(QWidget):
    """Виджет страницы настроек программы."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._update_worker = None
        self._new_exe_path = ""
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        title = QLabel("⚙️ Настройки программы")
        title.setObjectName("app_title")
        layout.addWidget(title)

        # Контейнер для скролла
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setContentsMargins(0, 0, 10, 0)
        scroll_layout.setSpacing(25)

        # ── Блок 1: Автоматизация ─────────────────────────────────
        group_main = QGroupBox("Автоматизация и поиск")
        group_main.setStyleSheet(
            "QGroupBox { font-weight: bold; color: #1e293b; border: 1px solid #cbd5e1; "
            "border-radius: 8px; margin-top: 15px; padding: 15px; } "
            "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }"
        )
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

        # ── Блок 2: Пути к файлам ────────────────────────────────
        group_paths = QGroupBox("Пути к файлам")
        group_paths.setStyleSheet(
            "QGroupBox { font-weight: bold; color: #1e293b; border: 1px solid #cbd5e1; "
            "border-radius: 8px; margin-top: 15px; padding: 15px; } "
            "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }"
        )
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

        # ── Блок 3: Скорость работы ──────────────────────────────
        group_speed = QGroupBox("Производительность")
        group_speed.setStyleSheet(
            "QGroupBox { font-weight: bold; color: #1e293b; border: 1px solid #cbd5e1; "
            "border-radius: 8px; margin-top: 15px; padding: 15px; } "
            "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }"
        )
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

        # ── Блок 4: Обновление программы ─────────────────────────
        group_update = QGroupBox("Обновление программы")
        group_update.setStyleSheet(
            "QGroupBox { font-weight: bold; color: #1e293b; border: 1px solid #cbd5e1; "
            "border-radius: 8px; margin-top: 15px; padding: 15px; } "
            "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }"
        )
        g_upd_layout = QVBoxLayout(group_update)

        lbl_cur_ver = QLabel(f"Текущая версия: {APP_VERSION}")
        lbl_cur_ver.setStyleSheet("color: #1e293b; font-size: 13px; font-weight: bold;")
        g_upd_layout.addWidget(lbl_cur_ver)

        lbl_upd_desc = QLabel("Проверить наличие новой версии и скачать обновление с GitHub.")
        lbl_upd_desc.setStyleSheet("color: #475569; font-size: 12px;")
        g_upd_layout.addWidget(lbl_upd_desc)

        self.btn_check_update = QPushButton("🔍 Проверить обновления")
        self.btn_check_update.setCursor(Qt.PointingHandCursor)
        self.btn_check_update.setStyleSheet("""
            QPushButton {
                background-color: #2563eb; color: white; font-size: 14px;
                font-weight: bold; padding: 12px 20px; border-radius: 8px; border: none;
            }
            QPushButton:hover { background-color: #1d4ed8; }
            QPushButton:pressed { background-color: #1e40af; }
            QPushButton:disabled { background-color: #94a3b8; }
        """)
        self.btn_check_update.clicked.connect(self._check_for_updates)
        g_upd_layout.addWidget(self.btn_check_update)

        self.btn_install_update = QPushButton("⬇️ Установить и перезапустить")
        self.btn_install_update.setCursor(Qt.PointingHandCursor)
        self.btn_install_update.setStyleSheet("""
            QPushButton {
                background-color: #16a34a; color: white; font-size: 14px;
                font-weight: bold; padding: 12px 20px; border-radius: 8px; border: none;
            }
            QPushButton:hover { background-color: #15803d; }
            QPushButton:pressed { background-color: #166534; }
            QPushButton:disabled { background-color: #94a3b8; }
        """)
        self.btn_install_update.clicked.connect(self._install_update)
        self.btn_install_update.setVisible(False)
        g_upd_layout.addWidget(self.btn_install_update)

        self.update_progress = QProgressBar()
        self.update_progress.setValue(0)
        self.update_progress.setMaximum(100)
        self.update_progress.setFixedHeight(14)
        self.update_progress.setVisible(False)
        self.update_progress.setStyleSheet("""
            QProgressBar { border: none; background-color: #e2e8f0; border-radius: 7px; }
            QProgressBar::chunk { background-color: #2563eb; border-radius: 7px; }
        """)
        g_upd_layout.addWidget(self.update_progress)

        self.lbl_update_status = QLabel("")
        self.lbl_update_status.setWordWrap(True)
        self.lbl_update_status.setStyleSheet("color: #475569; font-size: 12px; margin-top: 5px;")
        g_upd_layout.addWidget(self.lbl_update_status)

        scroll_layout.addWidget(group_update)
        scroll_layout.addStretch()

        # Оборачиваем в QScrollArea
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setStyleSheet(
            "QScrollArea { background: transparent; } "
            "QScrollArea > QWidget > QWidget { background: transparent; }"
        )
        layout.addWidget(scroll_area)

    # ── Обработчики ──────────────────────────────────────────────

    def _browse_macro_path(self):
        """Выбор файла макроса."""
        path, _ = QFileDialog.getOpenFileName(self, "Выберите файл макроса", "", "Excel (*.xlsm)")
        if path:
            self.edit_macro_path.setText(os.path.normpath(path))

    def _check_for_updates(self):
        """Запуск проверки и скачивания обновлений с GitHub Releases."""
        self.btn_check_update.setEnabled(False)
        self.btn_check_update.setText("⏳ Проверка...")
        self.btn_install_update.setVisible(False)
        self.update_progress.setVisible(True)
        self.update_progress.setValue(0)
        self.lbl_update_status.setText("Подключение к GitHub...")
        self.lbl_update_status.setStyleSheet("color: #2563eb; font-size: 12px; margin-top: 5px;")

        self._update_worker = UpdateWorker()
        self._update_worker.status.connect(self._on_update_status)
        self._update_worker.download_progress.connect(self.update_progress.setValue)
        self._update_worker.finished_ok.connect(self._on_update_finished)
        self._update_worker.start()

    def _on_update_status(self, text: str):
        self.lbl_update_status.setText(text)

    def _on_update_finished(self, success: bool, message: str):
        self.btn_check_update.setEnabled(True)
        self.btn_check_update.setText("🔍 Проверить обновления")
        self.lbl_update_status.setText(message)

        if success and self._update_worker and self._update_worker.new_exe_path:
            self._new_exe_path = self._update_worker.new_exe_path
            self.btn_install_update.setVisible(True)
            self.lbl_update_status.setStyleSheet("color: #16a34a; font-size: 12px; font-weight: bold; margin-top: 5px;")
        elif success:
            self.update_progress.setValue(100)
            self.lbl_update_status.setStyleSheet("color: #16a34a; font-size: 12px; font-weight: bold; margin-top: 5px;")
        else:
            self.lbl_update_status.setStyleSheet("color: #dc2626; font-size: 12px; margin-top: 5px;")

    def _install_update(self):
        if not self._new_exe_path:
            return
        reply = QMessageBox.question(
            self, "Подтверждение обновления",
            "Программа будет закрыта и перезапущена с новой версией.\nПродолжить?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )
        if reply == QMessageBox.Yes:
            apply_update_and_restart(self._new_exe_path)
