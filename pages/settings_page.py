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
    QGroupBox, QScrollArea, QMessageBox, QApplication,
    QListWidget, QListWidgetItem, QInputDialog
)

from version import APP_VERSION
import user_data
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
        group_main = QGroupBox("Автоматизация")
        g1_layout = QVBoxLayout(group_main)

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

        # ── Блок: Списки для формы ручного контроля ──────────────
        scroll_layout.addWidget(self._build_list_editor(
            "specialists", "Специалисты (поле B22)",
            "Список ФИО в выпадающем меню шага «Ручной контроль»."))
        scroll_layout.addWidget(self._build_list_editor(
            "epb_orgs", "Организации ЭПБ (поле B52)",
            "Список организаций в выпадающем меню шага «Ручной контроль»."))

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

        # ── Блок 4: Обновление программы ─────────────────────────
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

        # ── Блок: Диагностика ────────────────────────────────────
        group_diag = QGroupBox("Диагностика")
        g_diag = QVBoxLayout(group_diag)
        lbl_diag = QLabel("Файл с техническими ошибками — пришлите его разработчику, "
                          "если программа повела себя некорректно.")
        lbl_diag.setObjectName("hint")
        lbl_diag.setWordWrap(True)
        g_diag.addWidget(lbl_diag)
        diag_row = QHBoxLayout()
        btn_open_log = QPushButton("Открыть лог ошибок")
        btn_open_log.setObjectName("segment")
        btn_open_log.setCursor(Qt.PointingHandCursor)
        btn_open_log.clicked.connect(self._open_error_log)
        btn_open_logdir = QPushButton("Папка логов")
        btn_open_logdir.setObjectName("segment")
        btn_open_logdir.setCursor(Qt.PointingHandCursor)
        btn_open_logdir.clicked.connect(self._open_log_folder)
        diag_row.addWidget(btn_open_log)
        diag_row.addWidget(btn_open_logdir)
        diag_row.addStretch()
        g_diag.addLayout(diag_row)
        scroll_layout.addWidget(group_diag)

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

    # ── Редактор списков (специалисты / организации) ─────────────
    def _build_list_editor(self, name: str, title: str, hint: str) -> QGroupBox:
        group = QGroupBox(title)
        gl = QVBoxLayout(group)

        lbl_hint = QLabel(hint)
        lbl_hint.setObjectName("hint")
        lbl_hint.setWordWrap(True)
        gl.addWidget(lbl_hint)

        row = QHBoxLayout()
        lst = QListWidget()
        lst.setObjectName("point_list")
        lst.setMaximumHeight(160)
        for val in user_data.get_list(name):
            lst.addItem(QListWidgetItem(val))
        row.addWidget(lst, 1)

        btn_col = QVBoxLayout()
        btn_col.setSpacing(6)

        def save():
            items = [lst.item(i).text() for i in range(lst.count())]
            user_data.set_list(name, items)

        def add():
            text, ok = QInputDialog.getText(self, "Добавить", "Новое значение:")
            if ok and text.strip():
                lst.addItem(QListWidgetItem(text.strip()))
                save()

        def edit():
            it = lst.currentItem()
            if not it:
                return
            text, ok = QInputDialog.getText(self, "Изменить", "Значение:", text=it.text())
            if ok and text.strip():
                it.setText(text.strip())
                save()

        def remove():
            r = lst.currentRow()
            if r >= 0:
                lst.takeItem(r)
                save()

        def move(delta):
            r = lst.currentRow()
            nr = r + delta
            if r < 0 or nr < 0 or nr >= lst.count():
                return
            it = lst.takeItem(r)
            lst.insertItem(nr, it)
            lst.setCurrentRow(nr)
            save()

        def reset():
            if QMessageBox.question(
                    self, "Сброс списка",
                    "Вернуть список к стандартным значениям?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No) == QMessageBox.Yes:
                user_data.reset_list(name)
                lst.clear()
                for val in user_data.get_list(name):
                    lst.addItem(QListWidgetItem(val))

        for caption, handler in (
            ("Добавить", add), ("Изменить", edit), ("Удалить", remove),
            ("Вверх", lambda: move(-1)), ("Вниз", lambda: move(1)),
            ("Стандартные", reset),
        ):
            b = QPushButton(caption)
            b.setObjectName("segment")
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(handler)
            btn_col.addWidget(b)
        btn_col.addStretch()
        row.addLayout(btn_col)
        gl.addLayout(row)
        return group

    # ── Диагностика ──────────────────────────────────────────────
    def _log_dir(self) -> str:
        return os.path.join(os.getcwd(), "logs")

    def _open_error_log(self):
        path = os.path.join(self._log_dir(), "app_errors.log")
        if os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось открыть лог:\n{e}")
        else:
            QMessageBox.information(
                self, "Лог пуст",
                "Файл лога ошибок ещё не создан — значит, ошибок не было.")

    def _open_log_folder(self):
        d = self._log_dir()
        os.makedirs(d, exist_ok=True)
        try:
            os.startfile(d)
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось открыть папку:\n{e}")
