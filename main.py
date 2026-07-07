# -*- coding: utf-8 -*-
"""
ПНОС — Пакетный генератор отчётов.
Точка входа приложения.
"""

import sys
import os

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont

from main_window import MainWindow
from theme import build_stylesheet


def load_stylesheet() -> str:
    """Собирает QSS для текущей темы (light/dark) из QSettings."""
    settings = QSettings("PNOS", "PNOS")
    theme_name = settings.value("theme", "light")
    return build_stylesheet(theme_name)


def main():
    # Подчищаем остатки прошлого обновления (<EXE>.old, PNOS_update.exe),
    # чтобы на рабочем столе не копились лишние иконки.
    try:
        from updater import cleanup_update_leftovers
        cleanup_update_leftovers()
    except Exception:
        pass

    app = QApplication(sys.argv)

    # Шрифт по умолчанию
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    # Применяем стили
    stylesheet = load_stylesheet()
    if stylesheet:
        app.setStyleSheet(stylesheet)

    # Создаём и показываем окно
    window = MainWindow()
    window.show()

    # Фоновый прогрев тяжёлых модулей (pandas, aiohttp, win32com, python-docx).
    # Они импортируются лениво, поэтому старт быстрый; прогрев в фоне делает
    # первый клик по этапу мгновенным, не задерживая появление окна.
    _prewarm_heavy_imports()

    sys.exit(app.exec())


def _prewarm_heavy_imports():
    """Импортирует модули воркеров в фоновом потоке после показа окна."""
    import threading

    def _worker():
        import importlib
        for mod in (
            "workers.stage1_worker", "workers.stage2_worker", "workers.stage3_worker",
            "workers.stage4_worker", "workers.stage5_worker", "workers.stage6_worker",
            "report_dialog",
        ):
            try:
                importlib.import_module(mod)
            except Exception:
                pass

    threading.Thread(target=_worker, daemon=True).start()


if __name__ == "__main__":
    main()
