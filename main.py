# -*- coding: utf-8 -*-
"""
ПНОС — Пакетный генератор отчётов.
Точка входа приложения.
"""

import sys
import os

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont

from main_window import MainWindow
from utils import resource_path

def load_stylesheet() -> str:
    """Загружает QSS стили из файла."""
    style_path = resource_path(os.path.join("resources", "style.qss"))
    if os.path.exists(style_path):
        with open(style_path, "r", encoding="utf-8") as f:
            return f.read()
    return ""


def main():
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

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
