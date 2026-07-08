# -*- coding: utf-8 -*-
"""
ПНОС — Пакетный генератор отчётов.
Точка входа приложения.
"""

import sys
import os

from PySide6.QtCore import QSettings, Qt, QTimer, QRectF
from PySide6.QtWidgets import QApplication, QWidget
from PySide6.QtGui import QFont, QPixmap, QPainter, QPen, QColor

from main_window import MainWindow
from theme import build_stylesheet
from utils import resource_path


def load_stylesheet() -> str:
    """Собирает QSS для текущей темы (light/dark) из QSettings."""
    settings = QSettings("PNOS", "PNOS")
    theme_name = settings.value("theme", "light")
    return build_stylesheet(theme_name)


class SplashLoader(QWidget):
    """Лоадер запуска: иконка программы с вращающейся дугой вокруг неё."""

    SIZE = 180
    ICON = 84
    RING_MARGIN = 22

    def __init__(self):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
                         | Qt.SplashScreen)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(self.SIZE, self.SIZE)

        icon_path = resource_path(os.path.join("resources", "icon.png"))
        self._icon = QPixmap(icon_path)
        if not self._icon.isNull():
            self._icon = self._icon.scaled(
                self.ICON, self.ICON, Qt.KeepAspectRatio, Qt.SmoothTransformation)

        self._angle = 0
        self._timer = QTimer(self)
        self._timer.setInterval(16)
        self._timer.timeout.connect(self._tick)
        self._timer.start()

        # По центру экрана
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center().x() - self.SIZE // 2,
                  screen.center().y() - self.SIZE // 2)

    def _tick(self):
        self._angle = (self._angle + 6) % 360
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        # Иконка по центру
        if not self._icon.isNull():
            x = (self.width() - self._icon.width()) // 2
            y = (self.height() - self._icon.height()) // 2
            p.drawPixmap(x, y, self._icon)

        # Вращающаяся дуга вокруг иконки
        m = self.RING_MARGIN
        rect = QRectF(m, m, self.width() - 2 * m, self.height() - 2 * m)
        pen = QPen(QColor("#6366f1"), 4)
        pen.setCapStyle(Qt.RoundCap)
        p.setPen(pen)
        p.drawArc(rect, -self._angle * 16, 100 * 16)
        p.end()

    def finish(self):
        self._timer.stop()
        self.close()


def main():
    # Подчищаем остатки прошлого обновления (<EXE>.old, PNOS_update.exe),
    # чтобы на рабочем столе не копились лишние иконки.
    try:
        from updater import cleanup_update_leftovers
        cleanup_update_leftovers()
    except Exception:
        pass

    app = QApplication(sys.argv)

    # Лоадер: показывается минимум SPLASH_MS, пока окно готовится
    SPLASH_MS = 5000
    import time as _time
    splash = SplashLoader()
    splash.show()
    app.processEvents()
    splash_started = _time.time()

    # Шрифт по умолчанию
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    # Применяем стили
    stylesheet = load_stylesheet()
    if stylesheet:
        app.setStyleSheet(stylesheet)

    # Создаём окно; показываем его после того, как лоадер отработал своё
    window = MainWindow()
    elapsed_ms = int((_time.time() - splash_started) * 1000)
    remaining = max(0, SPLASH_MS - elapsed_ms)
    QTimer.singleShot(remaining, lambda: (window.show(), splash.finish()))

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
            "workers.manual3_worker",
        ):
            try:
                importlib.import_module(mod)
            except Exception:
                pass

    threading.Thread(target=_worker, daemon=True).start()


if __name__ == "__main__":
    main()
