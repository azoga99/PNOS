# -*- coding: utf-8 -*-
"""
ПНОС — Пакетный генератор отчётов.
Точка входа приложения.
"""

import sys
import os

from PySide6.QtCore import QSettings, Qt, QTimer, QRectF, QRect, QElapsedTimer
from PySide6.QtWidgets import QApplication, QWidget
from PySide6.QtGui import QFont, QPixmap, QPainter, QPen, QColor, QImage, QPainterPath

from main_window import MainWindow
from theme import build_stylesheet
from utils import resource_path


def load_stylesheet() -> str:
    """Собирает QSS для текущей темы (light/dark) из QSettings."""
    settings = QSettings("PNOS", "PNOS")
    theme_name = settings.value("theme", "light")
    return build_stylesheet(theme_name)


def _extract_logo_mark() -> QPixmap:
    """Вырезает из resources/icon.png сам знак логотипа (без текста «PNOS…»
    и белых полей): сканирует левую часть картинки и берёт границы непустых
    (не-белых) пикселей. Возвращает QPixmap знака или пустой при неудаче."""
    path = resource_path(os.path.join("resources", "icon.png"))
    img = QImage(path)
    if img.isNull():
        return QPixmap()
    img = img.convertToFormat(QImage.Format_ARGB32)
    w, h = img.width(), img.height()
    scan_w = int(w * 0.44)  # знак слева, текст «PNOS…» правее — сканируем только знак

    minx, miny, maxx, maxy = w, h, 0, 0
    found = False
    step = max(1, w // 400)
    for y in range(0, h, step):
        for x in range(0, scan_w, step):
            c = img.pixelColor(x, y)
            # «не-белый и не-прозрачный» пиксель = часть знака
            if c.alpha() > 20 and not (c.red() > 240 and c.green() > 240 and c.blue() > 240):
                found = True
                if x < minx: minx = x
                if y < miny: miny = y
                if x > maxx: maxx = x
                if y > maxy: maxy = y
    if not found or maxx <= minx or maxy <= miny:
        return QPixmap.fromImage(img)

    pad = int(min(w, h) * 0.02)
    minx = max(0, minx - pad); miny = max(0, miny - pad)
    maxx = min(w, maxx + pad); maxy = min(h, maxy + pad)
    return QPixmap.fromImage(img.copy(QRect(minx, miny, maxx - minx, maxy - miny)))


class SplashLoader(QWidget):
    """Лоадер запуска: знак логотипа на карточке + полоса прогресса,
    заполняющаяся за отведённое время (создаёт ощущение «прорисовки»)."""

    W, H = 300, 330
    DURATION_MS = 5000

    def __init__(self, duration_ms: int = 5000):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
                         | Qt.SplashScreen)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(self.W, self.H)
        self.DURATION_MS = duration_ms

        mark = _extract_logo_mark()
        self._mark = mark.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation) \
            if not mark.isNull() else QPixmap()

        self._t = QElapsedTimer()
        self._t.start()
        self._timer = QTimer(self)
        self._timer.setInterval(16)
        self._timer.timeout.connect(self.update)
        self._timer.start()

        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center().x() - self.W // 2, screen.center().y() - self.H // 2)

    def _progress(self) -> float:
        return max(0.0, min(1.0, self._t.elapsed() / self.DURATION_MS))

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        # Карточка со скруглением
        card = QRectF(10, 10, self.W - 20, self.H - 20)
        path = QPainterPath()
        path.addRoundedRect(card, 20, 20)
        p.fillPath(path, QColor("#ffffff"))
        p.setPen(QPen(QColor("#e2e8f0"), 1))
        p.drawPath(path)

        # Знак логотипа по центру верхней части
        if not self._mark.isNull():
            prog = self._progress()
            mx = (self.W - self._mark.width()) // 2
            my = 46
            # Лёгкое «проявление»: непрозрачность нарастает в первые 40% времени
            p.setOpacity(min(1.0, prog / 0.4) if prog < 0.4 else 1.0)
            p.drawPixmap(mx, my, self._mark)
            p.setOpacity(1.0)

        # Название
        p.setPen(QColor("#1e3a5f"))
        f = QFont("Segoe UI", 20, QFont.Bold)
        f.setLetterSpacing(QFont.AbsoluteSpacing, 2)
        p.setFont(f)
        p.drawText(QRect(0, 214, self.W, 34), Qt.AlignCenter, "ПНОС")

        p.setPen(QColor("#64748b"))
        p.setFont(QFont("Segoe UI", 9))
        p.drawText(QRect(0, 250, self.W, 20), Qt.AlignCenter, "Пакетный генератор отчётов")

        # Полоса прогресса-«прорисовки»
        bar = QRectF(48, 288, self.W - 96, 6)
        p.setPen(Qt.NoPen)
        p.fillRect(bar, QColor("#e2e8f0"))
        fill = QRectF(bar.x(), bar.y(), bar.width() * self._progress(), bar.height())
        fp = QPainterPath()
        fp.addRoundedRect(fill, 3, 3)
        p.fillPath(fp, QColor("#2563eb"))
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
    splash = SplashLoader(SPLASH_MS)
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
