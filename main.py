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


def _extract_logo_region(x0_frac: float, x1_frac: float) -> QImage:
    """Вырезает из resources/icon.png область с непустыми (не-белыми)
    пикселями в горизонтальной полосе [x0_frac..x1_frac] ширины картинки.
    Слева живёт знак, справа — текст «PNOS INDUSTRIAL SOFTWARE»."""
    path = resource_path(os.path.join("resources", "icon.png"))
    img = QImage(path)
    if img.isNull():
        return QImage()
    img = img.convertToFormat(QImage.Format_ARGB32)
    w, h = img.width(), img.height()
    x_from, x_to = int(w * x0_frac), int(w * x1_frac)

    minx, miny, maxx, maxy = w, h, 0, 0
    found = False
    step = max(1, w // 400)
    for y in range(0, h, step):
        for x in range(x_from, x_to, step):
            c = img.pixelColor(x, y)
            if c.alpha() > 20 and not (c.red() > 240 and c.green() > 240 and c.blue() > 240):
                found = True
                if x < minx: minx = x
                if y < miny: miny = y
                if x > maxx: maxx = x
                if y > maxy: maxy = y
    if not found or maxx <= minx or maxy <= miny:
        return img

    pad = int(min(w, h) * 0.02)
    minx = max(x_from, minx - pad); miny = max(0, miny - pad)
    maxx = min(x_to, maxx + pad); maxy = min(h, maxy + pad)
    return img.copy(QRect(minx, miny, maxx - minx, maxy - miny))


def _white_to_transparent(img: QImage, hard: int = 244, soft: int = 228) -> QImage:
    """Превращает белый фон вырезки в прозрачность: почти-белые пиксели
    (min(r,g,b) >= hard) становятся полностью прозрачными, зона soft..hard —
    плавный переход, чтобы края знака не были «рваными»."""
    img = img.convertToFormat(QImage.Format_ARGB32)
    for y in range(img.height()):
        for x in range(img.width()):
            c = img.pixelColor(x, y)
            if c.alpha() == 0:
                continue
            w = min(c.red(), c.green(), c.blue())
            if w >= hard:
                c.setAlpha(0)
                img.setPixelColor(x, y, c)
            elif w >= soft:
                c.setAlpha(int(c.alpha() * (hard - w) / (hard - soft)))
                img.setPixelColor(x, y, c)
    return img


def _silhouette(pm: QPixmap, color: QColor) -> QPixmap:
    """Одноцветный силуэт по альфа-каналу пиксмапа."""
    out = QPixmap(pm.size())
    out.fill(Qt.transparent)
    p = QPainter(out)
    p.drawPixmap(0, 0, pm)
    p.setCompositionMode(QPainter.CompositionMode_SourceIn)
    p.fillRect(out.rect(), color)
    p.end()
    return out


class SplashLoader(QWidget):
    """Лоадер запуска (минималистичный): квадратная белая карточка,
    по центру знак логотипа, «заливающийся» родным цветом снизу вверх."""

    W, H = 230, 230
    CARD = 190            # сторона карточки
    MARK_H = 110          # высота знака

    def __init__(self, duration_ms: int = 3000):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
                         | Qt.SplashScreen)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(self.W, self.H)
        self.DURATION_MS = duration_ms

        # Знак: масштабируем, убираем белый фон, готовим серый силуэт
        mark_img = _extract_logo_region(0.0, 0.44)
        self._mark = QPixmap()
        self._mark_gray = QPixmap()
        if not mark_img.isNull():
            scaled = _white_to_transparent(
                mark_img.scaledToHeight(self.MARK_H, Qt.SmoothTransformation))
            self._mark = QPixmap.fromImage(scaled)
            self._mark_gray = _silhouette(self._mark, QColor(148, 163, 184, 110))

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

        # Квадратная белая карточка со скруглением
        cx = (self.W - self.CARD) / 2
        cy = (self.H - self.CARD) / 2
        path = QPainterPath()
        path.addRoundedRect(QRectF(cx, cy, self.CARD, self.CARD), 18, 18)
        p.fillPath(path, QColor("#ffffff"))
        p.setPen(QPen(QColor("#e2e8f0"), 1))
        p.drawPath(path)

        if not self._mark.isNull():
            mx = (self.W - self._mark.width()) // 2
            my = (self.H - self._mark.height()) // 2
            # Полупрозрачный серый силуэт знака
            p.drawPixmap(mx, my, self._mark_gray)
            # Заливка цветом снизу вверх за всё время лоадера
            fill = self._progress()
            mh = self._mark.height()
            top = int(mh * (1.0 - fill))
            p.save()
            p.setClipRect(QRect(mx, my + top, self._mark.width(), mh - top))
            p.drawPixmap(mx, my, self._mark)
            p.restore()

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
    SPLASH_MS = 3000
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
