# -*- coding: utf-8 -*-
"""
Переиспользуемые виджеты интерфейса ПНОС.

StepCard — компактная карточка шага конвейера: номер-индикатор со статусом,
название, строка итога, кнопка действия, тонкий прогресс-бар и ховер-анимация
тени. Только PySide6 — тяжёлых импортов нет.
"""

from PySide6.QtCore import Qt, Signal, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QColor, QIcon, QPixmap, QPainter, QBrush
from PySide6.QtWidgets import (
    QFrame, QLabel, QPushButton, QProgressBar, QHBoxLayout, QVBoxLayout,
    QGraphicsDropShadowEffect, QSizePolicy, QComboBox,
)


class NoWheelComboBox(QComboBox):
    """QComboBox, игнорирующий колесо мыши: при прокрутке страницы значение
    в поле случайно не меняется. Раскрытый список скроллится как обычно."""

    def wheelEvent(self, event):
        event.ignore()

# Цвета статусов (совпадают по смыслу с палитрами theme.py)
STATUS_COLORS = {
    "idle": "#94a3b8",      # серый — не начат
    "run": "#6366f1",       # индиго — выполняется
    "done": "#10b981",      # зелёный — готово
    "error": "#ef4444",     # красный — ошибка
    "wait": "#f59e0b",      # янтарный — ждёт пользователя (ручной шаг)
}


def dot_icon(color: str, diameter: int = 10) -> QIcon:
    """Маленькая цветная точка-иконка (замена эмодзи в списках)."""
    pm = QPixmap(diameter, diameter)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QBrush(QColor(color)))
    p.setPen(Qt.NoPen)
    p.drawEllipse(0, 0, diameter, diameter)
    p.end()
    return QIcon(pm)


class StepCard(QFrame):
    """Карточка одного шага конвейера."""

    action_clicked = Signal()
    stop_clicked = Signal()

    GLYPHS = {"idle": None, "run": "…", "done": "✓", "error": "✕", "wait": "!"}

    def __init__(self, num: int, title: str, desc: str,
                 action_text: str = "Запустить", parent=None):
        super().__init__(parent)
        self.setObjectName("step_card")
        self.num = num
        self._status = "idle"
        self._desc = desc

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 10, 14, 10)
        root.setSpacing(6)

        row = QHBoxLayout()
        row.setSpacing(12)

        # Номер-индикатор
        self.badge = QLabel(str(num))
        self.badge.setObjectName("step_badge")
        self.badge.setFixedSize(30, 30)
        self.badge.setAlignment(Qt.AlignCenter)
        row.addWidget(self.badge)

        # Название + строка итога
        col = QVBoxLayout()
        col.setSpacing(1)
        self.lbl_title = QLabel(title)
        self.lbl_title.setObjectName("step_title")
        col.addWidget(self.lbl_title)
        self.lbl_summary = QLabel(desc)
        self.lbl_summary.setObjectName("step_summary")
        self.lbl_summary.setWordWrap(False)
        col.addWidget(self.lbl_summary)
        row.addLayout(col, 1)

        # Кнопки
        self.btn_stop = QPushButton("Стоп")
        self.btn_stop.setObjectName("step_stop")
        self.btn_stop.setCursor(Qt.PointingHandCursor)
        self.btn_stop.clicked.connect(self.stop_clicked.emit)
        self.btn_stop.hide()
        row.addWidget(self.btn_stop)

        self.btn_action = QPushButton(action_text)
        self.btn_action.setObjectName("step_action")
        self.btn_action.setCursor(Qt.PointingHandCursor)
        self.btn_action.setMinimumWidth(130)
        self.btn_action.clicked.connect(self.action_clicked.emit)
        row.addWidget(self.btn_action)

        root.addLayout(row)

        # Тонкий прогресс-бар (виден только при выполнении)
        self.progress = QProgressBar()
        self.progress.setObjectName("step_progress")
        self.progress.setFixedHeight(4)
        self.progress.setTextVisible(False)
        self.progress.hide()
        root.addWidget(self.progress)

        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        # Ховер-анимация тени
        self._shadow = QGraphicsDropShadowEffect(self)
        self._shadow.setColor(QColor(30, 41, 59, 55))
        self._shadow.setOffset(0, 2)
        self._shadow.setBlurRadius(0)
        self.setGraphicsEffect(self._shadow)
        self._anim = QPropertyAnimation(self._shadow, b"blurRadius", self)
        self._anim.setDuration(160)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)

        self._apply_status()

    # ── Ховер ──
    def enterEvent(self, event):
        self._animate_shadow(18)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._animate_shadow(0)
        super().leaveEvent(event)

    def _animate_shadow(self, target: float):
        self._anim.stop()
        self._anim.setStartValue(self._shadow.blurRadius())
        self._anim.setEndValue(target)
        self._anim.start()

    # ── API статуса ──
    def set_status(self, status: str):
        """idle / run / done / error / wait."""
        self._status = status
        self._apply_status()

    def status(self) -> str:
        return self._status

    def _apply_status(self):
        glyph = self.GLYPHS.get(self._status)
        self.badge.setText(glyph if glyph else str(self.num))
        self.badge.setProperty("status", self._status)
        self.badge.style().unpolish(self.badge)
        self.badge.style().polish(self.badge)

        running = self._status == "run"
        self.btn_stop.setVisible(running)
        self.progress.setVisible(running)
        self.btn_action.setEnabled(not running)

    def set_progress(self, value: int):
        self.progress.setValue(value)

    def set_summary(self, text: str):
        self.lbl_summary.setText(text if text else self._desc)

    def reset_summary(self):
        self.lbl_summary.setText(self._desc)

    def set_action_text(self, text: str):
        self.btn_action.setText(text)

    def set_action_enabled(self, enabled: bool):
        self.btn_action.setEnabled(enabled)
