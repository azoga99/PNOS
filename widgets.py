# -*- coding: utf-8 -*-
"""
Переиспользуемые виджеты интерфейса ПНОС.

SegmentBar — полоса из отдельных линий по этапам работы. DateCellDelegate
— ячейка таблицы с маской даты. NavRail — узкое боковое меню из значков с
подписями. EmptyState — подсказка по центру вместо пустого списка или таблицы.
Только PySide6 — тяжёлых импортов нет.
"""

import os

from PySide6.QtCore import Qt, Signal, QSize, QByteArray
from PySide6.QtGui import (
    QColor, QIcon, QPixmap, QPainter, QBrush, QGuiApplication,
)
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QFrame, QLabel, QHBoxLayout, QVBoxLayout, QSizePolicy, QComboBox,
    QWidget, QLineEdit, QStyledItemDelegate, QStackedWidget, QToolButton,
)

from utils import resource_path


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

# Цвета категорий сообщений единого лога
LOG_COLORS = {
    "info": "#94a3b8",
    "done": "#10b981",
    "wait": "#6366f1",
    "warn": "#f59e0b",
    "error": "#ef4444",
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


# ─── Полоса этапов ────────────────────────────────────────────────────

class SegmentBar(QWidget):
    """
    Полоса из отдельных линий: сколько этапов, столько сегментов.

    Обычный прогресс-бар здесь врал бы: число страниц, которые придётся
    прочитать, заранее неизвестно (модуль ходит по комплекту адресно), а этапы
    известны всегда. Заполненный сегмент = этап пройден, подсвеченный =
    выполняется сейчас; под полосой — название текущего этапа.

    Цвета — только из QSS (objectName «segment_line», свойство state).
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._names: list[str] = []
        self._lines: list[QFrame] = []

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(4)
        self._row = QHBoxLayout()
        self._row.setContentsMargins(0, 0, 0, 0)
        self._row.setSpacing(4)
        outer.addLayout(self._row)
        self.lbl = QLabel("")
        self.lbl.setObjectName("hint")
        outer.addWidget(self.lbl)
        self.setVisible(False)

    def set_phases(self, names):
        """Задать этапы и показать полосу пустой."""
        for line in self._lines:
            self._row.removeWidget(line)
            line.deleteLater()
        self._lines = []
        self._names = [str(n) for n in (names or [])]
        for _ in self._names:
            line = QFrame()
            line.setObjectName("segment_line")
            line.setProperty("state", "idle")
            line.setFixedHeight(4)
            line.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self._row.addWidget(line)
            self._lines.append(line)
        self.lbl.setText("")
        self.setVisible(bool(self._names))

    def set_current(self, index: int, name: str = ""):
        """Этапы до index — пройдены, index — выполняется."""
        for i, line in enumerate(self._lines):
            state = "done" if i < index else ("run" if i == index else "idle")
            if line.property("state") != state:
                line.setProperty("state", state)
                line.style().unpolish(line)
                line.style().polish(line)
        подпись = name or (self._names[index] if 0 <= index < len(self._names)
                           else "")
        self.lbl.setText(f"{подпись}…" if подпись else "Готово")

    def finish(self):
        self.set_current(len(self._names))

    def reset(self):
        self.set_phases([])


# ─── Ячейка таблицы с датой ───────────────────────────────────────────

DATE_MASK = "00.00.0000;_"


def date_cell_value(text: str) -> str:
    """
    Значение поля с маской даты: пустая маска («__.__.____») → пустая строка.

    Без этого в ячейку попадали бы подчёркивания, и дальше по коду они
    считались бы введённой датой.
    """
    s = (text or "").strip()
    if not any(c.isdigit() for c in s):
        return ""
    return s


class DateCellDelegate(QStyledItemDelegate):
    """Редактор ячейки таблицы с жёсткой маской ДД.ММ.ГГГГ (без календаря)."""

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setInputMask(DATE_MASK)
        return editor

    def setEditorData(self, editor, index):
        editor.setText((index.data() or "").strip())
        editor.setCursorPosition(0)

    def setModelData(self, editor, model, index):
        model.setData(index, date_cell_value(editor.text()))


# ─── Значки интерфейса ────────────────────────────────────────────────

def svg_pixmap(icon_name: str, color: str, size: int) -> QPixmap:
    """
    Однотонный значок resources/icons/<icon_name>.svg в цвете color.

    Значки нарисованы линиями цвета currentColor — он заменяется цветом токена
    темы, поэтому один файл служит обеим темам. Рисуется с запасом по
    плотности пикселей, чтобы при масштабе экрана 125–200 % не мылилось.
    """
    path = resource_path(os.path.join("resources", "icons", f"{icon_name}.svg"))
    try:
        with open(path, "r", encoding="utf-8") as f:
            svg = f.read().replace("currentColor", color)
    except OSError:
        svg = ""
    screen = QGuiApplication.primaryScreen()
    dpr = max(2.0, screen.devicePixelRatio() if screen else 1.0)
    pm = QPixmap(int(size * dpr), int(size * dpr))
    pm.fill(Qt.transparent)
    if svg:
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        QSvgRenderer(QByteArray(svg.encode("utf-8"))).render(p)
        p.end()
    pm.setDevicePixelRatio(dpr)
    return pm


# ─── Узкое боковое меню ───────────────────────────────────────────────

class _RailButton(QToolButton):
    """Пункт меню: значок, подпись под ним и счётчик в углу значка."""

    def __init__(self, caption: str, icon_size: int, parent=None):
        super().__init__(parent)
        self.setObjectName("rail_item")
        self.setText(caption)
        self.setCheckable(True)
        self.setAutoRaise(True)      # без него QToolButton не берёт значок наведения
        self.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        self.setIconSize(QSize(icon_size, icon_size))
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedSize(64, 54)
        self._icon_size = icon_size

        self.badge = QLabel("", self)
        self.badge.setObjectName("rail_badge")
        self.badge.setAlignment(Qt.AlignCenter)
        self.badge.setFixedHeight(18)
        self.badge.setMinimumWidth(18)
        self.badge.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.badge.hide()

    def set_badge(self, n: int):
        if n <= 0:
            self.badge.hide()
            return
        self.badge.setText(str(n) if n < 100 else "99+")
        self.badge.adjustSize()
        self._place_badge()
        self.badge.show()

    def _place_badge(self):
        # правый верхний угол значка, чуть наружу
        x = (self.width() + self._icon_size) // 2 - 8
        self.badge.move(min(x, self.width() - self.badge.width()), 1)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._place_badge()


class NavRail(QFrame):
    """
    Узкое боковое меню: значок программы сверху, пункты — значок и подпись
    мелким шрифтом под ним, «Настройки» отдельно внизу.

    Подписи и фон красит QSS (rail_item, rail_badge). Значки — SVG, их цвет
    подставляется из палитры темы в recolor(); после смены темы вызывайте
    recolor ещё раз.
    """

    item_clicked = Signal(str)

    WIDTH = 72
    ICON_SIZE = 22

    def __init__(self, logo_path: str = "", logo_tooltip: str = "", parent=None):
        super().__init__(parent)
        self.setObjectName("nav_rail")
        self.setFixedWidth(self.WIDTH)
        self._buttons: dict[str, _RailButton] = {}
        self._icon_names: dict[str, str] = {}
        self._palette: dict = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(4, 14, 4, 10)
        root.setSpacing(4)

        self.logo = QLabel()
        self.logo.setObjectName("rail_logo")
        self.logo.setAlignment(Qt.AlignCenter)
        self.logo.setToolTip(logo_tooltip)
        pm = QPixmap(logo_path) if logo_path else QPixmap()
        if not pm.isNull():
            screen = QGuiApplication.primaryScreen()
            dpr = screen.devicePixelRatio() if screen else 1.0
            pm = pm.scaled(int(32 * dpr), int(32 * dpr),
                           Qt.KeepAspectRatio, Qt.SmoothTransformation)
            pm.setDevicePixelRatio(dpr)
            self.logo.setPixmap(pm)
        root.addWidget(self.logo)
        root.addSpacing(12)

        self._top = QVBoxLayout()
        self._top.setSpacing(4)
        root.addLayout(self._top)
        root.addStretch(1)
        self._bottom = QVBoxLayout()
        self._bottom.setSpacing(4)
        root.addLayout(self._bottom)

    # ── Наполнение ──
    def add_item(self, key: str, caption: str, icon_name: str):
        self._add(self._top, key, caption, icon_name)

    def add_bottom_item(self, key: str, caption: str, icon_name: str):
        self._add(self._bottom, key, caption, icon_name)

    def _add(self, layout, key: str, caption: str, icon_name: str):
        btn = _RailButton(caption, self.ICON_SIZE, self)
        btn.clicked.connect(lambda checked=False, k=key: self.item_clicked.emit(k))
        layout.addWidget(btn, 0, Qt.AlignHCenter)
        self._buttons[key] = btn
        self._icon_names[key] = icon_name
        if self._palette:
            self._paint_icon(key)

    # ── Состояние ──
    def set_current(self, key: str):
        for k, btn in self._buttons.items():
            btn.setChecked(k == key)

    def set_badge(self, key: str, n: int):
        """Число в углу значка; 0 — скрыть."""
        btn = self._buttons.get(key)
        if btn:
            btn.set_badge(int(n or 0))

    def recolor(self, palette: dict):
        """Перекрасить значки под тему (palette — словарь theme.get_palette)."""
        self._palette = palette
        for key in self._buttons:
            self._paint_icon(key)

    def _paint_icon(self, key: str):
        name = self._icon_names[key]
        p = self._palette
        normal = svg_pixmap(name, p.get("rail_text", "#64748b"), self.ICON_SIZE)
        hover = svg_pixmap(name, p.get("rail_text_hover", "#1e293b"), self.ICON_SIZE)
        current = svg_pixmap(name, p.get("accent", "#6366f1"), self.ICON_SIZE)
        icon = QIcon()
        icon.addPixmap(normal, QIcon.Normal, QIcon.Off)
        icon.addPixmap(hover, QIcon.Active, QIcon.Off)
        icon.addPixmap(current, QIcon.Normal, QIcon.On)
        icon.addPixmap(current, QIcon.Active, QIcon.On)
        self._buttons[key].setIcon(icon)


# ─── Пустое состояние ─────────────────────────────────────────────────

class EmptyState(QStackedWidget):
    """
    Список или таблица — либо подсказка по центру на их месте.

    Пустая белая таблица ничего не говорит: непонятно, идёт ли работа,
    сломалось ли что-то и что нажать. Подсказка отвечает двумя строками —
    что происходит и что делать дальше.
    """

    def __init__(self, content: QWidget, title: str = "", text: str = "",
                 parent=None):
        super().__init__(parent)
        self.content = content
        self.addWidget(content)

        hint = QFrame()
        hint.setObjectName("empty_state")
        # Ширину строки держат поля (см. resizeEvent), а не maximumWidth
        # подписи: подпись с переносом и выравниванием по центру Qt считает
        # высотой в одну строку и обрезает остальное.
        self._hint_layout = QVBoxLayout(hint)
        self._hint_layout.setSpacing(6)
        self._hint_layout.addStretch(1)
        self.lbl_title = QLabel(title)
        self.lbl_title.setObjectName("empty_title")
        self.lbl_title.setAlignment(Qt.AlignCenter)
        self.lbl_title.setWordWrap(True)
        self._hint_layout.addWidget(self.lbl_title)
        self.lbl_text = QLabel(text)
        self.lbl_text.setObjectName("empty_text")
        self.lbl_text.setAlignment(Qt.AlignCenter)
        self.lbl_text.setWordWrap(True)
        self._hint_layout.addWidget(self.lbl_text)
        self._hint_layout.addStretch(1)
        self.addWidget(hint)

        if title or text:
            self.show_hint(title, text)

    def show_hint(self, title: str, text: str = ""):
        self.lbl_title.setText(title)
        self.lbl_text.setText(text)
        self.lbl_text.setVisible(bool(text))
        self.setCurrentIndex(1)

    def show_content(self):
        self.setCurrentIndex(0)

    def is_hint_shown(self) -> bool:
        return self.currentIndex() == 1

    TEXT_WIDTH = 460

    def resizeEvent(self, event):
        super().resizeEvent(event)
        side = max(24, (self.width() - self.TEXT_WIDTH) // 2)
        self._hint_layout.setContentsMargins(side, 24, side, 24)
