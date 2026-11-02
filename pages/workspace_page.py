# -*- coding: utf-8 -*-
"""
«Рабочее место» — одна страница вместо «Процесса» и «Обзора».

Сверху вниз: шапка (папка, переключатель «Пункты / Отчёт о прогоне», блок
прогона со «Стоп»), строка семи шагов с общим состоянием, панель запуска,
таблица пунктов с точками по шагам, легенда и сворачиваемый журнал.

Виджет самостоятельный: сам ничего не запускает и не читает с диска, только
показывает то, что ему передали (`set_points`, `set_run_status`, `add_log`), и
отдаёт намерения сигналами. Строки для таблицы строит `workspace_model`.
Только PySide6 — тяжёлых импортов нет (страница грузится при старте).
"""

import re

from PySide6.QtCore import Qt, Signal, QRect, QSize, QEvent
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap, QFontMetrics
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame,
    QLineEdit, QStackedWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QCheckBox, QListWidget, QListWidgetItem, QProgressBar, QSizePolicy,
    QStyledItemDelegate, QStyle,
)

from widgets import STATUS_COLORS, LOG_COLORS, dot_icon
from workspace_model import STEP_NAMES, FORM_STEP, next_step

STEPS = range(1, 8)

STATUS_TEXT = {
    "done": "выполнен",
    "run": "выполняется",
    "wait": "ждёт",
    "error": "ошибка",
    "idle": "не начат",
}

# Порядок и подписи легенды
LEGEND = [("done", "готово"), ("run", "идёт"), ("wait", "ждёт / предупреждение"),
          ("error", "ошибка"), ("idle", "не начат")]

# Колонки таблицы: отметка, пункт, 7 шагов, состояние
COL_CHECK = 0
COL_POINT = 1
COL_STEP0 = 2            # колонка шага N = COL_STEP0 + N - 1
COL_STATE = COL_STEP0 + 7

ROLE_STATUS = Qt.UserRole + 1
ROLE_POINT = Qt.UserRole + 2

DOT = 10


def status_pixmap(status: str, diameter: int = DOT) -> QPixmap:
    """
    Точка статуса. «Не начат» — пустое кольцо: серую заливку на светлой
    таблице легко принять за «готово» при беглом взгляде.
    """
    if status != "idle":
        return dot_icon(STATUS_COLORS.get(status, STATUS_COLORS["idle"]),
                        diameter).pixmap(diameter, diameter)
    pm = QPixmap(diameter, diameter)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setPen(QPen(QColor(STATUS_COLORS["idle"]), 1.6))
    p.setBrush(Qt.NoBrush)
    p.drawEllipse(1, 1, diameter - 2, diameter - 2)
    p.end()
    return pm


def plural_points(n: int) -> str:
    """«1 пункт», «3 пункта», «5 пунктов»."""
    if n % 10 == 1 and n % 100 != 11:
        word = "пункт"
    elif n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
        word = "пункта"
    else:
        word = "пунктов"
    return f"{n} {word}"


def parse_point_numbers(text: str) -> list[str]:
    """Номера из поля ввода: только цифры, без дублей, в порядке ввода."""
    seen = []
    for num in re.findall(r"\d+", text or ""):
        if num not in seen:
            seen.append(num)
    return seen


def aggregate_status(statuses: list[str]) -> str:
    """Общее состояние шага по всем пунктам."""
    if not statuses:
        return "idle"
    for s in ("run", "error", "wait"):
        if s in statuses:
            return s
    return "done" if all(s == "done" for s in statuses) else "idle"


class ElidedLabel(QLabel):
    """Однострочная надпись, обрезаемая многоточием по ширине."""

    def __init__(self, text: str = "", parent=None):
        super().__init__(parent)
        self._full = text
        self.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.setMinimumWidth(40)
        self._refresh()

    def setText(self, text: str):
        self._full = text or ""
        self._refresh()

    def fullText(self) -> str:
        return self._full

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._refresh()

    def _refresh(self):
        fm = QFontMetrics(self.font())
        shown = fm.elidedText(self._full, Qt.ElideRight, max(0, self.width() - 2))
        super().setText(shown)
        self.setToolTip(self._full if shown != self._full else "")


class StepDotDelegate(QStyledItemDelegate):
    """Ячейка шага: только точка по центру (иконка у QTableWidget прижата влево)."""

    def paint(self, painter, option, index):
        opt = option
        self.initStyleOption(opt, index)
        opt.text = ""
        opt.icon = type(opt.icon)()
        style = opt.widget.style() if opt.widget else None
        if style:
            style.drawControl(QStyle.CE_ItemViewItem, opt, painter, opt.widget)
        status = index.data(ROLE_STATUS) or "idle"
        r = option.rect
        x = r.x() + (r.width() - DOT) // 2
        y = r.y() + (r.height() - DOT) // 2
        painter.drawPixmap(x, y, status_pixmap(status))

    def sizeHint(self, option, index):
        return QSize(DOT + 16, DOT + 12)


class StepChip(QFrame):
    """Элемент строки шагов: точка, «N Название», счётчик «готово из всех»."""

    def __init__(self, step: int, parent=None):
        super().__init__(parent)
        self.setObjectName("ws_step_chip")
        self.step = step
        h = QHBoxLayout(self)
        h.setContentsMargins(10, 6, 10, 6)
        h.setSpacing(7)
        self.dot = QLabel()
        self.dot.setFixedSize(DOT, DOT)
        h.addWidget(self.dot)
        num = QLabel(str(step))
        num.setObjectName("ws_step_num")
        h.addWidget(num)
        name = QLabel(STEP_NAMES[step])
        name.setObjectName("ws_step_name")
        h.addWidget(name)
        h.addStretch()
        self.count = QLabel("")
        self.count.setObjectName("ws_step_count")
        h.addWidget(self.count)
        self.set_state("idle", 0, 0, {})

    def set_state(self, status: str, done: int, total: int, counts: dict):
        self.dot.setPixmap(status_pixmap(status))
        self.count.setText(f"{done}/{total}" if total else "")
        self.setProperty("status", status)
        self.style().unpolish(self)
        self.style().polish(self)
        if not total:
            self.setToolTip("Пунктов нет")
            return
        parts = [f"готово {done} из {total}"]
        for key, word in (("run", "идёт"), ("wait", "ждёт"), ("error", "ошибка")):
            if counts.get(key):
                parts.append(f"{word} {counts[key]}")
        self.setToolTip(f"Шаг {self.step} · {STEP_NAMES[self.step]}: "
                        + ", ".join(parts))


class WorkspacePage(QWidget):
    """Рабочее место: пункты и шаги на одной странице."""

    add_points_requested = Signal(list)        # номера-строки из поля
    run_step_requested = Signal(int, list)     # шаг 1–7, номера пунктов
    run_to_end_requested = Signal(list)        # номера пунктов
    open_form_requested = Signal(str)          # номер пункта
    stop_requested = Signal()
    refresh_requested = Signal()
    reset_requested = Signal()                 # диалог сброса статуса пунктов
    folder_requested = Signal()                # выбрать другую папку с пунктами

    LOG_EXPANDED_HEIGHT = 190

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("ws_page")
        self._rows: list[dict] = []
        self._checked: set[str] = set()
        self._busy = False
        self._run_step = None
        self._filling = False
        self._log_count = 0
        self._log_expanded = False

        root = QVBoxLayout(self)
        root.setContentsMargins(2, 2, 6, 2)
        root.setSpacing(10)

        root.addLayout(self._build_header())

        self.tabs = QStackedWidget()
        self.tabs.addWidget(self._build_points_tab())
        self.report_container = QWidget()
        self.report_container.setObjectName("ws_report")
        rl = QVBoxLayout(self.report_container)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(10)
        self.tabs.addWidget(self.report_container)
        root.addWidget(self.tabs, 1)

        root.addWidget(self._build_log())

        self.set_folder_info("", 0)
        self.set_run_status(None, None, 0, "")
        self.set_points([])

    # ═══════════════════════════════════════════════════════════════
    # Построение
    # ═══════════════════════════════════════════════════════════════

    def _build_header(self):
        head = QHBoxLayout()
        head.setSpacing(16)

        title_col = QVBoxLayout()
        title_col.setSpacing(0)
        title = QLabel("Рабочее место")
        title.setObjectName("app_title")
        title_col.addWidget(title)
        folder_row = QHBoxLayout()
        folder_row.setSpacing(6)
        self.lbl_folder = ElidedLabel()
        self.lbl_folder.setObjectName("ws_folder")
        folder_row.addWidget(self.lbl_folder, 1)
        self.btn_folder = QPushButton("Сменить")
        self.btn_folder.setObjectName("ws_link")
        self.btn_folder.setCursor(Qt.PointingHandCursor)
        self.btn_folder.setToolTip("Выбрать другую папку с пунктами")
        self.btn_folder.clicked.connect(self.folder_requested.emit)
        folder_row.addWidget(self.btn_folder)
        title_col.addLayout(folder_row)
        head.addLayout(title_col, 1)

        seg = QHBoxLayout()
        seg.setSpacing(0)
        self.btn_tab_points = QPushButton("Пункты")
        self.btn_tab_report = QPushButton("Отчёт о прогоне")
        for name, b in (("left", self.btn_tab_points), ("right", self.btn_tab_report)):
            b.setObjectName("ws_tab")
            b.setProperty("side", name)
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            seg.addWidget(b)
        self.btn_tab_points.clicked.connect(lambda: self.show_tab("points"))
        self.btn_tab_report.clicked.connect(lambda: self.show_tab("report"))
        head.addLayout(seg)

        # ── Блок прогона ──
        run = QFrame()
        run.setObjectName("ws_run")
        run.setFixedWidth(320)
        rh = QHBoxLayout(run)
        rh.setContentsMargins(12, 8, 8, 8)
        rh.setSpacing(10)
        rv = QVBoxLayout()
        rv.setSpacing(5)
        top = QHBoxLayout()
        top.setSpacing(8)
        self.lbl_run = ElidedLabel()
        self.lbl_run.setObjectName("ws_run_title")
        top.addWidget(self.lbl_run, 1)
        self.lbl_timer = QLabel("")
        self.lbl_timer.setObjectName("ws_run_timer")
        top.addWidget(self.lbl_timer)
        rv.addLayout(top)
        self.run_progress = QProgressBar()
        self.run_progress.setObjectName("ws_run_progress")
        self.run_progress.setTextVisible(False)
        self.run_progress.setRange(0, 100)
        self.run_progress.setFixedHeight(4)
        rv.addWidget(self.run_progress)
        rh.addLayout(rv, 1)
        self.btn_stop = QPushButton("Стоп")
        self.btn_stop.setObjectName("ws_stop")
        self.btn_stop.setCursor(Qt.PointingHandCursor)
        self.btn_stop.setToolTip("Остановить текущий шаг")
        self.btn_stop.clicked.connect(self.stop_requested.emit)
        rh.addWidget(self.btn_stop)
        head.addWidget(run)
        return head

    def _build_points_tab(self):
        tab = QWidget()
        v = QVBoxLayout(tab)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(10)

        # ── Строка шагов ──
        strip = QHBoxLayout()
        strip.setSpacing(6)
        self.step_chips = {}
        for step in STEPS:
            chip = StepChip(step)
            self.step_chips[step] = chip
            strip.addWidget(chip, 1)
        v.addLayout(strip)

        # ── Панель запуска ──
        bar = QHBoxLayout()
        bar.setSpacing(8)
        self.input_points = QLineEdit()
        self.input_points.setObjectName("ws_input")
        self.input_points.setPlaceholderText("Номера пунктов: 2758, 2759 3186")
        self.input_points.setMinimumWidth(180)
        self.input_points.returnPressed.connect(self._on_add)
        bar.addWidget(self.input_points, 1)
        self.btn_add = QPushButton("Добавить")
        self.btn_add.setObjectName("ws_secondary")
        self.btn_add.setCursor(Qt.PointingHandCursor)
        self.btn_add.setToolTip("Создать папки пунктов и собрать комплекты (шаг 1)")
        self.btn_add.clicked.connect(self._on_add)
        bar.addWidget(self.btn_add)
        self.btn_refresh = QPushButton("Обновить")
        self.btn_refresh.setObjectName("ws_secondary")
        self.btn_refresh.setCursor(Qt.PointingHandCursor)
        self.btn_refresh.setToolTip("Перечитать состояние пунктов из папки")
        self.btn_refresh.clicked.connect(self.refresh_requested.emit)
        bar.addWidget(self.btn_refresh)
        self.btn_reset = QPushButton("Сбросить…")
        self.btn_reset.setObjectName("ws_secondary")
        self.btn_reset.setCursor(Qt.PointingHandCursor)
        self.btn_reset.setToolTip("Сбросить статус пункта или шага, чтобы выполнить его заново")
        self.btn_reset.clicked.connect(self.reset_requested.emit)
        bar.addWidget(self.btn_reset)
        bar.addSpacing(12)
        self.btn_run = QPushButton("")
        self.btn_run.setObjectName("ws_primary")
        self.btn_run.setCursor(Qt.PointingHandCursor)
        self.btn_run.clicked.connect(self._on_run_step)
        bar.addWidget(self.btn_run)
        self.btn_to_end = QPushButton("До конца")
        self.btn_to_end.setObjectName("ws_secondary")
        self.btn_to_end.setCursor(Qt.PointingHandCursor)
        self.btn_to_end.setToolTip(
            "Выполнить по очереди все оставшиеся шаги для отмеченных пунктов")
        self.btn_to_end.clicked.connect(self._on_run_to_end)
        bar.addWidget(self.btn_to_end)
        v.addLayout(bar)

        # ── Таблица / подсказка для пустой ──
        self.table_stack = QStackedWidget()
        self.table = self._build_table()
        self.table_stack.addWidget(self.table)
        empty = QFrame()
        empty.setObjectName("ws_empty")
        ev = QVBoxLayout(empty)
        ev.setContentsMargins(24, 24, 24, 24)
        ev.addStretch()
        e1 = QLabel("Пунктов пока нет")
        e1.setObjectName("ws_empty_title")
        e1.setAlignment(Qt.AlignCenter)
        ev.addWidget(e1)
        e2 = QLabel("Введите номера пунктов в поле выше и нажмите «Добавить» — "
                    "программа создаст папки и соберёт комплекты.\n"
                    "Если комплект уже сложен в папке ПНОС вручную, он будет "
                    "взят как есть.")
        e2.setObjectName("ws_empty_text")
        e2.setAlignment(Qt.AlignCenter)
        e2.setWordWrap(True)
        ev.addWidget(e2)
        ev.addStretch()
        self.table_stack.addWidget(empty)
        v.addWidget(self.table_stack, 1)

        # ── Легенда ──
        legend = QHBoxLayout()
        legend.setSpacing(14)
        for status, text in LEGEND:
            item = QHBoxLayout()
            item.setSpacing(5)
            dot = QLabel()
            dot.setFixedSize(DOT, DOT)
            dot.setPixmap(status_pixmap(status))
            item.addWidget(dot)
            lbl = QLabel(text)
            lbl.setObjectName("ws_legend")
            item.addWidget(lbl)
            legend.addLayout(item)
        legend.addStretch()
        hint = QLabel(f"Нажмите на точку шага {FORM_STEP}, чтобы открыть форму пункта")
        hint.setObjectName("ws_legend")
        legend.addWidget(hint)
        v.addLayout(legend)
        return tab

    def _build_table(self):
        t = QTableWidget(0, COL_STATE + 1)
        t.setObjectName("ws_table")
        t.setHorizontalHeaderLabels(
            ["", "Пункт"] + [str(s) for s in STEPS] + ["Состояние"])
        for step in STEPS:
            t.horizontalHeaderItem(COL_STEP0 + step - 1).setToolTip(
                f"Шаг {step} · {STEP_NAMES[step]}")
        t.horizontalHeaderItem(COL_POINT).setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        t.horizontalHeaderItem(COL_STATE).setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        t.setEditTriggers(QTableWidget.NoEditTriggers)
        t.setSelectionMode(QTableWidget.NoSelection)
        t.setFocusPolicy(Qt.NoFocus)
        t.setShowGrid(False)
        t.setAlternatingRowColors(True)
        t.setWordWrap(False)
        t.verticalHeader().setVisible(False)
        t.verticalHeader().setDefaultSectionSize(34)
        t.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        hh = t.horizontalHeader()
        hh.setHighlightSections(False)
        hh.setSectionResizeMode(QHeaderView.Fixed)
        hh.resizeSection(COL_CHECK, 40)
        hh.resizeSection(COL_POINT, 96)
        hh.setDefaultAlignment(Qt.AlignCenter)
        for step in STEPS:
            hh.resizeSection(COL_STEP0 + step - 1, 38)
        hh.setSectionResizeMode(COL_STATE, QHeaderView.Stretch)
        hh.setMinimumSectionSize(30)

        delegate = StepDotDelegate(t)
        for step in STEPS:
            t.setItemDelegateForColumn(COL_STEP0 + step - 1, delegate)

        # «Отметить все» — настоящий QCheckBox поверх заголовка колонки отметок:
        # так он оформлен тем же QSS, что и отметки строк
        self.chk_all = QCheckBox(hh)
        self.chk_all.setObjectName("ws_check_all")
        self.chk_all.setToolTip("Отметить все пункты")
        self.chk_all.setCursor(Qt.PointingHandCursor)
        self.chk_all.clicked.connect(self._on_check_all)
        hh.sectionResized.connect(lambda *_: self._place_check_all())
        hh.geometriesChanged.connect(self._place_check_all)

        t.itemChanged.connect(self._on_item_changed)
        t.cellClicked.connect(self._on_cell_clicked)
        t.setMouseTracking(True)
        t.viewport().installEventFilter(self)
        return t

    def _build_log(self):
        box = QFrame()
        box.setObjectName("ws_log")
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 6, 8, 6)
        v.setSpacing(6)

        bar = QHBoxLayout()
        bar.setSpacing(8)
        cap = QLabel("Журнал")
        cap.setObjectName("ws_log_caption")
        bar.addWidget(cap)
        self.log_dot = QLabel()
        self.log_dot.setFixedSize(8, 8)
        bar.addWidget(self.log_dot)
        self.lbl_log_last = ElidedLabel()
        self.lbl_log_last.setObjectName("ws_log_last")
        bar.addWidget(self.lbl_log_last, 1)
        self.lbl_log_count = QLabel("")
        self.lbl_log_count.setObjectName("ws_log_count")
        bar.addWidget(self.lbl_log_count)
        self.btn_log = QPushButton("Развернуть")
        self.btn_log.setObjectName("ws_link")
        self.btn_log.setCursor(Qt.PointingHandCursor)
        self.btn_log.clicked.connect(lambda: self.set_log_expanded(not self._log_expanded))
        bar.addWidget(self.btn_log)
        v.addLayout(bar)

        self.log_list = QListWidget()
        self.log_list.setObjectName("ws_log_list")
        self.log_list.setFixedHeight(self.LOG_EXPANDED_HEIGHT)
        self.log_list.setWordWrap(True)
        self.log_list.hide()
        v.addWidget(self.log_list)

        self._log_bar = box
        box.installEventFilter(self)
        self._set_last_log("Сообщений пока нет", "info", empty=True)
        return box

    # ═══════════════════════════════════════════════════════════════
    # API для окна
    # ═══════════════════════════════════════════════════════════════

    def set_points(self, rows: list[dict]):
        """Заполняет таблицу. Отметки сохраняются по номеру пункта."""
        self._rows = [dict(r) for r in (rows or [])]
        present = {str(r.get("point", "")) for r in self._rows}
        self._checked &= present

        t = self.table
        self._filling = True
        t.setUpdatesEnabled(False)
        try:
            t.clearContents()
            t.setRowCount(len(self._rows))
            for i, row in enumerate(self._rows):
                self._fill_row(i, row)
        finally:
            t.setUpdatesEnabled(True)
            self._filling = False

        self.table_stack.setCurrentIndex(0 if self._rows else 1)
        self._update_step_strip()
        self._update_check_all()
        self._update_buttons()

    def selected_points(self) -> list[str]:
        """Отмеченные номера в порядке таблицы."""
        return [str(r["point"]) for r in self._rows if str(r["point"]) in self._checked]

    def set_run_status(self, step, point, progress: int, elapsed_text: str):
        """Блок прогона в шапке; step=None — прогона нет."""
        self._run_step = step
        if step is None:
            self.lbl_run.setText("Прогон не идёт")
            self.lbl_timer.setText("")
            self.run_progress.setValue(0)
            self.lbl_run.setProperty("active", False)
        else:
            text = f"Шаг {step} · {STEP_NAMES.get(step, '')}"
            if point:
                text = f"Шаг {step} · п.{point}"
            self.lbl_run.setText(text)
            self.lbl_timer.setText(elapsed_text or "")
            self.run_progress.setValue(max(0, min(100, int(progress or 0))))
            self.lbl_run.setProperty("active", True)
        self.lbl_run.style().unpolish(self.lbl_run)
        self.lbl_run.style().polish(self.lbl_run)
        self._update_buttons()

    def add_log(self, step: int, message: str, category: str = "info"):
        """Строка журнала: точка категории + «Шаг N · текст» (step=0 — без шага)."""
        if category not in LOG_COLORS:
            category = "info"
        text = f"Шаг {step} · {message}" if step else str(message)
        item = QListWidgetItem(text)
        item.setIcon(dot_icon(LOG_COLORS[category]))
        self.log_list.addItem(item)
        self.log_list.scrollToBottom()
        self._log_count += 1
        self._set_last_log(text, category)

    def set_folder_info(self, path: str, count: int):
        if path:
            self.lbl_folder.setText(f"{path}  ·  {plural_points(count)}")
        else:
            self.lbl_folder.setText("Папка ПНОС не выбрана")

    def set_busy(self, busy: bool):
        """Во время прогона запуск новых шагов и добавление пунктов заблокированы."""
        self._busy = bool(busy)
        self._update_buttons()

    def show_tab(self, name: str):
        idx = 1 if name == "report" else 0
        self.tabs.setCurrentIndex(idx)
        self.btn_tab_points.setChecked(idx == 0)
        self.btn_tab_report.setChecked(idx == 1)

    def set_log_expanded(self, expanded: bool):
        self._log_expanded = bool(expanded)
        self.log_list.setVisible(expanded)
        self.btn_log.setText("Свернуть" if expanded else "Развернуть")
        self._show_last_log()
        if expanded:
            self.log_list.scrollToBottom()

    # ═══════════════════════════════════════════════════════════════
    # Таблица
    # ═══════════════════════════════════════════════════════════════

    def _fill_row(self, i: int, row: dict):
        t = self.table
        point = str(row.get("point", ""))
        steps = row.get("steps") or {}

        chk = QTableWidgetItem()
        chk.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
        chk.setCheckState(Qt.Checked if point in self._checked else Qt.Unchecked)
        chk.setData(ROLE_POINT, point)
        t.setItem(i, COL_CHECK, chk)

        it = QTableWidgetItem(f"п.{point}")
        it.setFlags(Qt.ItemIsEnabled)
        if row.get("folder"):
            it.setToolTip(str(row["folder"]))
        t.setItem(i, COL_POINT, it)

        for step in STEPS:
            status = steps.get(step) or steps.get(str(step)) or "idle"
            cell = QTableWidgetItem()
            cell.setFlags(Qt.ItemIsEnabled)
            cell.setData(ROLE_STATUS, status)
            tip = f"Шаг {step} · {STEP_NAMES[step]}: {STATUS_TEXT.get(status, status)}"
            if status == "error" and row.get("state"):
                tip += f" — {row['state']}"
            if step == FORM_STEP:
                tip += "\nНажмите, чтобы открыть форму"
            cell.setToolTip(tip)
            t.setItem(i, COL_STEP0 + step - 1, cell)

        # Текст состояния — надпись с QSS-цветом категории (токены темы),
        # а не цвет элемента: жёлтый на белом иначе не читается
        kind = row.get("state_kind") or "info"
        state = ElidedLabel(str(row.get("state", "")))
        state.setObjectName("ws_state")
        state.setProperty("kind", kind)
        state.setContentsMargins(3, 0, 8, 0)
        t.setItem(i, COL_STATE, QTableWidgetItem())
        t.item(i, COL_STATE).setFlags(Qt.ItemIsEnabled)
        t.setCellWidget(i, COL_STATE, state)

    def _row_steps(self, row: dict) -> dict:
        steps = row.get("steps") or {}
        return {s: (steps.get(s) or steps.get(str(s)) or "idle") for s in STEPS}

    def _update_step_strip(self):
        total = len(self._rows)
        for step, chip in self.step_chips.items():
            statuses = [self._row_steps(r)[step] for r in self._rows]
            counts = {s: statuses.count(s) for s in set(statuses)}
            chip.set_state(aggregate_status(statuses), counts.get("done", 0),
                           total, counts)

    def _on_item_changed(self, item):
        if self._filling or item.column() != COL_CHECK:
            return
        point = item.data(ROLE_POINT)
        if item.checkState() == Qt.Checked:
            self._checked.add(point)
        else:
            self._checked.discard(point)
        self._update_check_all()
        self._update_buttons()

    def _on_check_all(self):
        check = bool(self._rows) and len(self._checked) < len(self._rows)
        self._checked = {str(r["point"]) for r in self._rows} if check else set()
        self._filling = True
        for i in range(self.table.rowCount()):
            self.table.item(i, COL_CHECK).setCheckState(
                Qt.Checked if check else Qt.Unchecked)
        self._filling = False
        self._update_check_all()
        self._update_buttons()

    def _update_check_all(self):
        n, total = len(self._checked), len(self._rows)
        self.chk_all.setEnabled(bool(total))
        self.chk_all.setTristate(False)
        if total and n == total:
            self.chk_all.setCheckState(Qt.Checked)
        elif n:
            self.chk_all.setCheckState(Qt.PartiallyChecked)
        else:
            self.chk_all.setCheckState(Qt.Unchecked)
        self.chk_all.setToolTip("Снять все отметки" if n else "Отметить все пункты")

    def _place_check_all(self):
        hh = self.table.horizontalHeader()
        size = self.chk_all.sizeHint()
        x = hh.sectionViewportPosition(COL_CHECK) + (hh.sectionSize(COL_CHECK) - size.width()) // 2
        y = (hh.height() - size.height()) // 2
        self.chk_all.setGeometry(QRect(x, y, size.width(), size.height()))

    def showEvent(self, event):
        super().showEvent(event)
        self._place_check_all()

    def _on_cell_clicked(self, row: int, col: int):
        if col == COL_STEP0 + FORM_STEP - 1 and 0 <= row < len(self._rows):
            self.open_form_requested.emit(str(self._rows[row]["point"]))
        elif col in (COL_POINT, COL_STATE):
            # Клик по строке — то же, что по отметке: целиться в 18 px не нужно
            item = self.table.item(row, COL_CHECK)
            if item is not None:
                item.setCheckState(Qt.Unchecked if item.checkState() == Qt.Checked
                                   else Qt.Checked)

    def eventFilter(self, obj, event):
        if obj is self.table.viewport() and event.type() == QEvent.MouseMove:
            col = self.table.columnAt(event.position().toPoint().x())
            row = self.table.rowAt(event.position().toPoint().y())
            hand = row >= 0 and col in (COL_STEP0 + FORM_STEP - 1, COL_POINT, COL_STATE)
            obj.setCursor(Qt.PointingHandCursor if hand else Qt.ArrowCursor)
        elif obj is getattr(self, "_log_bar", None) and event.type() == QEvent.MouseButtonRelease:
            if event.position().y() < 30:
                self.set_log_expanded(not self._log_expanded)
                return True
        return super().eventFilter(obj, event)

    # ═══════════════════════════════════════════════════════════════
    # Кнопки
    # ═══════════════════════════════════════════════════════════════

    def _selected_next_step(self):
        """Минимальный следующий шаг среди отмеченных; None — все готовы."""
        steps = [next_step(self._row_steps(r)) for r in self._rows
                 if str(r["point"]) in self._checked]
        steps = [s for s in steps if s is not None]
        return min(steps) if steps else None

    def _update_buttons(self):
        n = len(self._checked)
        running = self._busy or self._run_step is not None
        nxt = self._selected_next_step()

        if not n:
            self.btn_run.setText("Отметьте пункты в таблице")
            self.btn_run.setToolTip("Шаг запускается для отмеченных пунктов")
        elif nxt is None:
            self.btn_run.setText("Отмеченные пункты готовы")
            self.btn_run.setToolTip("У отмеченных пунктов выполнены все семь шагов")
        else:
            self.btn_run.setText(f"Шаг {nxt} · {STEP_NAMES[nxt]} для выбранных ({n})")
            tip = (f"Самый ранний невыполненный шаг среди отмеченных: {nxt}. "
                   "Пункты, где он уже выполнен, будут пропущены.")
            if nxt == FORM_STEP:
                tip = "Открыть форму ручного контроля для отмеченных пунктов"
            self.btn_run.setToolTip(tip)
        self.btn_run.setEnabled(bool(n) and nxt is not None and not running)
        # Пунктирная рамка — «кнопка ждёт отметок», а не «запуск запрещён»
        hint = not n or nxt is None
        if self.btn_run.property("hint") != hint:
            self.btn_run.setProperty("hint", hint)
            self.btn_run.style().unpolish(self.btn_run)
            self.btn_run.style().polish(self.btn_run)
        self.btn_to_end.setEnabled(bool(n) and nxt is not None and not running)
        self.btn_add.setEnabled(not running)
        self.btn_reset.setEnabled(not running)
        self.btn_folder.setEnabled(not running)
        self.input_points.setEnabled(not running)
        self.btn_stop.setEnabled(running)

    def _on_add(self):
        if not self.btn_add.isEnabled():
            return
        nums = parse_point_numbers(self.input_points.text())
        if not nums:
            self.input_points.setFocus()
            return
        self.input_points.clear()
        self.add_points_requested.emit(nums)

    def _on_run_step(self):
        nxt = self._selected_next_step()
        if nxt is not None and not self._busy:
            self.run_step_requested.emit(nxt, self.selected_points())

    def _on_run_to_end(self):
        if self._checked and not self._busy:
            self.run_to_end_requested.emit(self.selected_points())

    # ═══════════════════════════════════════════════════════════════
    # Журнал
    # ═══════════════════════════════════════════════════════════════

    def _set_last_log(self, text: str, category: str, empty: bool = False):
        self._last_log = (text, empty)
        self.lbl_log_last.setProperty("empty", empty)
        self.lbl_log_last.style().unpolish(self.lbl_log_last)
        self.lbl_log_last.style().polish(self.lbl_log_last)
        self.log_dot.setPixmap(dot_icon(LOG_COLORS[category], 8).pixmap(8, 8))
        self.lbl_log_count.setText(f"{self._log_count} зап." if self._log_count else "")
        self._show_last_log()

    def _show_last_log(self):
        """Последнее сообщение — только в свёрнутом журнале: раскрытый список
        и так показывает его последней строкой. Надпись не прячется, а
        пустеет, чтобы счётчик и кнопка остались прижаты вправо."""
        text, empty = self._last_log
        expanded = self._log_expanded
        self.lbl_log_last.setText("" if expanded else text)
        self.log_dot.setVisible(not empty and not expanded)
