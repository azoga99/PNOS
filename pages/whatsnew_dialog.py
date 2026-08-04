# -*- coding: utf-8 -*-
"""
Окно «Что нового» — журнал изменений программы.

Современный вид: временная шкала слева, карточки версий, цветные теги
категорий (Новое / Улучшено / Исправлено / Изменено), живой поиск по
изменениям. Данные берутся из changelog.py (разбор CHANGELOG.md).

Открывается из Настроек (весь журнал) и один раз автоматически после
обновления (только новые версии) — см. main._show_whats_new_if_updated.
"""

from __future__ import annotations

from PySide6.QtCore import Qt, QPoint, QSettings
from PySide6.QtGui import QPainter, QPen, QColor
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QScrollArea, QFrame
)

import changelog
import theme

# Цвет буллета/акцента по коду категории (токен палитры).
_CAT_TOKEN = {
    "new": "success",
    "improved": "info_blue",
    "fixed": "warn",
    "changed": "text_muted",
}


class _TimelineRail(QWidget):
    """Вертикальная шкала слева от карточки: линия + узловая точка."""

    DOT_Y = 26
    WIDTH = 36

    def __init__(self, line_color: str, dot_color: str, bg_color: str,
                 first: bool, last: bool, parent=None):
        super().__init__(parent)
        self.setFixedWidth(self.WIDTH)
        self._line = QColor(line_color)
        self._dot = QColor(dot_color)
        self._bg = QColor(bg_color)
        self._first = first
        self._last = last

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        cx = self.width() // 2

        top = self.DOT_Y if self._first else 0
        bottom = self.DOT_Y if self._last else self.height()
        if bottom > top:
            p.setPen(QPen(self._line, 2))
            p.drawLine(cx, top, cx, bottom)

        p.setPen(Qt.NoPen)
        # Кольцо фоном, чтобы точка «разрывала» линию.
        p.setBrush(self._bg)
        p.drawEllipse(QPoint(cx, self.DOT_Y), 8, 8)
        p.setBrush(self._dot)
        p.drawEllipse(QPoint(cx, self.DOT_Y), 5, 5)
        p.end()


class WhatsNewDialog(QDialog):
    """Диалог со списком изменений по версиям."""

    def __init__(self, entries: list[dict], title: str = "Что нового",
                 subtitle: str = "", parent=None):
        super().__init__(parent)
        self.setObjectName("whatsnew")
        self.setWindowTitle(title)
        self.setMinimumSize(620, 660)

        self._theme = QSettings("PNOS", "PNOS").value("theme", "light")
        self._entries = entries
        self._rows: list[dict] = []   # состояние для фильтрации поиском
        self._setup_ui(title, subtitle)

    # ── Построение UI ────────────────────────────────────────────
    def _c(self, token: str) -> str:
        return theme.color(self._theme, token)

    def _setup_ui(self, title: str, subtitle: str):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Шапка ────────────────────────────────────────────────
        header = QWidget()
        header.setObjectName("cl_header")
        h = QVBoxLayout(header)
        h.setContentsMargins(26, 22, 26, 18)
        h.setSpacing(4)

        lbl_title = QLabel(title)
        lbl_title.setObjectName("cl_head_title")
        h.addWidget(lbl_title)

        if subtitle:
            lbl_sub = QLabel(subtitle)
            lbl_sub.setObjectName("cl_head_sub")
            lbl_sub.setWordWrap(True)
            h.addWidget(lbl_sub)

        if len(self._entries) >= 4:
            self._search = QLineEdit()
            self._search.setObjectName("cl_search")
            self._search.setPlaceholderText("Поиск по изменениям…")
            self._search.setClearButtonEnabled(True)
            self._search.textChanged.connect(self._apply_filter)
            h.addSpacing(8)
            h.addWidget(self._search)
        else:
            self._search = None

        root.addWidget(header)

        # ── Лента версий ─────────────────────────────────────────
        scroll = QScrollArea()
        scroll.setObjectName("cl_scroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        body = QWidget()
        body.setObjectName("cl_body")
        self._body_layout = QVBoxLayout(body)
        self._body_layout.setContentsMargins(24, 18, 24, 24)
        self._body_layout.setSpacing(0)

        n = len(self._entries)
        for i, entry in enumerate(self._entries):
            self._body_layout.addWidget(
                self._build_entry(entry, first=(i == 0), last=(i == n - 1)))

        self._empty = QLabel("Ничего не найдено.")
        self._empty.setObjectName("cl_empty")
        self._empty.setAlignment(Qt.AlignCenter)
        self._empty.setVisible(False)
        self._body_layout.addWidget(self._empty)

        self._body_layout.addStretch()
        scroll.setWidget(body)
        root.addWidget(scroll, 1)

        # ── Подвал ───────────────────────────────────────────────
        footer = QWidget()
        footer.setObjectName("cl_footer")
        f = QHBoxLayout(footer)
        f.setContentsMargins(26, 12, 26, 16)
        f.addStretch()
        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("blue_button")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.clicked.connect(self.accept)
        f.addWidget(btn_close)
        root.addWidget(footer)

    def _build_entry(self, entry: dict, first: bool, last: bool) -> QWidget:
        row = QWidget()
        rl = QHBoxLayout(row)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(4)

        # Точка шкалы окрашивается по «главной» категории версии.
        main_code = entry["groups"][0]["code"] if entry["groups"] else "changed"
        rail = _TimelineRail(
            line_color=self._c("border_strong"),
            dot_color=self._c(_CAT_TOKEN.get(main_code, "text_muted")),
            bg_color=self._c("bg_content"),
            first=first, last=last)
        rl.addWidget(rail, 0, Qt.AlignTop)

        card = QFrame()
        card.setObjectName("cl_card")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(18, 14, 18, 14)
        cl.setSpacing(10)

        # Заголовок карточки: версия + дата.
        head = QHBoxLayout()
        head.setSpacing(10)
        pill = QLabel(f"Версия {entry['version']}")
        pill.setObjectName("cl_version")
        head.addWidget(pill)
        if entry.get("date"):
            date = QLabel(entry["date"])
            date.setObjectName("cl_date")
            head.addWidget(date)
        head.addStretch()
        cl.addLayout(head)

        item_states = []
        for g in entry["groups"]:
            group_w = QWidget()
            gl = QVBoxLayout(group_w)
            gl.setContentsMargins(0, 0, 0, 0)
            gl.setSpacing(6)

            chip = QLabel(g["label"].upper())
            chip.setObjectName("cl_chip")
            chip.setProperty("cat", g["code"])
            gl.addWidget(chip, 0, Qt.AlignLeft)

            for text in g["items"]:
                item_w = self._build_item(text, g["code"])
                gl.addWidget(item_w)
                item_states.append((item_w, text.lower()))

            cl.addWidget(group_w)
            item_states.append(("__group__", group_w))

        rl.addWidget(card, 1)

        self._rows.append({
            "row": row,
            "meta": f"{entry['version']} {entry.get('date','')}".lower(),
            "items": item_states,
        })
        return row

    def _build_item(self, text: str, code: str) -> QWidget:
        w = QWidget()
        hl = QHBoxLayout(w)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(8)

        bullet = QLabel("•")
        bullet.setObjectName("cl_bullet")
        bullet.setStyleSheet(
            f"color: {self._c(_CAT_TOKEN.get(code, 'text_muted'))}; "
            f"font-size: 15px; font-weight: 700;")
        hl.addWidget(bullet, 0, Qt.AlignTop)

        lbl = QLabel(text)
        lbl.setObjectName("cl_item")
        lbl.setWordWrap(True)
        hl.addWidget(lbl, 1)
        return w

    # ── Поиск ────────────────────────────────────────────────────
    def _apply_filter(self, query: str):
        q = query.lower().strip()
        any_visible = False

        for row in self._rows:
            meta_hit = bool(q) and q in row["meta"]
            row_visible = False
            pending_group = None
            group_has = False

            def close_group():
                if pending_group is not None:
                    pending_group.setVisible(group_has if q else True)

            for item, payload in row["items"]:
                if item == "__group__":
                    close_group()
                    pending_group = payload
                    group_has = False
                    continue
                hit = (not q) or meta_hit or (q in payload)
                item.setVisible(hit)
                if hit:
                    group_has = True
                    row_visible = True
            close_group()

            visible = row_visible if q else True
            row["row"].setVisible(visible)
            any_visible = any_visible or visible

        self._empty.setVisible(bool(q) and not any_visible)

    # ── Точки входа ──────────────────────────────────────────────
    @staticmethod
    def show_full(parent=None):
        """Показать весь журнал изменений (кнопка в Настройках)."""
        entries = changelog.load_entries()
        dlg = WhatsNewDialog(
            entries, title="Журнал изменений",
            subtitle="История версий программы ПНОС", parent=parent)
        dlg.exec()

    @staticmethod
    def show_since(version: str, parent=None) -> bool:
        """Показать изменения новее `version` (автопоказ после обновления).

        Возвращает True, если было что показать.
        """
        entries = changelog.entries_since(version)
        if not entries:
            return False
        dlg = WhatsNewDialog(
            entries, title="Что нового",
            subtitle="Программа обновлена. Вот что изменилось:", parent=parent)
        dlg.exec()
        return True
