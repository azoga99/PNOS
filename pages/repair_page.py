# -*- coding: utf-8 -*-
"""
Страница «Ремонт» — автономный модуль (ветка repair-module, этап Э1).

Поток: выбор папки пункта → «Проверить пункт» (фоновое сканирование: детекция
документов, светофор привязки, реестр, сечения) → правка реестра → отметка
заменённых сечений (колонка «Новый номинал») → предпросмотр «было → станет»
с галочками → «Применить» (бэкап → Excel COM → Word → журнал; при ошибке
автооткат) / «Выгрузить чек-лист» (ручной путь) / «Откатить последнюю».

Страница ничего не пишет без предпросмотра и явной кнопки (принцип П-2);
все данные живут в обычных файлах пункта (П-1). Тяжёлые модули импортируются
лениво — быстрый старт приложения не страдает.
"""

from __future__ import annotations

import os

from PySide6.QtCore import Qt, Signal, QPointF
from PySide6.QtGui import QImage, QPixmap, QIcon
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QFrame, QLabel,
    QLineEdit, QPushButton, QFileDialog, QListWidget, QListWidgetItem,
    QTableWidget, QTableWidgetItem, QTabWidget, QHeaderView, QAbstractItemView,
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem, QTextEdit, QCheckBox,
)

from widgets import dot_icon


class MapView(QGraphicsView):
    """Просмотр коррозионной карты: зум колесом, панорама перетаскиванием,
    клик (без движения мыши) отдаёт координаты пикселя изображения."""

    clicked = Signal(float, float)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._scene = QGraphicsScene(self)
        self.setScene(self._scene)
        self._pix_item: QGraphicsPixmapItem | None = None
        self._press_pos = None
        self._fitted = False
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)

    def set_image(self, qimage: QImage):
        self._scene.clear()
        self._pix_item = self._scene.addPixmap(QPixmap.fromImage(qimage))
        self._scene.setSceneRect(self._pix_item.boundingRect())
        self._marks = []
        if not self._fitted:
            self.fitInView(self._pix_item, Qt.KeepAspectRatio)
            self._fitted = True

    def show_at(self, x: float, y: float, box=None, zoom: float = 3.0):
        """
        Подвести взгляд к точке: рамка вокруг подписи и центрирование.

        Это и убирает шаг «найти сечение на схеме»: программа знает координаты
        подписи и сама показывает нужное место.
        """
        from PySide6.QtGui import QPen, QColor
        from PySide6.QtCore import QRectF

        if self._pix_item is None:
            return
        for it in getattr(self, "_marks", []):
            self._scene.removeItem(it)
        self._marks = []
        if box is not None:
            pad = 6
            rect = QRectF(box.x1 - pad, box.y1 - pad,
                          (box.x2 - box.x1) + 2 * pad,
                          (box.y2 - box.y1) + 2 * pad)
            pen = QPen(QColor(0, 170, 90))
            pen.setWidth(3)
            pen.setCosmetic(True)
            self._marks.append(self._scene.addRect(rect, pen))
        self.resetTransform()
        self.scale(zoom, zoom)
        self.centerOn(x, y)

    def wheelEvent(self, event):
        factor = 1.2 if event.angleDelta().y() > 0 else 1 / 1.2
        self.scale(factor, factor)

    def mousePressEvent(self, event):
        self._press_pos = event.position()
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event)
        if self._press_pos is None or self._pix_item is None:
            return
        moved = (event.position() - self._press_pos).manhattanLength()
        self._press_pos = None
        if moved > 6:
            return                      # это была панорама, не клик
        scene_pos: QPointF = self.mapToScene(event.position().toPoint())
        if self._pix_item.boundingRect().contains(scene_pos):
            self.clicked.emit(scene_pos.x(), scene_pos.y())

_DOT = {
    "info": "#4A90D9", "done": "#4CAF50", "wait": "#FF9800",
    "warn": "#FF9800", "error": "#F44336",
}
_VERDICT = {"ok": ("done", "совпадает с пунктом"),
            "mismatch": ("error", "похоже, документ другого пункта"),
            "unknown": ("warn", "в имени данных нет — проверьте по превью")}

CHECKLIST_NAME = "_чек-лист_ремонта.txt"


class RepairPage(QWidget):
    """Автономная страница модуля «Ремонт»."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._workers = {}
        self._scan: dict = {}          # результат RepairScanWorker
        self._preview = None           # (ops_by_change, changes, svedenia)
        # строка «Сведения о ремонте»: собранная программой и правка сотрудника
        self._sved_generated = ""
        self._sved_edited = False
        self._sved_programmatic = False
        # состояние вкладки «Карта»
        self._map_path: str | None = None
        self._map_orig = None          # numpy-массив исходной карты
        self._map_arr = None           # numpy-массив с правками
        self._map_edits: list = []     # [(label_xy, donor_xy, описание)]
        self._map_pending = None       # первый клик пары (label_xy)
        self._map_labels: list = []    # подписи сечений, найденные на карте
        self._map_by_number: dict = {} # номер сечения → список подписей
        self._map_shown = None         # (номер, новое значение, подпись) на экране
        self._map_cycle: dict = {}     # номер → какое из мест показано
        # состояние превью документа (вкладка 1)
        self._doc_path: str | None = None
        self._doc_page_idx = 0
        self._doc_pages = 0
        self._id_map = None            # карта разделов комплекта ИД
        self._build_ui()

    # ═══════════════════════════════════ UI ═══════════════════════════

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(10)

        self._header = QWidget()
        head = QVBoxLayout(self._header)
        head.setContentsMargins(0, 0, 0, 0)
        head.setSpacing(6)
        title = QLabel("Ремонт")
        title.setObjectName("app_title")
        head.addWidget(title)
        sub = QLabel(
            "Учёт ремонтов трубопровода: акты отбраковки и исполнительная "
            "документация → лист «Ремонт», номиналы заменённых сечений на "
            "листах «Импорт», строка «Сведения о ремонте» в отчёте. "
            "Ничего не записывается без предпросмотра; перед каждой записью "
            "создаётся бэкап, доступен откат.")
        sub.setObjectName("hint")
        sub.setWordWrap(True)
        head.addWidget(sub)
        outer.addWidget(self._header)

        # ── Выбор пункта ──
        panel = QFrame()
        panel.setObjectName("settings_panel")
        g = QGridLayout(panel)
        g.setContentsMargins(14, 12, 14, 12)
        g.setHorizontalSpacing(10)
        g.addWidget(QLabel("Папка пункта"), 0, 0)
        self.edit_folder = QLineEdit()
        self.edit_folder.setPlaceholderText("Папка пункта п.NNNN")
        g.addWidget(self.edit_folder, 0, 1)
        btn_browse = QPushButton("📂")
        btn_browse.setObjectName("browse_btn")
        btn_browse.setFixedWidth(40)
        btn_browse.setCursor(Qt.PointingHandCursor)
        btn_browse.clicked.connect(self._browse)
        g.addWidget(btn_browse, 0, 2)
        self.btn_scan = QPushButton("Проверить пункт")
        self.btn_scan.setObjectName("green_button")
        self.btn_scan.setCursor(Qt.PointingHandCursor)
        self.btn_scan.clicked.connect(self._start_scan)
        g.addWidget(self.btn_scan, 0, 3)
        self.btn_verify = QPushButton("Верификация")
        self.btn_verify.setObjectName("segment")
        self.btn_verify.setCursor(Qt.PointingHandCursor)
        self.btn_verify.setToolTip(
            "Проверка согласованности пункта (реестр ↔ отчёт, зеркальность "
            "листов «Импорт», #REF!, задвижки в таблице 3). Работает для "
            "любого пункта — и сделанного вручную.")
        self.btn_verify.clicked.connect(self._verify)
        g.addWidget(self.btn_verify, 0, 4)
        self.btn_rollback = QPushButton("Откатить последнюю")
        self.btn_rollback.setObjectName("segment")
        self.btn_rollback.setCursor(Qt.PointingHandCursor)
        self.btn_rollback.clicked.connect(self._rollback)
        g.addWidget(self.btn_rollback, 0, 5)
        self._panel = panel
        outer.addWidget(panel)

        # ── Вкладки ──
        self.tabs = QTabWidget()
        outer.addWidget(self.tabs, 1)

        # Вкладка 1: документы и реестр
        tab_docs = QWidget()
        v1 = QVBoxLayout(tab_docs)
        v1.setSpacing(6)
        lbl_docs = QLabel("Документы ремонта — клик по документу открывает "
                          "превью страниц справа")
        lbl_docs.setObjectName("section_label")
        v1.addWidget(lbl_docs)

        docs_row = QHBoxLayout()
        docs_row.setSpacing(8)
        left_col = QVBoxLayout()
        self.list_docs = QListWidget()
        self.list_docs.setObjectName("activity_feed")
        self.list_docs.currentItemChanged.connect(self._doc_selected)
        left_col.addWidget(self.list_docs, 1)
        docs_row.addLayout(left_col, 1)

        prev_col = QVBoxLayout()
        prev_col.setSpacing(4)
        self.doc_view = MapView()
        self.doc_view.setMinimumWidth(300)
        prev_col.addWidget(self.doc_view, 1)
        nav = QHBoxLayout()
        self.btn_pg_prev = QPushButton("◀")
        self.btn_pg_prev.setObjectName("segment")
        self.btn_pg_prev.setFixedWidth(36)
        self.btn_pg_prev.clicked.connect(lambda: self._doc_page(-1))
        nav.addWidget(self.btn_pg_prev)
        self.lbl_pg = QLabel("—")
        self.lbl_pg.setObjectName("hint")
        self.lbl_pg.setAlignment(Qt.AlignCenter)
        nav.addWidget(self.lbl_pg, 1)
        self.btn_pg_next = QPushButton("▶")
        self.btn_pg_next.setObjectName("segment")
        self.btn_pg_next.setFixedWidth(36)
        self.btn_pg_next.clicked.connect(lambda: self._doc_page(1))
        nav.addWidget(self.btn_pg_next)
        btn_open = QPushButton("Открыть файл")
        btn_open.setObjectName("segment")
        btn_open.setCursor(Qt.PointingHandCursor)
        btn_open.clicked.connect(self._doc_open_external)
        nav.addWidget(btn_open)
        self.btn_ocr = QPushButton("Распознать (OCR)")
        self.btn_ocr.setObjectName("segment")
        self.btn_ocr.setCursor(Qt.PointingHandCursor)
        self.btn_ocr.setToolTip(
            "Windows OCR по первым страницам документа: номера и даты актов, "
            "материал, типоразмеры. Заполняются только пустые ячейки реестра — "
            "введённое вами не перезаписывается. Рукописные даты (день в "
            "грифе) распознаются плохо — их вписывайте по превью.")
        self.btn_ocr.clicked.connect(self._doc_ocr)
        nav.addWidget(self.btn_ocr)
        prev_col.addLayout(nav)
        docs_row.addLayout(prev_col, 1)
        v1.addLayout(docs_row, 1)
        lbl_reg = QLabel("Реестр ремонтов (лист «Ремонт»; запись = комплект ИД)")
        lbl_reg.setObjectName("section_label")
        v1.addWidget(lbl_reg)
        self.tbl_registry = QTableWidget(0, 8)
        self.tbl_registry.setHorizontalHeaderLabels(
            ["Статус ИД", "Дата выполнения", "Номера актов", "Даты актов",
             "Замена", "Материал", "Файлы", "Примечание"])
        self.tbl_registry.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.tbl_registry.horizontalHeader().setStretchLastSection(True)
        self.tbl_registry.verticalHeader().setVisible(False)
        # Высота строки с запасом: в неё встаёт редактор ячейки, иначе текст
        # при вводе обрезается по высоте.
        self.tbl_registry.verticalHeader().setDefaultSectionSize(30)
        v1.addWidget(self.tbl_registry, 1)
        row1 = QHBoxLayout()
        for text, slot in (("Добавить запись", self._reg_add),
                           ("Удалить запись", self._reg_del),
                           ("Заполнить из документов", self._reg_prefill)):
            b = QPushButton(text)
            b.setObjectName("segment")
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(slot)
            row1.addWidget(b)
        row1.addStretch()
        hint1 = QLabel("Статус ИД: есть / ожидается / не будет. Даты — ДД.ММ.ГГГГ; "
                       "дата акта — из грифа УТВЕРЖДАЮ, не с титульника.")
        hint1.setObjectName("hint")
        row1.addWidget(hint1)
        v1.addLayout(row1)
        self.tabs.addTab(tab_docs, "Документы и реестр")

        # Вкладка 2: сечения
        tab_sec = QWidget()
        v2 = QVBoxLayout(tab_sec)
        v2.setSpacing(6)
        hint2 = QLabel(
            "Отметьте заменённые сечения: впишите номинал новой трубы в колонку "
            "«Новый номинал» (например 10 или 8.0; для перехода — «12 / 10»). "
            "Пустая ячейка = сечение не менялось. Задвижки и бобышки по правилам "
            "не обновляются — при необходимости впишите номинал явно.")
        hint2.setObjectName("hint")
        hint2.setWordWrap(True)
        v2.addWidget(hint2)
        self.tbl_sections = QTableWidget(0, 7)
        self.tbl_sections.setHorizontalHeaderLabels(
            ["Сечение", "Элемент", "Тисп", "Диам", "Замеры", "Мин",
             "Новый номинал"])
        self.tbl_sections.verticalHeader().setVisible(False)
        self.tbl_sections.verticalHeader().setDefaultSectionSize(30)
        self.tbl_sections.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tbl_sections.horizontalHeader().setStretchLastSection(True)

        # Таблица слева, схема дефектного участка из акта — справа: на схеме
        # заменённый элемент выделен красным, и сверять кандидатов удобнее,
        # не выходя из программы.
        sec_row = QHBoxLayout()
        sec_row.setSpacing(8)
        sec_row.addWidget(self.tbl_sections, 3)
        self.scheme_view = MapView()
        self.scheme_view.setMinimumWidth(280)
        self.scheme_view.setVisible(False)
        sec_row.addWidget(self.scheme_view, 2)
        v2.addLayout(sec_row, 1)

        row2 = QHBoxLayout()
        self.chk_hide_small = QCheckBox("Скрыть задвижки и бобышки")
        self.chk_hide_small.setChecked(True)
        self.chk_hide_small.setToolTip(
            "По правилам их номиналы не обновляются. Снимите галочку, если "
            "нужно вписать номинал такому элементу явно.")
        self.chk_hide_small.toggled.connect(self._apply_sections_filter)
        row2.addWidget(self.chk_hide_small)

        self.btn_scheme = QPushButton("Схема из акта")
        self.btn_scheme.setObjectName("segment")
        self.btn_scheme.setCursor(Qt.PointingHandCursor)
        self.btn_scheme.setCheckable(True)
        self.btn_scheme.setToolTip(
            "Показать лист акта со схемой дефектного участка: заменяемый "
            "элемент выделен на ней красным")
        self.btn_scheme.toggled.connect(self._toggle_scheme)
        row2.addWidget(self.btn_scheme)

        self.btn_suggest = QPushButton("Подобрать по документам")
        self.btn_suggest.setObjectName("segment")
        self.btn_suggest.setCursor(Qt.PointingHandCursor)
        self.btn_suggest.setToolTip(
            "Сопоставить строки таблицы замены из документов с сечениями по "
            "типу элемента и диаметру. Если подходит несколько сечений, все "
            "они помечаются — выбрать место поможет карта.")
        self.btn_suggest.clicked.connect(self._suggest_sections)
        row2.addWidget(self.btn_suggest)
        row2.addStretch()
        hint_sug = QLabel("Зелёная точка — сечение одно; оранжевая — подходящих "
                          "несколько, номинал не подставляется")
        hint_sug.setObjectName("hint")
        row2.addWidget(hint_sug)
        v2.addLayout(row2)
        self.tabs.addTab(tab_sec, "Сечения")

        # Вкладка 3: предпросмотр
        tab_prev = QWidget()
        v3 = QVBoxLayout(tab_prev)
        v3.setSpacing(6)
        # Строка «Сведения о ремонте» — главный результат работы модуля,
        # поэтому она показывается целиком и её можно поправить руками до
        # записи: формулировки в отчётах набивались годами, и последнее слово
        # за экспертом.
        lbl_sv_title = QLabel("Строка «Сведения о ремонте» — попадёт в таблицу 1 "
                              "отчёта. Можно править перед записью.")
        lbl_sv_title.setObjectName("section_label")
        lbl_sv_title.setWordWrap(True)
        v3.addWidget(lbl_sv_title)

        lbl_now = QLabel("Сейчас в отчёте")
        lbl_now.setObjectName("hint")
        v3.addWidget(lbl_now)
        self.txt_sved_now = QTextEdit()
        self.txt_sved_now.setReadOnly(True)
        self.txt_sved_now.setFixedHeight(58)
        v3.addWidget(self.txt_sved_now)

        self.lbl_sved = QLabel("Станет")
        self.lbl_sved.setObjectName("hint")
        v3.addWidget(self.lbl_sved)
        self.txt_sved = QTextEdit()
        self.txt_sved.setFixedHeight(86)
        self.txt_sved.setPlaceholderText(
            "Строка появится после проверки пункта")
        self.txt_sved.textChanged.connect(self._on_sved_edited)
        v3.addWidget(self.txt_sved)

        row_sv = QHBoxLayout()
        self.btn_sved_reset = QPushButton("Вернуть сгенерированную")
        self.btn_sved_reset.setObjectName("segment")
        self.btn_sved_reset.setCursor(Qt.PointingHandCursor)
        self.btn_sved_reset.setToolTip(
            "Восстановить строку, собранную программой из реестра ремонтов")
        self.btn_sved_reset.clicked.connect(self._reset_svedenia)
        row_sv.addWidget(self.btn_sved_reset)
        row_sv.addStretch()
        v3.addLayout(row_sv)

        self.tbl_preview = QTableWidget(0, 4)
        self.tbl_preview.setHorizontalHeaderLabels(["✓", "Где", "Было", "Станет"])
        self.tbl_preview.verticalHeader().setVisible(False)
        self.tbl_preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # Значения бывают длинными — переносим по словам и тянем высоту строк,
        # иначе «Станет» показывается обрезанным с многоточием.
        self.tbl_preview.setWordWrap(True)
        hdr = self.tbl_preview.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.setSectionResizeMode(3, QHeaderView.Stretch)
        v3.addWidget(self.tbl_preview, 1)
        row3 = QHBoxLayout()
        self.btn_apply = QPushButton("Применить выбранное")
        self.btn_apply.setObjectName("green_button")
        self.btn_apply.setCursor(Qt.PointingHandCursor)
        self.btn_apply.setEnabled(False)
        self.btn_apply.clicked.connect(self._apply)
        row3.addWidget(self.btn_apply)
        self.btn_checklist = QPushButton("Выгрузить чек-лист")
        self.btn_checklist.setObjectName("segment")
        self.btn_checklist.setCursor(Qt.PointingHandCursor)
        self.btn_checklist.setEnabled(False)
        self.btn_checklist.clicked.connect(self._export_checklist)
        row3.addWidget(self.btn_checklist)
        row3.addStretch()
        v3.addLayout(row3)
        self.tabs.addTab(tab_prev, "Предпросмотр")

        # Вкладка 4: карта
        tab_map = QWidget()
        v4 = QHBoxLayout(tab_map)
        v4.setSpacing(8)
        self.map_view = MapView()
        self.map_view.clicked.connect(self._on_map_clicked)
        v4.addWidget(self.map_view, 1)

        side = QVBoxLayout()
        side.setSpacing(6)
        self.lbl_map_status = QLabel("Карта не загружена")
        self.lbl_map_status.setObjectName("form_field_label")
        self.lbl_map_status.setWordWrap(True)
        self.lbl_map_status.setMaximumWidth(260)
        side.addWidget(self.lbl_map_status)
        hint4 = QLabel(
            "Кликните по подписи заменённого сечения и введите новое значение. "
            "Шрифт Arial подбирается под размер подписи, цвет берётся из неё же "
            "(сечение выделено оранжевым — останется оранжевым). Линии схемы "
            "не страдают. Колесо — зум, перетаскивание — панорама.")
        hint4.setObjectName("hint")
        hint4.setWordWrap(True)
        hint4.setMaximumWidth(260)
        side.addWidget(hint4)

        self.btn_map_load = QPushButton("Перезагрузить карту")
        self.btn_map_load.setObjectName("segment")
        self.btn_map_load.setCursor(Qt.PointingHandCursor)
        self.btn_map_load.clicked.connect(self._map_load)
        side.addWidget(self.btn_map_load)

        lbl_ba = QLabel("До / после")
        lbl_ba.setObjectName("section_label")
        side.addWidget(lbl_ba)
        self.lbl_map_before = QLabel("—")
        self.lbl_map_before.setMaximumWidth(260)
        side.addWidget(self.lbl_map_before)
        self.lbl_map_after = QLabel("—")
        self.lbl_map_after.setMaximumWidth(260)
        side.addWidget(self.lbl_map_after)

        # Правки, которые предстоит внести: берутся из колонки «Новый номинал»
        # вкладки «Сечения». Программа сама находит подпись на карте, поэтому
        # искать сечение глазами не нужно.
        lbl_plan = QLabel("Правки по сечениям")
        lbl_plan.setObjectName("section_label")
        side.addWidget(lbl_plan)
        self.list_map_plan = QListWidget()
        self.list_map_plan.setObjectName("activity_feed")
        self.list_map_plan.setMaximumWidth(260)
        self.list_map_plan.setToolTip(
            "Клик — перейти к подписи на карте. Если сечений с таким номером "
            "несколько, повторный клик показывает следующее место.")
        self.list_map_plan.itemClicked.connect(self._map_plan_clicked)
        side.addWidget(self.list_map_plan, 1)

        row_plan = QHBoxLayout()
        self.btn_map_apply_one = QPushButton("Проставить показанное")
        self.btn_map_apply_one.setObjectName("segment")
        self.btn_map_apply_one.setCursor(Qt.PointingHandCursor)
        self.btn_map_apply_one.setEnabled(False)
        self.btn_map_apply_one.setToolTip(
            "Записать новое значение в подпись, к которой перешли")
        self.btn_map_apply_one.clicked.connect(self._map_apply_shown)
        row_plan.addWidget(self.btn_map_apply_one)
        side.addLayout(row_plan)
        self.btn_map_apply_unique = QPushButton("Проставить все однозначные")
        self.btn_map_apply_unique.setObjectName("segment")
        self.btn_map_apply_unique.setCursor(Qt.PointingHandCursor)
        self.btn_map_apply_unique.setEnabled(False)
        self.btn_map_apply_unique.setToolTip(
            "Проставить значения там, где подпись сечения на карте одна. "
            "Места с несколькими подписями остаются на ваше решение.")
        self.btn_map_apply_unique.clicked.connect(self._map_apply_unique)
        side.addWidget(self.btn_map_apply_unique)

        lbl_done = QLabel("Внесённые правки")
        lbl_done.setObjectName("section_label")
        side.addWidget(lbl_done)
        self.list_map_edits = QListWidget()
        self.list_map_edits.setObjectName("activity_feed")
        self.list_map_edits.setMaximumWidth(260)
        self.list_map_edits.setMaximumHeight(110)
        side.addWidget(self.list_map_edits)

        self.btn_map_undo = QPushButton("Отменить последнюю правку")
        self.btn_map_undo.setObjectName("segment")
        self.btn_map_undo.setCursor(Qt.PointingHandCursor)
        self.btn_map_undo.setEnabled(False)
        self.btn_map_undo.clicked.connect(self._map_undo)
        side.addWidget(self.btn_map_undo)
        self.btn_map_save = QPushButton("Сохранить карту (с бэкапом)")
        self.btn_map_save.setObjectName("green_button")
        self.btn_map_save.setCursor(Qt.PointingHandCursor)
        self.btn_map_save.setEnabled(False)
        self.btn_map_save.clicked.connect(self._map_save)
        side.addWidget(self.btn_map_save)
        v4.addLayout(side)
        self.tabs.addTab(tab_map, "Карта")

        self.tabs.currentChanged.connect(self._on_tab_changed)

        # ── Журнал страницы ──
        lbl_log = QLabel("Журнал")
        lbl_log.setObjectName("section_label")
        outer.addWidget(lbl_log)
        self.log_list = QListWidget()
        self.log_list.setObjectName("activity_feed")
        self.log_list.setMaximumHeight(110)
        outer.addWidget(self.log_list)

    # ═════════════════════════════ Помощники ═══════════════════════════

    def _log(self, message: str, category: str = "info"):
        item = QListWidgetItem(message)
        item.setIcon(dot_icon(_DOT.get(category, _DOT["info"])))
        self.log_list.addItem(item)
        self.log_list.scrollToBottom()

    def _browse(self):
        folder = QFileDialog.getExistingDirectory(self, "Папка пункта")
        if folder:
            self.edit_folder.setText(os.path.normpath(folder))

    def set_embedded(self, embedded: bool = True):
        """Встроенный режим (внутри «Ручного контроля»): скрыть заголовок и
        панель выбора пункта — пунктом управляет host через load_point()."""
        self._header.setVisible(not embedded)
        self._panel.setVisible(not embedded)

    def load_point(self, folder: str):
        """Загрузить пункт и сразу просканировать (для host-страницы)."""
        self.edit_folder.setText(os.path.normpath(folder))
        self._start_scan()

    def _point_folder(self) -> str:
        return self.edit_folder.text().strip()

    # ═════════════════════════════ Сканирование ════════════════════════

    def _start_scan(self):
        folder = self._point_folder()
        if not os.path.isdir(folder):
            self._log("Укажите существующую папку пункта", "error")
            return
        from workers.repair_worker import RepairScanWorker
        self.btn_scan.setEnabled(False)
        w = RepairScanWorker(folder, parent=self)
        w.log.connect(lambda m: self._log(m, "info"))
        w.info.connect(self._log)
        w.result.connect(self._on_scan)
        w.finished_ok.connect(lambda ok: self.btn_scan.setEnabled(True))
        self._workers["scan"] = w
        w.start()

    def _on_scan(self, res: dict):
        self._scan = res
        self._preview = None
        docs = res.get("docs")

        self.list_docs.clear()
        if not docs or not docs.has_repair:
            item = QListWidgetItem("Документов ремонта нет — у пункта состояние "
                                   "«ремонт не проводился»")
            item.setIcon(dot_icon(_DOT["info"]))
            self.list_docs.addItem(item)
            self.tbl_registry.setRowCount(0)
            self.tbl_sections.setRowCount(0)
            self._refresh_preview()
            return

        unit, line = res.get("unit", ""), res.get("line", "")
        for f, parsed, verdict in res.get("verdicts", []):
            cat, note = _VERDICT[verdict]
            extra = []
            if parsed.acts:
                extra.append("акт " + ", ".join(parsed.acts))
            if parsed.date:
                extra.append(parsed.date)
            text = f"{os.path.basename(f)} — {note}"
            if extra:
                text += f" ({'; '.join(extra)})"
            item = QListWidgetItem(text)
            item.setIcon(dot_icon(_DOT[cat]))
            item.setToolTip(f)
            self.list_docs.addItem(item)
        if unit or line:
            self._log(f"Карточка пункта: установка {unit or '—'}, "
                      f"трубопровод № {line or '—'}", "info")

        src = "лист «Ремонт»" if res.get("registry_existed") else "имена документов (черновик)"
        self._log(f"Реестр: {src}, записей {len(res.get('records', []))}", "info")
        self._fill_registry(res.get("records", []))
        self._fill_sections(res.get("sections", []))
        self._refresh_preview()
        self._map_autoload()

    # ═════════════════════════ Превью документа ════════════════════════

    def _doc_selected(self, current, _previous=None):
        if current is None:
            return
        path = current.toolTip()
        if not path or not os.path.isfile(path):
            return
        try:
            from repair.pdf_preview import page_count
            self._doc_path = path
            self._doc_pages = page_count(path)
            self._doc_page_idx = 0
            self._doc_show()
        except Exception as e:
            self._log(f"Превью не открылось: {e}", "error")

    def _doc_page(self, delta: int):
        if not self._doc_path:
            return
        idx = self._doc_page_idx + delta
        if 0 <= idx < self._doc_pages:
            self._doc_page_idx = idx
            self._doc_show()

    def _doc_show(self):
        from repair.pdf_preview import render_page
        arr = render_page(self._doc_path, self._doc_page_idx)
        self.doc_view._fitted = False
        self.doc_view.set_image(self._np_to_qimage(arr))
        self.lbl_pg.setText(
            f"стр. {self._doc_page_idx + 1} / {self._doc_pages} · "
            f"{os.path.basename(self._doc_path)[:40]}")
        self.btn_pg_prev.setEnabled(self._doc_page_idx > 0)
        self.btn_pg_next.setEnabled(self._doc_page_idx < self._doc_pages - 1)

    def _doc_open_external(self):
        if self._doc_path and os.path.isfile(self._doc_path):
            os.startfile(self._doc_path)

    def _doc_ocr(self):
        if not self._doc_path:
            self._log("Сначала выберите документ в списке", "warn")
            return
        from workers.repair_worker import RepairOcrWorker
        self.btn_ocr.setEnabled(False)
        w = RepairOcrWorker(self._doc_path, parent=self)
        w.log.connect(lambda m: self._log(m, "info"))
        w.info.connect(self._log)
        w.result.connect(self._on_ocr)
        w.map_ready.connect(self._on_id_map)
        w.zone_hint.connect(self._on_zone_hint)
        w.finished_ok.connect(lambda ok: self.btn_ocr.setEnabled(True))
        self._workers["ocr"] = w
        w.start()

    def _on_id_map(self, page_map):
        """Карта разделов комплекта ИД: показываем, где что лежит."""
        self._id_map = page_map
        if not page_map:
            return
        if not page_map.get("разделы"):
            self._log(f"Реестр ИД найден на стр. {page_map['страница_реестра']}, "
                      f"но таблица не разобрана — разделы придётся искать глазами",
                      "warn")
            return
        from repair.id_map import quality_cert_section, find_section, SEC_ACT
        части = []
        акт = find_section(page_map, SEC_ACT)
        if акт:
            части.append(f"копия акта — стр. {акт['pdf'][0]}–{акт['pdf'][1]}")
        cert = quality_cert_section(page_map)
        if cert:
            части.append(f"удостоверение о качестве — стр. "
                         f"{cert['pdf'][0]}–{cert['pdf'][1]}")
        if части:
            self._log("В комплекте: " + "; ".join(части), "info")

    def _on_zone_hint(self, hint):
        """
        Показать сотруднику фрагмент, который OCR прочитать не может.

        Рукописные даты и штампы Windows OCR не распознаёт вовсе, поэтому
        программа открывает нужный кроп сама — иначе пришлось бы листать
        комплект на сотню страниц.
        """
        if not hint or not self._doc_path:
            return
        try:
            from repair.pdf_preview import render_zone
            полоса = hint.get("полоса") or [0.6, 1.0]
            arr = render_zone(self._doc_path, hint["page"] - 1,
                              полоса[0], полоса[1],
                              rotation=hint.get("поворот", 0), scale=2.8)
            self.doc_view._fitted = False
            self.doc_view.set_image(self._np_to_qimage(arr))
            self._doc_page_idx = hint["page"] - 1
            подпись = hint.get("зачем", "фрагмент")
            self.lbl_pg.setText(f"стр. {hint['page']} · {подпись}")
            if not hint.get("найдено"):
                self._log(f"Впишите «{подпись}» в реестр по этому фрагменту "
                          f"(стр. {hint['page']})", "wait")
        except Exception as e:
            self._log(f"Фрагмент не показан: {e}", "warn")

    def _on_ocr(self, info):
        from repair.doc_extract import merge_into_records
        records = self._registry_records()
        filled = merge_into_records(records, info)
        if filled:
            self._fill_registry(records)
            self._log(f"Реестр: заполнено полей из распознанного — {filled} "
                      f"(только пустые ячейки)", "done")
        else:
            self._log("Распознанное не добавило новых полей (заполнено "
                      "вручную или номера актов не совпали)", "info")
        # сверка содержимого с карточкой пункта — второй уровень светофора
        unit, line = self._scan.get("unit", ""), self._scan.get("line", "")
        if info.unit and unit and info.unit != unit:
            self._log(f"Внимание: в документе установка {info.unit}, "
                      f"у пункта {unit} — проверьте привязку", "error")
        if info.line and line and info.line != line:
            self._log(f"Внимание: в документе линия №{info.line}, "
                      f"у пункта №{line} — проверьте привязку", "error")

    # ═════════════════════════════ Реестр ══════════════════════════════

    def _fill_registry(self, records):
        from repair.registry import records_to_rows
        rows = records_to_rows(records)
        self.tbl_registry.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, val in enumerate(row[1:]):        # без колонки «№»
                self.tbl_registry.setItem(r, c, QTableWidgetItem(val))

    def _registry_records(self):
        from repair.registry import rows_to_records
        rows = []
        for r in range(self.tbl_registry.rowCount()):
            row = [str(r + 1)]
            for c in range(8):
                it = self.tbl_registry.item(r, c)
                row.append(it.text() if it else "")
            rows.append(row)
        return rows_to_records(rows)

    def _reg_add(self):
        self.tbl_registry.insertRow(self.tbl_registry.rowCount())

    def _reg_del(self):
        r = self.tbl_registry.currentRow()
        if r >= 0:
            self.tbl_registry.removeRow(r)

    def _reg_prefill(self):
        docs = self._scan.get("docs")
        if not docs:
            self._log("Сначала «Проверить пункт»", "warn")
            return
        from repair.detect import build_records_from_docs
        self._fill_registry(build_records_from_docs(docs))
        self._log("Реестр заполнен из документов заново (черновик)", "info")

    # ═════════════════════════════ Сечения ═════════════════════════════

    def _fill_sections(self, sections):
        self.tbl_sections.setRowCount(len(sections))
        for r, s in enumerate(sections):
            vals = [str(s.number), s.element, s.thickness_nom, s.diameter,
                    " / ".join(m for m in s.measures if m), s.tmin, ""]
            for c, val in enumerate(vals):
                it = QTableWidgetItem(val)
                if c < 6:
                    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                self.tbl_sections.setItem(r, c, it)
            # строка листа — в данных первой ячейки
            self.tbl_sections.item(r, 0).setData(Qt.UserRole, s.row)
        self._apply_sections_filter()

    def _apply_sections_filter(self):
        """
        Скрыть строки, которые по правилам не обновляются.

        Задвижек и бобышек в пункте бывает больше, чем труб, и они занимают
        экран, мешая искать заменённое сечение. Строку с уже вписанным
        номиналом не прячем никогда — иначе правка исчезнет из виду.
        """
        from repair.config import REPAIR_CONFIG

        мелкие = [e.lower() for e in
                  (list(REPAIR_CONFIG.get("SMALL_ELEMENTS") or [])
                   + list(REPAIR_CONFIG.get("EXCLUDE_FROM_TABLE3") or []))]
        прятать = self.chk_hide_small.isChecked()
        скрыто = 0
        for r in range(self.tbl_sections.rowCount()):
            эл = self.tbl_sections.item(r, 1)
            ном = self.tbl_sections.item(r, 6)
            мелкий = bool(эл) and any(эл.text().strip().lower().startswith(m)
                                      for m in мелкие)
            заполнен = bool(ном) and bool(ном.text().strip())
            надо_скрыть = прятать and мелкий and not заполнен
            self.tbl_sections.setRowHidden(r, надо_скрыть)
            скрыто += int(надо_скрыть)
        if прятать and скрыто:
            self.chk_hide_small.setText(f"Скрыть задвижки и бобышки ({скрыто})")
        else:
            self.chk_hide_small.setText("Скрыть задвижки и бобышки")

    def _toggle_scheme(self, показать: bool):
        """Показать/скрыть лист акта со схемой дефектного участка."""
        if not показать:
            self.scheme_view.setVisible(False)
            return
        docs = self._scan.get("docs")
        файлы = list(getattr(docs, "act_files", []) or [])
        if not файлы:
            self._log("Отдельного акта отбраковки в пункте нет — схему "
                      "смотрите в комплекте ИД", "warn")
            self.btn_scheme.setChecked(False)
            return
        try:
            from repair.pdf_preview import page_count, render_zone
            путь = файлы[0]
            # Схема — на последнем листе акта, над таблицей замены.
            стр = page_count(путь)
            arr = render_zone(путь, стр - 1, 0.0, 0.75, scale=2.0)
            self.scheme_view._fitted = False
            self.scheme_view.set_image(self._np_to_qimage(arr))
            self.scheme_view.setVisible(True)
            self._log(f"Схема дефектного участка: {os.path.basename(путь)}, "
                      f"лист {стр}", "info")
        except Exception as e:
            self._log(f"Схему показать не удалось: {e}", "warn")
            self.btn_scheme.setChecked(False)

    def _replacement_lines(self) -> list[str]:
        """Строки таблицы замены из реестра (колонка «Замена»)."""
        out = []
        for r in range(self.tbl_registry.rowCount()):
            it = self.tbl_registry.item(r, 4)
            текст = (it.text().strip() if it else "")
            for часть in текст.replace(";", "\n").splitlines():
                часть = часть.strip()
                if часть:
                    out.append(часть)
        return out

    def _suggest_sections(self):
        """
        Подобрать сечения по таблице замены из документов.

        Однозначного ответа обычно нет: в пункте бывает по два десятка труб
        одного диаметра. Поэтому номинал подставляется только там, где сечение
        одно, а прочие кандидаты просто помечаются — выбрать место помогает
        карта, где видно, какой участок заменён.
        """
        from repair.suggest import suggest_sections, describe

        строки = self._replacement_lines()
        сечения = self._scan.get("sections", [])
        if not строки:
            self._log("В реестре не заполнена колонка «Замена» — подбирать "
                      "не по чему", "warn")
            return
        if not сечения:
            self._log("Сечения не прочитаны — сначала «Проверить пункт»", "warn")
            return

        # снять прежние пометки
        for r in range(self.tbl_sections.rowCount()):
            it = self.tbl_sections.item(r, 0)
            if it is not None:
                it.setIcon(QIcon())
                it.setToolTip("")

        по_строке = {}
        for r in range(self.tbl_sections.rowCount()):
            it = self.tbl_sections.item(r, 0)
            if it is not None:
                по_строке[it.data(Qt.UserRole)] = r

        всего = подставлено = 0
        for s in suggest_sections(строки, сечения):
            self._log(describe(s), "done" if s.кандидаты else "warn")
            for c in s.кандидаты:
                r = по_строке.get(c.row)
                if r is None:
                    continue
                всего += 1
                it = self.tbl_sections.item(r, 0)
                it.setIcon(dot_icon(_DOT["done" if s.однозначно else "warn"]))
                it.setToolTip(f"{s.строка} → предлагаемый номинал {s.номинал}")
                if s.однозначно:
                    цель = self.tbl_sections.item(r, 6)
                    if цель is not None and not цель.text().strip():
                        цель.setText(s.номинал)
                        подставлено += 1
        self._log(f"Подбор: помечено сечений {всего}, номинал подставлен "
                  f"в {подставлено}", "info")
        self._refresh_map_plan()

    def _collect_nominals(self):
        """Из колонки «Новый номинал» → ({row: float}, {row: (толщ, диам)}, ошибки)."""
        rows: dict[int, float] = {}
        perehod: dict[int, tuple] = {}
        errors: list[str] = []
        secs = {s.row: s for s in self._scan.get("sections", [])}
        for r in range(self.tbl_sections.rowCount()):
            it = self.tbl_sections.item(r, 6)
            text = (it.text().strip() if it else "")
            if not text:
                continue
            sheet_row = self.tbl_sections.item(r, 0).data(Qt.UserRole)
            sec = secs.get(sheet_row)
            num = self.tbl_sections.item(r, 0).text()
            if "/" in text:
                # переход: «12 / 10» — толщина текстом, номинал = первое число
                try:
                    first = float(text.split("/")[0].replace(",", "."))
                except ValueError:
                    errors.append(f"сечение {num}: не разобрано «{text}»")
                    continue
                rows[sheet_row] = first
                diam = sec.diameter if sec else ""
                perehod[sheet_row] = (text, diam)
            else:
                try:
                    rows[sheet_row] = float(text.replace(",", "."))
                except ValueError:
                    errors.append(f"сечение {num}: не разобрано «{text}»")
        return rows, perehod, errors

    # ═════════════════════════════ Предпросмотр ════════════════════════

    def _on_tab_changed(self, index: int):
        вкладка = self.tabs.tabText(index)
        if вкладка == "Предпросмотр":
            self._refresh_preview()
            # Пересчёт высоты строк — уже на видимой вкладке: колонки тянутся
            # по ширине, и при скрытой вкладке высота считается по неверной.
            self.tbl_preview.resizeRowsToContents()
        elif вкладка == "Карта":
            # Номиналы могли вписать уже после загрузки карты.
            self._refresh_map_plan()

    def _refresh_preview(self):
        from repair.nominals import apply_nominals
        from repair.imports import build_write_ops
        from repair.svedenia import build_svedenia

        self.tbl_preview.setRowCount(0)
        self._preview = None
        self.btn_apply.setEnabled(False)
        self.btn_checklist.setEnabled(False)

        sections = self._scan.get("sections", [])
        if not self._scan or not self._scan.get("docs") \
                or not self._scan["docs"].has_repair:
            self._set_svedenia_text("")
            self.txt_sved_now.setPlainText("")
            return

        records = self._registry_records()
        собранная = build_svedenia(records)
        now = self._scan.get("svedenia_now")
        self._sved_generated = собранная
        self.txt_sved_now.setPlainText(now or "(строка в отчёте не найдена)")
        # Правку сотрудника не затираем: пересборка предпросмотра случается на
        # каждом переключении вкладки.
        if self._sved_edited:
            self.lbl_sved.setText("Станет — изменено вручную")
        else:
            self.lbl_sved.setText("Станет")
            self._set_svedenia_text(собранная)
        svedenia = self._svedenia_text()

        rows, perehod, errors = self._collect_nominals()
        for e in errors:
            self._log(e, "error")
        new, changes = apply_nominals(sections, rows, perehod, by_row=True)

        # операции, сгруппированные по строке листа — для согласования с галочками
        ops_all = build_write_ops(sections, new)
        ops_by_row = {}
        for op in ops_all:
            ops_by_row.setdefault(op.row, []).append(op)

        self.tbl_preview.setRowCount(len(changes) + 1)
        # строка 0 — Word
        chk = QTableWidgetItem()
        chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        chk.setCheckState(Qt.Checked)
        self.tbl_preview.setItem(0, 0, chk)
        self.tbl_preview.setItem(0, 1, QTableWidgetItem("Word · таблица 1"))
        self.tbl_preview.setItem(0, 2, QTableWidgetItem(now or "(не найдена)"))
        self.tbl_preview.setItem(0, 3, QTableWidgetItem(svedenia))

        for i, ch in enumerate(changes, start=1):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Checked)
            self.tbl_preview.setItem(i, 0, chk)
            self.tbl_preview.setItem(i, 1, QTableWidgetItem(f"Импорт · {ch.where}"))
            self.tbl_preview.setItem(i, 2, QTableWidgetItem(ch.before))
            self.tbl_preview.setItem(i, 3, QTableWidgetItem(ch.after))
        self.tbl_preview.resizeRowsToContents()

        # сопоставление галочка → операции: по номеру строки листа
        sheet_rows_in_order = [s.row for s in new if s.row in rows]
        self._preview = {
            "records": records,
            "svedenia": svedenia,
            "ops_by_row": ops_by_row,
            "sheet_rows_in_order": sheet_rows_in_order,
            "changes": changes,
        }
        has_any = bool(changes) or bool(records)
        self.btn_apply.setEnabled(has_any)
        self.btn_checklist.setEnabled(has_any)

    # ── строка «Сведения о ремонте»: показ и правка ──────────────────

    def _svedenia_text(self) -> str:
        return " ".join(self.txt_sved.toPlainText().split()).strip()

    def _set_svedenia_text(self, text: str):
        """Записать строку в поле, не пометив её как правку сотрудника."""
        self._sved_programmatic = True
        try:
            self.txt_sved.setPlainText(text)
        finally:
            self._sved_programmatic = False

    def _on_sved_edited(self):
        if self._sved_programmatic:
            return
        self._sved_edited = True
        self.lbl_sved.setText("Станет — изменено вручную")
        # В таблице предпросмотра показываем именно то, что будет записано.
        it = self.tbl_preview.item(0, 3)
        if it is not None:
            it.setText(self._svedenia_text())
            self.tbl_preview.resizeRowsToContents()

    def _reset_svedenia(self):
        self._sved_edited = False
        self._set_svedenia_text(self._sved_generated)
        self.lbl_sved.setText("Станет")
        self._refresh_preview_keep_checks()
        self._log("Строка «Сведения о ремонте» восстановлена по реестру", "info")

    def _gather_enabled(self):
        """Считать галочки предпросмотра → (ops, svedenia_or_None)."""
        p = self._preview or {}
        if not p:
            return [], None
        sved = None
        it = self.tbl_preview.item(0, 0)
        if it is not None and it.checkState() == Qt.Checked:
            # Берём текст из поля: сотрудник мог его поправить.
            sved = self._svedenia_text() or p["svedenia"]
        ops = []
        order = p["sheet_rows_in_order"]
        for i, sheet_row in enumerate(order, start=1):
            it = self.tbl_preview.item(i, 0)
            if it is not None and it.checkState() == Qt.Checked:
                ops.extend(p["ops_by_row"].get(sheet_row, []))
        return ops, sved

    # ═════════════════════════════ Применение ══════════════════════════

    def _apply(self):
        folder = self._point_folder()
        xlsm = self._scan.get("xlsm")
        if not xlsm:
            self._log("В пункте не найден ПНОС_ТТП.xlsm", "error")
            return
        self._refresh_preview_keep_checks()
        ops, sved = self._gather_enabled()
        records = (self._preview or {}).get("records", [])
        if not ops and sved is None:
            self._log("Нечего применять: все галочки сняты", "warn")
            return

        from workers.repair_worker import RepairApplyWorker
        self.btn_apply.setEnabled(False)
        w = RepairApplyWorker(folder, xlsm, self._scan.get("report"),
                              ops, records, sved, parent=self)
        w.log.connect(lambda m: self._log(m, "info"))
        w.info.connect(self._log)
        w.finished_ok.connect(self._on_applied)
        self._workers["apply"] = w
        w.start()

    def _refresh_preview_keep_checks(self):
        """Перед применением пересобрать операции, сохранив снятые галочки."""
        states = [self.tbl_preview.item(r, 0).checkState()
                  for r in range(self.tbl_preview.rowCount())
                  if self.tbl_preview.item(r, 0)]
        self._refresh_preview()
        for r, st in enumerate(states[:self.tbl_preview.rowCount()]):
            it = self.tbl_preview.item(r, 0)
            if it is not None:
                it.setCheckState(st)

    def _on_applied(self, ok: bool):
        self.btn_apply.setEnabled(True)
        if ok:
            self._log("Готово. Повторное сканирование пункта...", "done")
            self._start_scan()

    def _verify(self):
        folder = self._point_folder()
        if not os.path.isdir(folder):
            self._log("Укажите папку пункта", "error")
            return
        from workers.repair_worker import RepairVerifyWorker
        self.btn_verify.setEnabled(False)
        w = RepairVerifyWorker(folder, parent=self)
        w.log.connect(lambda m: self._log(m, "info"))
        w.info.connect(self._log)
        w.result.connect(self._on_verified)
        w.finished_ok.connect(lambda ok: self.btn_verify.setEnabled(True))
        self._workers["verify"] = w
        w.start()

    def _on_verified(self, findings):
        cat = {"ok": "done", "warn": "warn", "error": "error"}
        for f in findings:
            self._log(f"[{f.code} {f.where}] {f.message}", cat.get(f.level, "info"))

    def _rollback(self):
        folder = self._point_folder()
        if not os.path.isdir(folder):
            self._log("Укажите папку пункта", "error")
            return
        from repair.safety import latest_backup, restore_backup, append_journal
        b = latest_backup(folder)
        if not b:
            self._log("Бэкапов у пункта нет", "warn")
            return
        restored = restore_backup(b, folder)
        append_journal(folder, "Откат по кнопке",
                       [f"восстановлено файлов: {len(restored)}",
                        f"бэкап: {b}"])
        self._log(f"Откат из {os.path.basename(b)}: восстановлено "
                  f"{len(restored)} файла(ов)", "done")
        self._start_scan()

    # ═════════════════════════════ Карта ══════════════════════════════

    @staticmethod
    def _find_map_path(folder: str) -> str | None:
        """Рабочая карта пункта: Первичка/NNNN.png (оригинал ОТН «NNNN..png»
        не трогается никогда)."""
        import re
        from repair.detect import _find_subfolder
        perv = _find_subfolder(folder, "Первичка")
        if not perv:
            return None
        m = re.search(r"(\d+)", os.path.basename(folder))
        num = m.group(1) if m else None
        cands = []
        for f in sorted(os.listdir(perv)):
            if not f.lower().endswith(".png") or f.lower().endswith("..png"):
                continue
            p = os.path.join(perv, f)
            if num and os.path.splitext(f)[0] == num:
                return p
            cands.append(p)
        return cands[0] if cands else None

    @staticmethod
    def _np_to_qimage(arr) -> QImage:
        import numpy as np
        a = np.ascontiguousarray(arr)
        h, w = a.shape[:2]
        return QImage(a.data, w, h, 3 * w, QImage.Format_RGB888).copy()

    def _map_autoload(self):
        """
        Подгрузить карту сразу при проверке пункта.

        Путь к рабочей карте выводится из номера пункта («Первичка/NNNN.png»),
        так что просить сотрудника нажимать «Загрузить карту» незачем. Если
        карты нет — молча ничего не делаем: у пункта её может и не быть.
        """
        folder = self._point_folder()
        if not os.path.isdir(folder) or not self._find_map_path(folder):
            self.lbl_map_status.setText(
                "Рабочая карта (Первичка/NNNN.png) в пункте не найдена")
            return
        self._map_load()

    def _map_load(self):
        folder = self._point_folder()
        if not os.path.isdir(folder):
            self._log("Укажите папку пункта", "error")
            return
        path = self._find_map_path(folder)
        if not path:
            self._log("Рабочая карта (Первичка/NNNN.png) не найдена", "error")
            return
        import numpy as np
        from PIL import Image
        self._map_path = path
        self._map_orig = np.array(Image.open(path).convert("RGB"))
        self._map_arr = self._map_orig.copy()
        self._map_edits = []
        self._map_pending = None
        self.list_map_edits.clear()
        self.map_view._fitted = False
        self.map_view.set_image(self._np_to_qimage(self._map_arr))
        self.btn_map_undo.setEnabled(False)
        self.btn_map_save.setEnabled(False)
        self.lbl_map_status.setText(
            f"{os.path.basename(path)} · кликните по подписи заменённого сечения")
        self._log(f"Карта загружена: {os.path.basename(path)}", "info")
        self._map_cycle = {}
        self.list_map_edits.clear()
        self._map_scan_labels()

    # ── подписи сечений на карте и план правок ───────────────────────

    def _map_scan_labels(self):
        """Найти подписи сечений на загруженной карте (фоном)."""
        if self._map_arr is None:
            return
        from workers.repair_worker import MapLabelsWorker
        self._map_labels = []
        self._map_by_number = {}
        self.lbl_map_status.setText(
            f"{os.path.basename(self._map_path or '')} · ищу подписи сечений…")
        w = MapLabelsWorker(self._map_arr.copy(), parent=self)
        w.log.connect(lambda m: self._log(m, "info"))
        w.info.connect(self._log)
        w.result.connect(self._on_map_labels)
        self._workers["map_labels"] = w
        w.start()

    def _on_map_labels(self, labels):
        from repair.map_labels import labels_by_number
        self._map_labels = labels or []
        self._map_by_number = labels_by_number(self._map_labels)
        self.lbl_map_status.setText(
            f"{os.path.basename(self._map_path or '')} · "
            f"распознано сечений {len(self._map_by_number)}")
        self._refresh_map_plan()

    def _planned_edits(self) -> list[tuple[int, str]]:
        """Что предстоит поправить на карте: (номер сечения, новое значение)
        из колонки «Новый номинал» вкладки «Сечения»."""
        out = []
        for r in range(self.tbl_sections.rowCount()):
            it_new = self.tbl_sections.item(r, 6)
            текст = (it_new.text().strip() if it_new else "")
            if not текст:
                continue
            it_num = self.tbl_sections.item(r, 0)
            if it_num is None or not it_num.text().strip().isdigit():
                continue
            out.append((int(it_num.text().strip()), текст))
        return out

    def _refresh_map_plan(self):
        """Список правок с состоянием привязки к карте."""
        # Пересборка случается и при возврате на вкладку — выбранную подпись
        # при этом терять нельзя, иначе кнопка «Проставить показанное» гаснет
        # ровно тогда, когда сотрудник к ней тянется.
        было = self._map_shown
        self.list_map_plan.clear()
        план = self._planned_edits()
        if было is not None and (было[0], было[1]) not in план:
            self._map_shown = None
            self.btn_map_apply_one.setEnabled(False)
        if not план:
            item = QListWidgetItem("Ничего не отмечено: впишите новые номиналы "
                                   "на вкладке «Сечения»")
            item.setIcon(dot_icon(_DOT["info"]))
            self.list_map_plan.addItem(item)
            self.btn_map_apply_unique.setEnabled(False)
            return

        однозначных = 0
        for номер, значение in план:
            места = self._map_by_number.get(номер, [])
            if len(места) == 1:
                цвет, хвост = "done", "подпись найдена"
                однозначных += 1
            elif len(места) > 1:
                цвет, хвост = "warn", f"мест на карте: {len(места)} — выберите"
            else:
                цвет, хвост = "wait", "подпись не найдена, правьте кликом"
            item = QListWidgetItem(f"сечение {номер} → {значение} · {хвост}")
            item.setIcon(dot_icon(_DOT[цвет]))
            item.setData(Qt.UserRole, (номер, значение))
            self.list_map_plan.addItem(item)
        self.btn_map_apply_unique.setEnabled(однозначных > 0
                                            and self._map_arr is not None)

    def _map_plan_clicked(self, item):
        """Перейти к подписи сечения; при нескольких местах — к следующему."""
        данные = item.data(Qt.UserRole)
        if not данные or self._map_arr is None:
            return
        номер, значение = данные
        места = self._map_by_number.get(номер, [])
        if not места:
            self._log(f"Сечение {номер}: подпись на карте не распознана — "
                      f"кликните по ней сами", "warn")
            return
        i = self._map_cycle.get(номер, -1) + 1
        if i >= len(места):
            i = 0
        self._map_cycle[номер] = i
        подпись = места[i]
        self.map_view.show_at(*подпись.point, box=подпись.bbox)
        self._map_shown = (номер, значение, подпись)
        self.btn_map_apply_one.setEnabled(True)
        уточнение = f" (место {i + 1} из {len(места)})" if len(места) > 1 else ""
        self._log(f"Сечение {номер}{уточнение}: подпись показана на карте",
                  "info")

    @staticmethod
    def _map_value_text(значение: str) -> str:
        """Значение в том виде, как оно стоит на схеме: «(6.0)»."""
        text = (значение or "").strip()
        if not text.startswith("("):
            text = f"({text})" if not text.endswith(")") else f"({text}"
        return text

    def _map_write(self, номер: int, значение: str, подпись) -> bool:
        """Записать новое значение в подпись сечения на карте."""
        from repair.map_editor import render_value, render_value_in
        значение = self._map_value_text(значение)
        try:
            if подпись.value_bbox is not None:
                # Рамка значения известна из распознанной подписи — берём её.
                # Выращивать рамку от точки нельзя: рядом с подписью стоит
                # красный крестик замера, и рамка «раздувается» на него.
                rep = render_value_in(self._map_arr, подпись.value_bbox,
                                      значение, label=подпись.bbox)
            else:
                rep = render_value(self._map_arr, подпись.point, значение)
        except Exception as e:
            self._log(f"Сечение {номер}: правка не удалась — {e}", "error")
            return False
        self._map_edits.append((подпись.point, None,
                                f"сечение {номер} → {значение}"))
        item = QListWidgetItem(f"сечение {номер} → {значение}")
        item.setIcon(dot_icon(_DOT["done"]))
        self.list_map_edits.addItem(item)
        self.lbl_map_before.setPixmap(
            self._crop_pixmap(self._map_orig, rep["label"]))
        self.lbl_map_after.setPixmap(
            self._crop_pixmap(self._map_arr, rep["label"]))
        self.map_view.set_image(self._np_to_qimage(self._map_arr))
        self.map_view.show_at(*подпись.point, box=rep["label"])
        self.btn_map_undo.setEnabled(True)
        self.btn_map_save.setEnabled(True)
        return True

    def _map_apply_shown(self):
        if not self._map_shown:
            return
        номер, значение, подпись = self._map_shown
        if self._map_write(номер, значение, подпись):
            self._log(f"Сечение {номер}: на карте записано {значение}", "done")

    def _map_apply_unique(self):
        """
        Проставить значения там, где подпись сечения на карте одна.

        Места с несколькими одинаковыми сечениями намеренно не трогаются:
        какое из них заменено, видно только по схеме — решает сотрудник.
        """
        сделано = отложено = 0
        for номер, значение in self._planned_edits():
            места = self._map_by_number.get(номер, [])
            if len(места) != 1:
                отложено += 1
                continue
            if self._map_write(номер, значение, места[0]):
                сделано += 1
        self._log(f"Карта: проставлено {сделано}"
                  + (f", отложено {отложено} (несколько мест или подпись "
                     f"не найдена)" if отложено else ""),
                  "done" if сделано else "warn")

    def _crop_pixmap(self, arr, bbox, pad=30) -> QPixmap:
        h, w = arr.shape[:2]
        crop = arr[max(0, bbox.y1 - pad):min(h, bbox.y2 + pad),
                   max(0, bbox.x1 - pad):min(w, bbox.x2 + pad)]
        pm = QPixmap.fromImage(self._np_to_qimage(crop))
        return pm.scaledToWidth(250, Qt.SmoothTransformation) \
            if pm.width() > 250 else pm

    def _on_map_clicked(self, x: float, y: float):
        """Клик по подписи → ввод нового значения (шрифт Arial и цвет
        берутся из самой подписи, поэтому оранжевые сечения остаются
        оранжевыми)."""
        if self._map_arr is None:
            return
        from PySide6.QtWidgets import QInputDialog
        from repair.map_editor import label_bbox_at, find_value_span

        xy = (int(x), int(y))
        current = ""
        try:
            lb = label_bbox_at(self._map_arr, *xy)
            if lb is not None and find_value_span(self._map_arr, lb) is not None:
                self.lbl_map_before.setPixmap(self._crop_pixmap(self._map_arr, lb))
        except Exception:
            pass

        text, ok = QInputDialog.getText(
            self, "Новое значение",
            "Значение в скобках (как на схеме, напр. «(6.0)»):", text=current)
        if not ok:
            self.lbl_map_status.setText("Отменено. Кликните по подписи сечения")
            return
        text = text.strip()
        if not text:
            return
        if not text.startswith("("):
            text = f"({text})" if not text.endswith(")") else f"({text}"
        self._map_apply_text(xy, text)

    def _map_apply_text(self, label_xy, text):
        from repair.map_editor import render_value
        before = self._map_arr.copy()
        try:
            rep = render_value(self._map_arr, label_xy, text)
        except ValueError as e:
            self._map_arr = before
            self._log(f"Карта: {e}", "error")
            self.lbl_map_status.setText(f"Не получилось: {e}. "
                                        f"Кликните по подписи сечения")
            return
        lb = rep["label"]
        self._map_edits.append((label_xy, text))
        self.lbl_map_before.setPixmap(self._crop_pixmap(self._map_orig, lb))
        self.lbl_map_after.setPixmap(self._crop_pixmap(self._map_arr, lb))
        item = QListWidgetItem(
            f"правка {len(self._map_edits)}: {rep['text']} · Arial {rep['font_size']} · "
            f"цвет RGB{rep['color']}")
        item.setIcon(dot_icon(_DOT["done"]))
        self.list_map_edits.addItem(item)
        self.map_view.set_image(self._np_to_qimage(self._map_arr))
        self.btn_map_undo.setEnabled(True)
        self.btn_map_save.setEnabled(True)
        self.lbl_map_status.setText(
            f"Правок: {len(self._map_edits)}. Проверьте «до/после» и "
            f"кликните по следующей подписи или сохраните карту")

    def _map_undo(self):
        if not self._map_edits:
            return
        from repair.map_editor import render_value
        self._map_edits.pop()
        self.list_map_edits.takeItem(self.list_map_edits.count() - 1)
        self._map_arr = self._map_orig.copy()
        for label_xy, text in self._map_edits:
            try:
                render_value(self._map_arr, label_xy, text)
            except ValueError:
                pass
        self.map_view.set_image(self._np_to_qimage(self._map_arr))
        self.btn_map_undo.setEnabled(bool(self._map_edits))
        self.btn_map_save.setEnabled(bool(self._map_edits))
        self._log("Последняя правка карты отменена", "info")

    def _map_save(self):
        if self._map_arr is None or not self._map_path:
            return
        from PIL import Image
        from repair.safety import make_backup, append_journal
        folder = self._point_folder()
        b = make_backup(folder, [self._map_path])
        Image.fromarray(self._map_arr).save(self._map_path)
        append_journal(folder, "Правка коррозионной карты",
                       [f"файл: {os.path.basename(self._map_path)}",
                        f"правок: {len(self._map_edits)}",
                        f"бэкап: {b}"])
        self._log(f"Карта сохранена ({len(self._map_edits)} правок), "
                  f"бэкап и журнал созданы", "done")
        self._map_orig = self._map_arr.copy()
        self._map_edits = []
        self.list_map_edits.clear()
        self.btn_map_undo.setEnabled(False)
        self.btn_map_save.setEnabled(False)

    def _export_checklist(self):
        folder = self._point_folder()
        p = self._preview or {}
        if not p:
            return
        lines = ["ЧЕК-ЛИСТ РУЧНОЙ ПРАВКИ (модуль «Ремонт»)", ""]
        lines.append("1. Строка «Сведения о ремонте» (таблица 1 отчёта):")
        # В чек-лист идёт то, что реально будет записано, — с правкой сотрудника.
        lines.append(f"   {self._svedenia_text() or p['svedenia']}")
        lines.append("")
        lines.append("2. Листы «Импорт» и «Импорт (ПР)» (одинаково):")
        for ch in p["changes"]:
            lines.append(f"   {ch.where}: {ch.before}  →  {ch.after}")
        lines.append("")
        lines.append("3. Коррозионная карта: те же сечения → номинал.")
        path = os.path.join(folder, CHECKLIST_NAME)
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        self._log(f"Чек-лист сохранён: {os.path.basename(path)}", "done")
