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

from widgets import dot_icon, SegmentBar, DateCellDelegate


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

# Колонки таблицы сечений. Правка разнесена на три поля: у перехода Тисп
# записана двусторонне («6 / 6»), и одно поле «новый номинал» затирало её
# половиной. Замер по умолчанию выводится из Тисп, но правится отдельно.
COL_NUM, COL_ELEM, COL_T, COL_D, COL_MEAS, COL_MIN = range(6)
COL_NEW_T, COL_NEW_D, COL_NEW_M = 6, 7, 8
NEW_COLS = (COL_NEW_T, COL_NEW_D, COL_NEW_M)

# Колонки таблицы реестра (колонка «№» на экран не выводится).
REG_COL_DATE_DONE = 1      # «Дата выполнения» — вводится по маске


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
        # Состояние вкладки «Карта». Карта бывает из нескольких листов, поэтому
        # всё, что относится к изображению, живёт в списке листов:
        #   {"path", "num", "arr", "labels", "by_number", "edits"}
        # Исходный массив листа не храним: лист весит до 37 МБ, а откат
        # перечитывает файл с диска — это редкая операция.
        self._sheets: list[dict] = []
        self._sheet_idx = 0            # какой лист показан
        self._map_shown = None         # (номер, значение, лист, подпись) на экране
        self._map_cycle: dict = {}     # номер → какое из мест показано
        self._map_dirty = False        # есть ли несохранённые правки
        # состояние превью документа (вкладка 1)
        self._doc_path: str | None = None
        self._doc_page_idx = 0
        self._doc_pages = 0
        self._id_map = None            # карта разделов комплекта ИД
        # Очередь распознавания: все документы ремонта пункта подряд, без
        # единого нажатия. Кроп даты показывается в конце всей очереди —
        # его даёт только комплект ИД, а он в списке обычно не первый.
        self._ocr_queue: list[str] = []
        self._ocr_idx = 0
        self._ocr_phase_offset = 0     # сколько сегментов полосы уже пройдено
        self._ocr_phase_sizes: list[int] = []
        self._ocr_hints: list[tuple[str, dict]] = []
        self._ocr_running = False
        # служебные флаги таблицы сечений и состояния записи
        self._filling_sections = False
        self._applied_summary = ""     # итог последней успешной записи
        self._applied_folder = ""      # пункт, к которому этот итог относится
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
        prev_col.addLayout(nav)
        # Полоса этапов распознавания — под превью, где и появится кроп даты.
        self.seg_ocr = SegmentBar()
        prev_col.addWidget(self.seg_ocr)
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
        # «Дата выполнения» вводится по маске ДД.ММ.ГГГГ: её вписывают руками
        # по кропу блока подписей, и формат должен быть единственно возможным.
        # «Даты актов» маской не закрываем — там бывает несколько дат через «;».
        self._date_delegate = DateCellDelegate(self.tbl_registry)
        self.tbl_registry.setItemDelegateForColumn(
            REG_COL_DATE_DONE, self._date_delegate)
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
            "Отметьте заменённые сечения тремя правыми колонками: «Новый Тисп» "
            "(«6», для перехода «6 / 5»), «Новый диаметр» («108 / 89») и «Новый "
            "замер» — он подставляется из Тисп сам и правится вручную. Пустое "
            "поле = не меняется, поэтому у перехода можно поправить только "
            "замер, оставив двустороннюю Тисп как есть. Задвижки и бобышки по "
            "правилам не обновляются — при необходимости впишите значения явно.")
        hint2.setObjectName("hint")
        hint2.setWordWrap(True)
        v2.addWidget(hint2)
        self.tbl_sections = QTableWidget(0, 9)
        self.tbl_sections.setHorizontalHeaderLabels(
            ["Сечение", "Элемент", "Тисп", "Диам", "Замеры", "Мин",
             "Новый Тисп", "Новый диаметр", "Новый замер"])
        self.tbl_sections.itemChanged.connect(self._on_section_edited)
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

        # Пока не нажата «Применить выбранное», всё вышесказанное живёт только
        # на экране. На п.2763 это и вышло боком: сечение поправили, кнопку не
        # нажали — и пункт ушёл на этапы 4–7 со старыми замерами.
        self.lbl_apply_state = QLabel("")
        self.lbl_apply_state.setObjectName("apply_state")
        self.lbl_apply_state.setWordWrap(True)
        self.lbl_apply_state.setVisible(False)
        v3.addWidget(self.lbl_apply_state)

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
        # Вкладку регистрируем после карты: предпросмотр с кнопкой записи —
        # последний шаг, и порядок вкладок должен совпадать с порядком работы.

        # Вкладка 3: карта
        tab_map = QWidget()
        v4 = QHBoxLayout(tab_map)
        v4.setSpacing(8)

        map_col = QVBoxLayout()
        map_col.setSpacing(6)
        # Переключатель листов: карта пункта бывает разбита на 2–3 листа.
        # При одном листе ряд скрыт — вид остаётся прежним.
        self.row_sheets = QHBoxLayout()
        self.row_sheets.setSpacing(4)
        self._sheet_buttons: list = []
        self.wid_sheets = QWidget()
        self.wid_sheets.setLayout(self.row_sheets)
        self.wid_sheets.setVisible(False)
        map_col.addWidget(self.wid_sheets)
        self.map_view = MapView()
        self.map_view.clicked.connect(self._on_map_clicked)
        map_col.addWidget(self.map_view, 1)
        self.seg_map = SegmentBar()
        map_col.addWidget(self.seg_map)
        v4.addLayout(map_col, 1)

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
        self.tabs.addTab(tab_prev, "Предпросмотр")

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

    def _start_scan(self, run_ocr: bool = True):
        """
        Перечитать пункт с диска.

        `run_ocr=False` — обновление после записи или отката: документы с тех
        пор не менялись, и гонять по ним распознавание заново (это минуты на
        комплект ИД) незачем.
        """
        folder = self._point_folder()
        if not os.path.isdir(folder):
            self._log("Укажите существующую папку пункта", "error")
            return
        from workers.repair_worker import RepairScanWorker
        self._scan_with_ocr = run_ocr
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
        # Документы распознаются сами, все подряд: реестр заполняется без
        # единого нажатия, а в конце появляется кроп с датой.
        if self.list_docs.count():
            self.list_docs.setCurrentRow(0)
        if getattr(self, "_scan_with_ocr", True):
            self._start_ocr_queue()

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
        # Распознавание здесь не запускаем: очередь по всем документам уже
        # идёт (или прошла) — клик по списку только листает превью.

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

    # ═══════════════════ Распознавание: очередь документов ═════════════

    def _start_ocr_queue(self):
        """
        Распознать все документы ремонта пункта подряд, без нажатий.

        Раньше запускался только открытый документ, а первым в списке лежит акт
        отбраковки — короткий, и кропа даты он не даёт: кроп рождается в блоке
        подписей комплекта ИД, который стоит вторым. Поэтому очередь идёт по
        всем документам, а кроп показывается в конце всей очереди.
        """
        from repair.pdf_preview import page_count
        from workers.repair_worker import RepairOcrWorker

        if self._ocr_running:
            return
        docs = self._scan.get("docs")
        файлы = list(getattr(docs, "all_files", []) or [])
        if not файлы:
            return

        очередь, размеры, этапы = [], [], []
        for p in файлы:
            try:
                страниц = page_count(p)
            except Exception as e:
                # Битые PDF встречаются (п.2704, файл 222-1-406): очередь из-за
                # одного такого останавливаться не должна.
                self._log(f"{os.path.basename(p)}: файл не читается ({e}) — "
                          f"пропускаю", "warn")
                continue
            набор = (RepairOcrWorker.PHASES_ACT
                     if страниц <= RepairOcrWorker.SHORT_DOC
                     else RepairOcrWorker.PHASES_ID)
            очередь.append(p)
            размеры.append(len(набор))
            этапы += list(набор)

        if not очередь:
            self._log("Ни один документ ремонта не удалось открыть", "error")
            return

        self._ocr_queue = очередь
        self._ocr_phase_sizes = размеры
        self._ocr_idx = 0
        self._ocr_phase_offset = 0
        self._ocr_hints = []
        self.seg_ocr.set_phases(этапы)
        self._log(f"Распознавание: документов в очереди {len(очередь)}", "info")
        self._ocr_next()

    def _ocr_next(self):
        """Запустить распознавание очередного документа очереди."""
        if self._ocr_idx >= len(self._ocr_queue):
            self._finish_ocr_queue()
            return
        from workers.repair_worker import RepairOcrWorker

        path = self._ocr_queue[self._ocr_idx]
        self._ocr_running = True
        w = RepairOcrWorker(path, parent=self)
        w.log.connect(lambda m: self._log(m, "info"))
        w.info.connect(self._log)
        w.result.connect(self._on_ocr)
        w.map_ready.connect(self._on_id_map)
        w.zone_hint.connect(lambda h, p=path: self._collect_hint(p, h))
        w.phase.connect(self._on_ocr_phase)
        w.finished_ok.connect(self._on_doc_ocr_finished)
        self._workers["ocr"] = w
        w.start()

    def _on_ocr_phase(self, index: int, name: str):
        """Этап текущего документа → сегмент общей полосы."""
        всего = len(self._ocr_queue)
        подпись = name
        if всего > 1 and name:
            подпись = f"Документ {self._ocr_idx + 1} из {всего} · {name}"
        self.seg_ocr.set_current(self._ocr_phase_offset + index, подпись)

    def _on_doc_ocr_finished(self, ok: bool):
        self._ocr_running = False
        if self._ocr_idx < len(self._ocr_phase_sizes):
            self._ocr_phase_offset += self._ocr_phase_sizes[self._ocr_idx]
        self._ocr_idx += 1
        self._ocr_next()

    def _collect_hint(self, path: str, hint):
        """Отложить кроп до конца очереди: сначала все документы, потом дата."""
        if hint:
            self._ocr_hints.append((path, hint))

    def _finish_ocr_queue(self):
        """
        Очередь пройдена: показать кроп с датой и поставить курсор в ячейку.

        Приоритет у «даты выполнения работ» из комплекта ИД: это то, что
        сотрудник вписывает руками. Если её нет (у пункта только акт), покажем
        гриф УТВЕРЖДАЮ — чтобы кроп в конце был всегда.
        """
        self.seg_ocr.finish()
        self._ocr_running = False
        if not self._ocr_hints:
            return
        путь, hint = next(
            ((p, h) for p, h in self._ocr_hints
             if h.get("зачем") == "дата выполнения работ"), self._ocr_hints[0])
        self._show_zone(путь, hint)

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

    def _show_zone(self, path: str, hint):
        """
        Показать сотруднику фрагмент, который OCR прочитать не может.

        Рукописные даты и штампы Windows OCR не распознаёт вовсе, поэтому
        программа открывает нужный кроп сама — иначе пришлось бы листать
        комплект на сотню страниц. Путь передаётся явно: кроп показывается в
        конце очереди, когда открытым в превью может быть уже другой документ.
        """
        if not hint or not path or not os.path.isfile(path):
            return
        try:
            from repair.pdf_preview import render_zone
            полоса = hint.get("полоса") or [0.6, 1.0]
            arr = render_zone(path, hint["page"] - 1, полоса[0], полоса[1],
                              rotation=hint.get("поворот", 0), scale=2.8)
            self._doc_path = path
            self._doc_page_idx = hint["page"] - 1
            self.doc_view._fitted = False
            self.doc_view.set_image(self._np_to_qimage(arr))
            подпись = hint.get("зачем", "фрагмент")
            self.lbl_pg.setText(f"стр. {hint['page']} · {подпись}")
            # Выделить документ в списке, чтобы было видно, откуда фрагмент.
            for i in range(self.list_docs.count()):
                if self.list_docs.item(i).toolTip() == path:
                    self.list_docs.blockSignals(True)
                    self.list_docs.setCurrentRow(i)
                    self.list_docs.blockSignals(False)
                    break
            if not hint.get("найдено"):
                self._log(f"Впишите «{подпись}» в реестр по этому фрагменту "
                          f"({os.path.basename(path)}, стр. {hint['page']})",
                          "wait")
                self._focus_date_cell()
        except Exception as e:
            self._log(f"Фрагмент не показан: {e}", "warn")

    def _focus_date_cell(self):
        """
        Поставить курсор в пустую ячейку «Дата выполнения» и открыть редактор.

        Кроп показан — остаётся набрать цифры; маска сама расставит точки.
        Ячейку с уже введённой датой не трогаем.
        """
        for r in range(self.tbl_registry.rowCount()):
            it = self.tbl_registry.item(r, REG_COL_DATE_DONE)
            if it is None:
                continue
            if it.text().strip():
                continue
            self.tbl_registry.setCurrentCell(r, REG_COL_DATE_DONE)
            self.tbl_registry.editItem(it)
            return

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
        self._filling_sections = True
        try:
            self.tbl_sections.setRowCount(len(sections))
            for r, s in enumerate(sections):
                vals = [str(s.number), s.element, s.thickness_nom, s.diameter,
                        " / ".join(m for m in s.measures if m), s.tmin,
                        "", "", ""]
                for c, val in enumerate(vals):
                    it = QTableWidgetItem(val)
                    if c not in NEW_COLS:
                        it.setFlags(it.flags() & ~Qt.ItemIsEditable)
                    self.tbl_sections.setItem(r, c, it)
                # строка листа — в данных первой ячейки
                self.tbl_sections.item(r, COL_NUM).setData(Qt.UserRole, s.row)
        finally:
            self._filling_sections = False
        self._apply_sections_filter()

    def _cell(self, row: int, col: int) -> str:
        it = self.tbl_sections.item(row, col)
        return it.text().strip() if it else ""

    def _on_section_edited(self, item):
        """
        Правка ячейки: замер подставляется из новой Тисп.

        Труба новая — мерить нечего, поэтому во все точки замеров идёт её
        номинал. Подставляем только в пустую ячейку: введённое сотрудником
        значение важнее вычисленного.
        """
        if self._filling_sections or item.column() != COL_NEW_T:
            return
        from repair.nominals import measure_from_thickness

        замер = measure_from_thickness(item.text())
        цель = self.tbl_sections.item(item.row(), COL_NEW_M)
        if замер and цель is not None and not цель.text().strip():
            self._filling_sections = True
            try:
                цель.setText(замер)
            finally:
                self._filling_sections = False

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
            эл = self.tbl_sections.item(r, COL_ELEM)
            мелкий = bool(эл) and any(эл.text().strip().lower().startswith(m)
                                      for m in мелкие)
            заполнен = any(self._cell(r, c) for c in NEW_COLS)
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
            # Если на схеме есть красное выделение — показываем сразу его
            # крупным планом: искать заменяемый участок глазами по всему листу
            # не нужно, а замер рядом с ним и есть ключ к номеру сечения.
            крупно = self._act_marker_view()
            if крупно is not None:
                arr, путь, стр, сколько = крупно
                # Числом отметок не хвастаемся: одна дуга распадается на
                # несколько пятен, и «мест: 6» при трёх настоящих врало бы.
                хвост = ("выделенные участки" if сколько > 1
                         else "выделенный участок")
                подпись = (f"Схема, {хвост} — {os.path.basename(путь)}, "
                           f"лист {стр}")
            else:
                from repair.pdf_preview import page_count, render_zone
                путь = файлы[0]
                # Схема — на последнем листе акта, над таблицей замены.
                стр = page_count(путь)
                arr = render_zone(путь, стр - 1, 0.0, 0.75, scale=2.0)
                подпись = (f"Схема дефектного участка: "
                           f"{os.path.basename(путь)}, лист {стр} "
                           f"(выделение не найдено — смотрите весь лист)")
            self.scheme_view._fitted = False
            self.scheme_view.set_image(self._np_to_qimage(arr))
            self.scheme_view.setVisible(True)
            self._log(подпись, "info")
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

        from repair.nominals import measure_from_thickness

        # снять прежние пометки
        for r in range(self.tbl_sections.rowCount()):
            it = self.tbl_sections.item(r, COL_NUM)
            if it is not None:
                it.setIcon(QIcon())
                it.setToolTip("")

        по_строке = {}
        for r in range(self.tbl_sections.rowCount()):
            it = self.tbl_sections.item(r, COL_NUM)
            if it is not None:
                по_строке[it.data(Qt.UserRole)] = r

        # Выделенный красным участок на схеме — третий, независимый признак.
        по_схеме = self._marker_sections()

        всего = подставлено = 0
        for s in suggest_sections(строки, сечения):
            self._log(describe(s), "done" if s.кандидаты else "warn")
            # Пересечение «что сказано в акте» и «что выделено на схеме» —
            # самый надёжный ответ: на п.2687 рядом с выделением три сечения,
            # и только одно из них отвод, как в таблице замены.
            точные = [c for c in s.кандидаты if c.number in по_схеме] or \
                     self._marker_match(s, по_схеме, сечения)
            for c in s.кандидаты:
                r = по_строке.get(c.row)
                if r is None:
                    continue
                всего += 1
                it = self.tbl_sections.item(r, COL_NUM)
                it.setIcon(dot_icon(_DOT["done" if s.однозначно else "warn"]))
                it.setToolTip(f"{s.строка} → Тисп {s.номинал}, "
                              f"диаметр {s.диаметр}")
                if s.однозначно and not any(self._cell(r, cc) for cc in NEW_COLS):
                    # Из строки замены известны и толщина, и диаметр — они
                    # раскладываются по своим колонкам, а замер выводится
                    # из толщины.
                    self._set_new_values(r, s.номинал, s.диаметр,
                                         measure_from_thickness(s.номинал))
                    подставлено += 1

            if len(точные) == 1:
                c = точные[0]
                r = по_строке.get(c.row)
                if r is not None:
                    it = self.tbl_sections.item(r, COL_NUM)
                    it.setIcon(dot_icon(_DOT["done"]))
                    it.setToolTip(f"{s.строка} → выделено на схеме")
                    self._log(
                        f"На схеме выделен участок: {s.элемент} → сечение "
                        f"{c.number} (было {c.diameter}×{c.thickness_nom}, "
                        f"замер {c.tmin})", "done")
                    if not any(self._cell(r, cc) for cc in NEW_COLS):
                        self._set_new_values(r, s.номинал, s.диаметр,
                                             measure_from_thickness(s.номинал))
                        подставлено += 1
        self._log(f"Подбор: помечено сечений {всего}, номинал подставлен "
                  f"в {подставлено}", "info")
        self._refresh_map_plan()

    # ── выделенный красным участок на схеме ──────────────────────────

    #: Радиус вокруг выделения, в котором ищутся подписи сечений. Подписи стоят
    #: не на трубе, а сбоку на выносках, поэтому ближайшая не обязательно та —
    #: берём круг и сужаем типом элемента из акта.
    MARKER_RADIUS = 260

    #: Масштаб рендера страницы акта при поиске выделения. Схема на листе
    #: нарисована мелко: при 3.0 полоса у половины актов не доживает до порога
    #: толщины, при 5.0 находится у всех пяти проверенных.
    ACT_RENDER_SCALE = 5.0

    #: Сколько пикселей вокруг выделения показывать в «Схеме из акта».
    SCHEME_ZOOM_RADIUS = 700

    def _marker_sections(self) -> set:
        """
        Номера сечений рядом с красным выделением на схеме.

        Источник только один — исходник ОТН из «Первички»: он попиксельно
        совпадает с рабочей картой, поэтому координаты переносятся без всякой
        подгонки. Схема внутри акта — другой чертёж, и связать её с картой
        можно лишь через подписанные рядом замеры; проверка на п.2763 показала,
        что OCR читает их неверно (6.9 → «9.6», настоящее 2.6 пропущено) и
        выдаёт чужое сечение. Поэтому схему акта программа не толкует, а
        показывает сотруднику — см. «Схема из акта».
        """
        return self._marker_from_original()

    def _marker_from_original(self) -> set:
        """Выделение на исходнике ОТН: координаты переносятся 1:1."""
        import numpy as np
        from PIL import Image
        from repair.highlight import find_markers, labels_near
        from repair.map_sheets import original_map_file

        найдено = set()
        for s in self._sheets:
            исходник = original_map_file(s["path"])
            if not os.path.isfile(исходник) or not s["labels"]:
                continue
            try:
                arr = np.array(Image.open(исходник).convert("RGB"))
                метки = find_markers(arr)
            except Exception as e:
                self._log(f"Исходник {os.path.basename(исходник)} не прочитан: {e}",
                          "warn")
                continue
            # Выделений бывает несколько: заменяют не один кусок, а участок.
            for m in метки:
                рядом = labels_near(s["labels"], m.center, self.MARKER_RADIUS)
                найдено |= {l.number for l in рядом}
                self._log(f"На исходнике {os.path.basename(исходник)} выделен "
                          f"участок; рядом сечения: "
                          f"{', '.join(str(l.number) for l in рядом) or '—'}",
                          "info")
        return найдено

    def _act_marker_view(self):
        """
        Найти выделенные участки на схеме акта и вернуть их крупным планом.

        Возвращает (массив, путь, страница, сколько выделений) или None.
        Программа схему не толкует — показывает: замер у выделения читается
        глазами за секунду, а OCR на нём ошибается (п.2763: 6.9 прочиталось
        как «9.6», настоящее 2.6 пропущено).

        Кроп берётся по всем выделениям сразу: у п.2688 их три, и показать
        одно значило бы спрятать два остальных.
        """
        from repair.highlight import find_markers, union_bbox
        from repair.pdf_preview import page_count, render_page

        docs = self._scan.get("docs")
        for путь in list(getattr(docs, "act_files", []) or []):
            try:
                страниц = page_count(путь)
            except Exception:
                continue
            for стр in range(страниц):
                try:
                    arr = render_page(путь, стр, scale=self.ACT_RENDER_SCALE)
                    метки = find_markers(arr)
                except Exception:
                    continue
                if not метки:
                    continue
                общая = union_bbox(метки)
                R = self.SCHEME_ZOOM_RADIUS
                h, w = arr.shape[:2]
                y1, y2 = max(0, общая.y1 - R), min(h, общая.y2 + R)
                x1, x2 = max(0, общая.x1 - R), min(w, общая.x2 + R)
                return arr[y1:y2, x1:x2], путь, стр + 1, len(метки)
        return None

    @staticmethod
    def _marker_match(s, по_схеме: set, сечения) -> list:
        """
        Сечения из круга выделения, подходящие по типу элемента из акта.

        Нужно, когда подбор по типоразмеру нашёл не то: у п.2687 акт ставит
        отвод Ø57 вместо Ø45, поиск по диаметру выдал два чужих отвода, а
        рядом с выделением лежат сечения 28 (труба), 29 (отвод) и 31 (переход)
        — по типу подходит ровно одно.
        """
        from repair.suggest import _similar, ELEMENT_MIN

        if not по_схеме or not s.элемент:
            return []
        return [sec for sec in сечения
                if sec.number in по_схеме
                and _similar(s.элемент, sec.element) >= ELEMENT_MIN]

    def _set_new_values(self, row: int, тисп: str, диаметр: str, замер: str):
        """Записать три поля правки разом, не дёргая автоподстановку замера."""
        self._filling_sections = True
        try:
            for col, val in ((COL_NEW_T, тисп), (COL_NEW_D, диаметр),
                             (COL_NEW_M, замер)):
                it = self.tbl_sections.item(row, col)
                if it is not None:
                    it.setText(val)
        finally:
            self._filling_sections = False

    def _collect_edits(self):
        """Три правых колонки → ({строка листа: SectionEdit}, ошибки)."""
        from repair.nominals import SectionEdit

        edits: dict[int, SectionEdit] = {}
        errors: list[str] = []
        for r in range(self.tbl_sections.rowCount()):
            тисп = self._cell(r, COL_NEW_T)
            диам = self._cell(r, COL_NEW_D)
            замер = self._cell(r, COL_NEW_M)
            if not (тисп or диам or замер):
                continue
            num = self._cell(r, COL_NUM)
            if замер:
                # Замер уходит в ячейки Excel числом — проверяем сразу, иначе
                # ошибка всплывёт уже во время записи в книгу.
                try:
                    float(замер.replace(",", "."))
                except ValueError:
                    errors.append(f"сечение {num}: замер «{замер}» — не число")
                    continue
            sheet_row = self.tbl_sections.item(r, COL_NUM).data(Qt.UserRole)
            edits[sheet_row] = SectionEdit(thickness=тисп, diameter=диам,
                                           measure=замер)
        return edits, errors

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
        from repair.nominals import apply_edits
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
            self.lbl_apply_state.setVisible(False)
            return

        records = self._registry_records()
        собранная = build_svedenia(records)
        now = self._scan.get("svedenia_now")
        # На ручном контроле (этап 3) отчёта ещё нет — его создаёт этап 5.
        # Это норма, а не ошибка: строка «Сведения о ремонте» вносится позже.
        есть_отчёт = bool(self._scan.get("report"))
        self._sved_generated = собранная
        self.txt_sved_now.setPlainText(
            now if now else ("(отчёт ещё не создан — появится на этапе 5)"
                             if not есть_отчёт else
                             "(строка в отчёте не найдена)"))
        # Правку сотрудника не затираем: пересборка предпросмотра случается на
        # каждом переключении вкладки.
        if self._sved_edited:
            self.lbl_sved.setText("Станет — изменено вручную")
        else:
            self.lbl_sved.setText("Станет")
            self._set_svedenia_text(собранная)
        svedenia = self._svedenia_text()

        edits, errors = self._collect_edits()
        for e in errors:
            self._log(e, "error")
        new, changes = apply_edits(sections, edits, by_row=True)

        # операции, сгруппированные по строке листа — для согласования с галочками
        ops_all = build_write_ops(sections, new)
        ops_by_row = {}
        for op in ops_all:
            ops_by_row.setdefault(op.row, []).append(op)

        self.tbl_preview.setRowCount(len(changes) + 1)
        # строка 0 — Word
        chk = QTableWidgetItem()
        chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        chk.setCheckState(Qt.Checked if есть_отчёт else Qt.Unchecked)
        self.tbl_preview.setItem(0, 0, chk)
        self.tbl_preview.setItem(0, 1, QTableWidgetItem("Word · таблица 1"))
        self.tbl_preview.setItem(
            0, 2, QTableWidgetItem(now if now else
                                   ("(не найдена)" if есть_отчёт
                                    else "(отчёта ещё нет)")))
        self.tbl_preview.setItem(
            0, 3, QTableWidgetItem(svedenia if есть_отчёт else
                                   f"{svedenia}\n\n(будет внесено, когда "
                                   f"отчёт создан на этапе 5)"))

        for i, ch in enumerate(changes, start=1):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Checked)
            self.tbl_preview.setItem(i, 0, chk)
            self.tbl_preview.setItem(i, 1, QTableWidgetItem(f"Импорт · {ch.where}"))
            self.tbl_preview.setItem(i, 2, QTableWidgetItem(ch.before))
            self.tbl_preview.setItem(i, 3, QTableWidgetItem(ch.after))
        self.tbl_preview.resizeRowsToContents()

        # Сопоставление галочка → операции по номеру строки листа. Порядок тот
        # же, в каком apply_edits собирал changes: обход исходных сечений.
        sheet_rows_in_order = [s.row for s in sections
                               if s.row in edits and not edits[s.row].is_empty()]
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
        # «Не записано» показываем не по факту наличия строк, а по расхождению
        # с файлами: после записи книга и отчёт уже содержат нужное, и плашка
        # честно гаснет. Пока отчёта нет, строку сравнивать не с чем — иначе
        # плашка горела бы всегда и о записи в Excel ничего не говорила.
        строка_разошлась = (есть_отчёт and bool(svedenia)
                            and svedenia != " ".join((now or "").split()))
        реестр_не_заведён = bool(records) and not self._scan.get("registry_existed")
        self._set_apply_state(len(changes),
                              bool(changes) or строка_разошлась
                              or реестр_не_заведён,
                              есть_отчёт)

    def _map_state_text(self) -> str:
        """
        Строка о состоянии карты для плашки предпросмотра.

        Карта живёт отдельно: её сохраняет своя кнопка, и «Применить выбранное»
        её не трогает. Порядок нажатий значения не имеет, но человеку об этом
        должно быть сказано — иначе непонятно, не потерялись ли правки карты.
        """
        if not self._sheets:
            return ""
        не_сохранено = self._total_edits()
        if не_сохранено:
            return (f" Карта: {не_сохранено} правок не сохранено — вернитесь на "
                    f"вкладку «Карта» и нажмите «Сохранить карту».")
        план = len(self._planned_edits())
        if план:
            return (f" Карта сохраняется отдельно своей кнопкой; отмеченных "
                    f"сечений: {план}.")
        return ""

    def _set_apply_state(self, n_changes: int, надо_писать: bool,
                         есть_отчёт: bool = True):
        """Плашка «записано / ещё не записано» под таблицей предпросмотра."""
        lbl = self.lbl_apply_state
        записано = (self._applied_summary
                    if self._applied_folder == self._point_folder() else "")
        if записано:
            записано += self._map_state_text()
        if надо_писать:
            части = []
            if n_changes:
                части.append(f"сечений к правке: {n_changes}")
            if not есть_отчёт:
                части.append("строка отчёта — после этапа 5")
            хвост = f" ({'; '.join(части)})" if части else ""
            lbl.setText(
                f"Изменения ещё не записаны в файлы пункта{хвост}. "
                f"Нажмите «Применить выбранное» — до этого листы «Импорт» и "
                f"лист «Ремонт» остаются прежними, и следующие этапы соберут "
                f"отчёт по старым замерам.")
            lbl.setProperty("state", "wait")
        elif записано:
            lbl.setText(записано)
            lbl.setProperty("state", "done")
        else:
            lbl.setVisible(False)
            return
        lbl.setVisible(True)
        lbl.style().unpolish(lbl)
        lbl.style().polish(lbl)

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
            self._applied_folder = self._point_folder()
            куда = ["листы «Импорт» и «Импорт (ПР)»", "лист «Ремонт»"]
            хвост = ""
            if self._scan.get("report"):
                куда.append("строка «Сведения о ремонте» в отчёте")
            else:
                # Отчёт создаёт этап 5 — обещать записанную строку нельзя.
                хвост = (" Строка «Сведения о ремонте» будет внесена, когда "
                         "отчёт создан на этапе 5 — вернитесь сюда и нажмите "
                         "«Применить выбранное» ещё раз.")
            self._applied_summary = (
                f"Записано в файлы пункта: {', '.join(куда)}. Бэкап и журнал "
                f"созданы.{хвост}")
            self._log("Готово. Перечитываю пункт...", "done")
            self._start_scan(run_ocr=False)

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
        self._start_scan(run_ocr=False)

    # ═════════════════════════════ Карта ══════════════════════════════

    @staticmethod
    def _find_map_paths(folder: str) -> list[str]:
        """
        Рабочие листы коррозионной карты пункта, по порядку.

        Карта бывает разбита на 2–3 листа («2704_1.png», «2704_2.png»), а рядом
        лежат исходники ОТН с лишней точкой в имени («2704._1.png») — их правка
        запрещена. Отбор и порядок — в repair.map_sheets, там же он и покрыт
        тестами на реальных наборах имён.
        """
        from repair.detect import _find_subfolder
        from repair.map_sheets import working_map_files

        perv = _find_subfolder(folder, "Первичка")
        if not perv:
            return []
        return [os.path.join(perv, f)
                for f in working_map_files(sorted(os.listdir(perv)))]

    @staticmethod
    def _np_to_qimage(arr) -> QImage:
        import numpy as np
        a = np.ascontiguousarray(arr)
        h, w = a.shape[:2]
        return QImage(a.data, w, h, 3 * w, QImage.Format_RGB888).copy()

    def _map_autoload(self):
        """
        Подгрузить карту сразу при проверке пункта.

        Пути к листам выводятся из содержимого «Первички», так что просить
        сотрудника нажимать «Загрузить карту» незачем. Если карты нет — молча
        ничего не делаем: у пункта её может и не быть.
        """
        folder = self._point_folder()
        if not os.path.isdir(folder) or not self._find_map_paths(folder):
            self.lbl_map_status.setText(
                "Рабочая карта (Первичка/NNNN.png) в пункте не найдена")
            return
        self._map_load()

    def _map_load(self):
        folder = self._point_folder()
        if not os.path.isdir(folder):
            self._log("Укажите папку пункта", "error")
            return
        пути = self._find_map_paths(folder)
        if not пути:
            self._log("Рабочая карта (Первичка/NNNN.png) не найдена", "error")
            return
        import numpy as np
        from PIL import Image
        from repair.map_sheets import sheet_number

        self._sheets = []
        for p in пути:
            self._sheets.append({
                "path": p,
                "num": sheet_number(p),
                "arr": np.array(Image.open(p).convert("RGB")),
                "labels": [],
                "by_number": {},
                "edits": [],       # [(номер, значение, рамка подписи)]
            })
        self._sheet_idx = 0
        self._map_cycle = {}
        self._map_shown = None
        self._map_dirty = False
        self.list_map_edits.clear()
        self._build_sheet_buttons()
        self._show_sheet(0, fit=True)
        self.btn_map_undo.setEnabled(False)
        self.btn_map_save.setEnabled(False)
        листы = ", ".join(os.path.basename(s["path"]) for s in self._sheets)
        self._log(f"Карта загружена: {листы}"
                  + (f" ({len(self._sheets)} листа)" if len(self._sheets) > 1
                     else ""), "info")
        self._map_scan_labels()

    # ── листы карты ──────────────────────────────────────────────────

    def _build_sheet_buttons(self):
        """Ряд «Лист 1 / Лист 2 / …»; при одном листе ряд скрыт."""
        for b in self._sheet_buttons:
            self.row_sheets.removeWidget(b)
            b.deleteLater()
        self._sheet_buttons = []
        if len(self._sheets) < 2:
            self.wid_sheets.setVisible(False)
            return
        for i, s in enumerate(self._sheets):
            b = QPushButton(f"Лист {s['num']}")
            b.setObjectName("segment")
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _=False, idx=i: self._show_sheet(idx))
            self.row_sheets.addWidget(b)
            self._sheet_buttons.append(b)
        self.row_sheets.addStretch()
        self.wid_sheets.setVisible(True)

    def _sheet(self, idx: int | None = None) -> dict | None:
        i = self._sheet_idx if idx is None else idx
        return self._sheets[i] if 0 <= i < len(self._sheets) else None

    def _show_sheet(self, idx: int, fit: bool = False):
        """Показать лист карты на экране."""
        s = self._sheet(idx)
        if s is None:
            return
        self._sheet_idx = idx
        for i, b in enumerate(self._sheet_buttons):
            b.setChecked(i == idx)
        if fit:
            self.map_view._fitted = False
        self.map_view.set_image(self._np_to_qimage(s["arr"]))
        self._update_map_status()

    def _update_map_status(self):
        s = self._sheet()
        if s is None:
            return
        сечений = len(s["by_number"])
        хвост = (f"распознано сечений {сечений}" if сечений
                 else "подписи ещё не разобраны")
        лист = f"лист {s['num']} · " if len(self._sheets) > 1 else ""
        self.lbl_map_status.setText(
            f"{лист}{os.path.basename(s['path'])} · {хвост}")

    # ── подписи сечений на карте и план правок ───────────────────────

    def _map_scan_labels(self):
        """Найти подписи сечений на всех листах карты (фоном)."""
        if not self._sheets:
            return
        from workers.repair_worker import MapLabelsWorker

        for s in self._sheets:
            s["labels"] = []
            s["by_number"] = {}
        self.seg_map.set_phases([f"Лист {s['num']}" for s in self._sheets])
        self.seg_map.set_current(0)
        self.lbl_map_status.setText("Ищу подписи сечений на карте…")
        # Массивы отдаём без копий: find_labels только читает, а копия каждого
        # листа — это лишние десятки мегабайт.
        w = MapLabelsWorker([s["arr"] for s in self._sheets], parent=self)
        w.log.connect(lambda m: self._log(m, "info"))
        w.info.connect(self._log)
        w.sheet_done.connect(self._on_sheet_labels)
        w.finished_ok.connect(lambda ok: self.seg_map.finish())
        self._workers["map_labels"] = w
        w.start()

    def _on_sheet_labels(self, idx: int, labels):
        from repair.map_labels import labels_by_number
        from repair.map_sheets import MIN_SECTION_LABELS

        s = self._sheet(idx)
        if s is None:
            return
        s["labels"] = labels or []
        s["by_number"] = labels_by_number(s["labels"])
        # Страховка от правки исходника ОТН: у него разметки толщинометрии нет,
        # подписей находится 0–1, тогда как у рабочего листа — от 66.
        if len(s["by_number"]) < MIN_SECTION_LABELS:
            s["readonly"] = True
            self._log(f"Лист {s['num']} ({os.path.basename(s['path'])}): "
                      f"подписей сечений почти нет — похоже, это исходник ОТН, "
                      f"править его не буду", "warn")
        self.seg_map.set_current(idx + 1)
        if idx == self._sheet_idx:
            self._update_map_status()
        self._refresh_map_plan()

    def _sheets_by_number(self) -> dict:
        """{индекс листа: множество номеров сечений} — для подсказки по соседям."""
        return {i: set(s["by_number"]) for i, s in enumerate(self._sheets)
                if not s.get("readonly")}

    def _places_for(self, номер: int) -> list[tuple[int, object]]:
        """Все места сечения на карте: [(индекс листа, подпись), …]."""
        места = []
        for i, s in enumerate(self._sheets):
            if s.get("readonly"):
                continue
            for l in s["by_number"].get(номер, []):
                места.append((i, l))
        return места

    def _planned_edits(self) -> list[tuple[int, str]]:
        """
        Что предстоит поправить на карте: (номер сечения, новое значение).

        На схеме в скобках стоит замер, а не Тисп, поэтому берётся колонка
        «Новый замер»; если она пуста — значение выводится из новой Тисп.
        """
        from repair.nominals import measure_from_thickness

        out = []
        for r in range(self.tbl_sections.rowCount()):
            текст = (self._cell(r, COL_NEW_M)
                     or measure_from_thickness(self._cell(r, COL_NEW_T)))
            if not текст:
                continue
            номер = self._cell(r, COL_NUM)
            if not номер.isdigit():
                continue
            out.append((int(номер), текст))
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

        from repair.map_sheets import likely_sheet

        много_листов = len(self._sheets) > 1
        по_листам = self._sheets_by_number() if много_листов else {}
        однозначных = 0
        for номер, значение in план:
            места = self._places_for(номер)
            листы = sorted({i for i, _ in места})
            if len(места) == 1:
                цвет = "done"
                хвост = "подпись найдена"
                if много_листов:
                    хвост = f"лист {self._sheets[листы[0]]['num']} · " + хвост
                однозначных += 1
            elif len(места) > 1:
                цвет = "warn"
                разбивка = ", ".join(
                    f"лист {self._sheets[i]['num']} — "
                    f"{sum(1 for j, _ in места if j == i)}" for i in листы)
                хвост = f"мест: {len(места)} ({разбивка}) — выберите"
                if len(листы) > 1:
                    # Сечения нумеруются вдоль трубопровода: лист, где лежат
                    # соседние номера, почти наверняка держит и это сечение.
                    # Подсказываем, но не проставляем — решает сотрудник.
                    вероятный = likely_sheet(номер, по_листам)
                    if вероятный is not None:
                        хвост += (f"; вероятно лист "
                                  f"{self._sheets[вероятный]['num']} (по соседям)")
            else:
                цвет, хвост = "wait", "подпись не найдена, правьте кликом"
            item = QListWidgetItem(f"сечение {номер} → {значение} · {хвост}")
            item.setIcon(dot_icon(_DOT[цвет]))
            item.setData(Qt.UserRole, (номер, значение))
            self.list_map_plan.addItem(item)
        self.btn_map_apply_unique.setEnabled(однозначных > 0 and
                                             bool(self._sheets))

    def _map_plan_clicked(self, item):
        """Перейти к подписи сечения; при нескольких местах — к следующему."""
        данные = item.data(Qt.UserRole)
        if not данные or not self._sheets:
            return
        номер, значение = данные
        места = self._places_for(номер)
        if not места:
            self._log(f"Сечение {номер}: подпись на карте не распознана — "
                      f"кликните по ней сами", "warn")
            return
        i = self._map_cycle.get(номер, -1) + 1
        if i >= len(места):
            i = 0
        self._map_cycle[номер] = i
        лист, подпись = места[i]
        if лист != self._sheet_idx:
            self._show_sheet(лист)
        self.map_view.show_at(*подпись.point, box=подпись.bbox)
        self._map_shown = (номер, значение, лист, подпись)
        self.btn_map_apply_one.setEnabled(True)
        уточнение = f" (место {i + 1} из {len(места)})" if len(места) > 1 else ""
        где = (f", лист {self._sheets[лист]['num']}"
               if len(self._sheets) > 1 else "")
        self._log(f"Сечение {номер}{уточнение}{где}: подпись показана на карте",
                  "info")

    @staticmethod
    def _map_value_text(значение: str) -> str:
        """Значение в том виде, как оно стоит на схеме: «(6.0)»."""
        text = (значение or "").strip()
        if not text.startswith("("):
            text = f"({text})" if not text.endswith(")") else f"({text}"
        return text

    def _map_write(self, номер: int, значение: str, лист: int, подпись) -> bool:
        """Записать новое значение в подпись сечения на указанном листе."""
        from repair.map_editor import render_value, render_value_in

        s = self._sheet(лист)
        if s is None or s.get("readonly"):
            return False
        значение = self._map_value_text(значение)
        arr = s["arr"]
        # Кроп «до» снимаем с массива перед записью: держать копию каждого
        # листа только ради него — это десятки лишних мегабайт.
        было = arr.copy() if arr.size < 60_000_000 else None
        try:
            if подпись.value_bbox is not None:
                # Рамка значения известна из распознанной подписи — берём её.
                # Выращивать рамку от точки нельзя: рядом с подписью стоит
                # красный крестик замера, и рамка «раздувается» на него.
                rep = render_value_in(arr, подпись.value_bbox,
                                      значение, label=подпись.bbox)
            else:
                rep = render_value(arr, подпись.point, значение)
        except Exception as e:
            self._log(f"Сечение {номер}: правка не удалась — {e}", "error")
            return False

        s["edits"].append((номер, значение, подпись))
        self._map_dirty = True
        где = f" · лист {s['num']}" if len(self._sheets) > 1 else ""
        item = QListWidgetItem(f"сечение {номер} → {значение}{где}")
        item.setIcon(dot_icon(_DOT["done"]))
        item.setData(Qt.UserRole, лист)
        self.list_map_edits.addItem(item)
        if было is not None:
            self.lbl_map_before.setPixmap(self._crop_pixmap(было, rep["label"]))
        self.lbl_map_after.setPixmap(self._crop_pixmap(arr, rep["label"]))
        if лист != self._sheet_idx:
            self._show_sheet(лист)
        else:
            self.map_view.set_image(self._np_to_qimage(arr))
        self.map_view.show_at(*подпись.point, box=rep["label"])
        self.btn_map_undo.setEnabled(True)
        self.btn_map_save.setEnabled(True)
        return True

    def _map_apply_shown(self):
        if not self._map_shown:
            return
        номер, значение, лист, подпись = self._map_shown
        if self._map_write(номер, значение, лист, подпись):
            self._log(f"Сечение {номер}: на карте записано {значение}", "done")

    def _map_apply_unique(self):
        """
        Проставить значения там, где подпись сечения на карте одна.

        Проход идёт сразу по всем листам: замена одного акта нередко попадает
        на два листа, и заставлять открывать каждый и жать кнопку заново
        незачем. Места с несколькими одинаковыми сечениями намеренно не
        трогаются: какое из них заменено, видно только по схеме — решает
        сотрудник.
        """
        сделано, отложено = 0, 0
        по_листам: dict[int, int] = {}
        for номер, значение in self._planned_edits():
            места = self._places_for(номер)
            if len(места) != 1:
                отложено += 1
                continue
            лист, подпись = места[0]
            if self._map_write(номер, значение, лист, подпись):
                сделано += 1
                по_листам[лист] = по_листам.get(лист, 0) + 1
        разбивка = ", ".join(
            f"лист {self._sheets[i]['num']} — {n}"
            for i, n in sorted(по_листам.items())) if len(self._sheets) > 1 else ""
        self._log(f"Карта: проставлено {сделано}"
                  + (f" ({разбивка})" if разбивка else "")
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
        s = self._sheet()
        if s is None or s.get("readonly"):
            return
        from PySide6.QtWidgets import QInputDialog
        from repair.map_editor import label_bbox_at, find_value_span

        arr = s["arr"]
        xy = (int(x), int(y))
        current = ""
        try:
            lb = label_bbox_at(arr, *xy)
            if lb is not None and find_value_span(arr, lb) is not None:
                self.lbl_map_before.setPixmap(self._crop_pixmap(arr, lb))
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
        """Ручная правка по клику: подпись ищется на показанном листе."""
        from repair.map_editor import render_value

        s = self._sheet()
        if s is None:
            return
        arr = s["arr"]
        before = arr.copy()
        try:
            rep = render_value(arr, label_xy, text)
        except ValueError as e:
            s["arr"] = before
            self._log(f"Карта: {e}", "error")
            self.lbl_map_status.setText(f"Не получилось: {e}. "
                                        f"Кликните по подписи сечения")
            return
        lb = rep["label"]
        s["edits"].append((None, text, label_xy))
        self._map_dirty = True
        self.lbl_map_before.setPixmap(self._crop_pixmap(before, lb))
        self.lbl_map_after.setPixmap(self._crop_pixmap(arr, lb))
        где = f" · лист {s['num']}" if len(self._sheets) > 1 else ""
        item = QListWidgetItem(
            f"{rep['text']} · Arial {rep['font_size']} · "
            f"цвет RGB{rep['color']}{где}")
        item.setIcon(dot_icon(_DOT["done"]))
        item.setData(Qt.UserRole, self._sheet_idx)
        self.list_map_edits.addItem(item)
        self.map_view.set_image(self._np_to_qimage(arr))
        self.btn_map_undo.setEnabled(True)
        self.btn_map_save.setEnabled(True)
        self.lbl_map_status.setText(
            f"Правок: {self._total_edits()}. Проверьте «до/после» и кликните "
            f"по следующей подписи или сохраните карту")

    def _total_edits(self) -> int:
        return sum(len(s["edits"]) for s in self._sheets)

    def _map_undo(self):
        """
        Отменить последнюю правку — на том листе, где она была сделана.

        Лист перечитывается с диска и правки повторяются: держать копию
        исходного массива каждого листа только ради отката слишком дорого,
        а сама операция редкая.
        """
        item = self.list_map_edits.item(self.list_map_edits.count() - 1)
        if item is None:
            return
        лист = item.data(Qt.UserRole) or 0
        s = self._sheet(лист)
        if s is None or not s["edits"]:
            return
        import numpy as np
        from PIL import Image
        from repair.map_editor import render_value, render_value_in

        s["edits"].pop()
        self.list_map_edits.takeItem(self.list_map_edits.count() - 1)
        s["arr"] = np.array(Image.open(s["path"]).convert("RGB"))
        for номер, значение, привязка in s["edits"]:
            try:
                if номер is None:
                    render_value(s["arr"], привязка, значение)
                elif привязка.value_bbox is not None:
                    render_value_in(s["arr"], привязка.value_bbox, значение,
                                    label=привязка.bbox)
                else:
                    render_value(s["arr"], привязка.point, значение)
            except Exception:
                pass
        if лист == self._sheet_idx:
            self.map_view.set_image(self._np_to_qimage(s["arr"]))
        else:
            self._show_sheet(лист)
        есть = self._total_edits() > 0
        self._map_dirty = есть
        self.btn_map_undo.setEnabled(есть)
        self.btn_map_save.setEnabled(есть)
        self._log("Последняя правка карты отменена", "info")

    def _map_save(self):
        """
        Сохранить все изменённые листы: один бэкап, одна запись в журнале.

        Замена одного акта нередко ложится на два листа, и разносить это на
        два сохранения значило бы плодить бэкапы одного и того же события.
        """
        изменённые = [s for s in self._sheets if s["edits"]]
        if not изменённые:
            return
        from PIL import Image
        from repair.safety import make_backup, append_journal

        folder = self._point_folder()
        b = make_backup(folder, [s["path"] for s in изменённые])
        строки = []
        for s in изменённые:
            Image.fromarray(s["arr"]).save(s["path"])
            строки.append(f"лист {s['num']} ({os.path.basename(s['path'])}): "
                          f"правок {len(s['edits'])}")
        всего = self._total_edits()
        append_journal(folder, "Правка коррозионной карты",
                       строки + [f"бэкап: {b}"])
        листов = f" на {len(изменённые)} листах" if len(изменённые) > 1 else ""
        self._log(f"Карта сохранена ({всего} правок{листов}), "
                  f"бэкап и журнал созданы", "done")
        for s in изменённые:
            s["edits"] = []
        self._map_dirty = False
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
