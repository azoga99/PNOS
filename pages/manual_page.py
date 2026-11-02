# -*- coding: utf-8 -*-
"""
Страница «Ручной контроль» (шаг 3): форма паспорта и встроенный модуль «Ремонт».

Примесь к MainWindow: методы работают с тем же self, что и окно. Общее
для всех страниц (_get_local_root, _add_activity, _switch_page, _workers)
живёт в main_window.py.
"""

import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QProgressBar,
    QFrame, QLineEdit, QMessageBox, QStackedWidget, QListWidget,
    QListWidgetItem, QScrollArea,
)

from config import CONFIG
from pages.repair_page import RepairPage
from widgets import dot_icon, STATUS_COLORS


class ManualPageMixin:
    """Страница «Ручной контроль» (шаг 3): форма паспорта и встроенный модуль «Ремонт»."""

    # ═══════════════════════════════════════════════════════════════
    # Страница «Ручной контроль» (шаг 3)
    # ═══════════════════════════════════════════════════════════════

    def _make_date_edit(self):
        """Поле даты с жёсткой маской дд.мм.гггг (без календаря)."""
        de = QLineEdit()
        de.setInputMask("00.00.0000;_")
        de.setCursorPosition(0)
        return de

    def _date_edit_value(self, de) -> str:
        """Строка дд.мм.гггг или '' если поле пустое (одни подчёркивания)."""
        s = de.text().strip()
        if not any(c.isdigit() for c in s):
            return ""
        return s

    def _set_date_edit(self, de, value: str):
        de.setText((value or "").strip())
        de.setCursorPosition(0)

    @staticmethod
    def _date_ok_or_empty(s: str) -> bool:
        """Пустая строка — ок; иначе должна быть существующая дата дд.мм.гггг."""
        s = (s or "").strip()
        if not s:
            return True
        from PySide6.QtCore import QDate
        d = QDate.fromString(s, "dd.MM.yyyy")
        return d.isValid() and 1900 <= d.year() <= 2100

    def _form_field(self, label_text: str, widget, required: bool = False, note: str = ""):
        block = QWidget()
        v = QVBoxLayout(block)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(4)
        head = QHBoxLayout()
        head.setSpacing(4)
        lbl = QLabel(label_text)
        lbl.setObjectName("form_field_label")
        head.addWidget(lbl)
        if required:
            star = QLabel("*")
            star.setObjectName("form_field_req")
            head.addWidget(star)
        head.addStretch()
        v.addLayout(head)
        v.addWidget(widget)
        if note:
            n = QLabel(note)
            n.setObjectName("field_note")
            n.setWordWrap(True)
            v.addWidget(n)
        return block

    def _setup_manual_page(self):
        from widgets import NoWheelComboBox as QComboBox
        from PySide6.QtGui import QIntValidator

        self.manual_drafts = {}
        self.manual_current = None
        self._manual_loading = False

        outer = QVBoxLayout(self.page_manual)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(10)

        head_row = QHBoxLayout()
        head_col = QVBoxLayout()
        head_col.setSpacing(2)
        head_col.addWidget(self._page_title("Ручной контроль"))
        sub = QLabel("Заполните данные паспорта для каждого пункта — они будут записаны "
                     "на лист «паспорт» автоматически.")
        sub.setObjectName("card_description")
        sub.setWordWrap(True)
        head_col.addWidget(sub)
        head_row.addLayout(head_col, 1)
        btn_back = QPushButton("К процессу")
        btn_back.setObjectName("ghost_button")
        btn_back.setCursor(Qt.PointingHandCursor)
        btn_back.clicked.connect(lambda: self._switch_page(self.PAGE_PROCESS))
        head_row.addWidget(btn_back, 0, Qt.AlignTop)
        outer.addLayout(head_row)

        split = QHBoxLayout()
        split.setSpacing(12)

        left = QVBoxLayout()
        left.setSpacing(6)
        lbl_pts = QLabel("Пункты")
        lbl_pts.setObjectName("section_label")
        left.addWidget(lbl_pts)
        self.manual_point_list = QListWidget()
        self.manual_point_list.setObjectName("point_list")
        self.manual_point_list.setMinimumWidth(230)
        self.manual_point_list.setMaximumWidth(280)
        self.manual_point_list.currentItemChanged.connect(self._on_manual_point_selected)
        left.addWidget(self.manual_point_list, 1)
        legend = QLabel(
            f'<span style="color:{STATUS_COLORS["done"]}">●</span> заполнен&nbsp;&nbsp;'
            f'<span style="color:{STATUS_COLORS["idle"]}">●</span> не заполнен&nbsp;&nbsp;'
            f'<span style="color:{STATUS_COLORS["wait"]}">●</span> пропущен&nbsp;&nbsp;'
            f'<span style="color:{STATUS_COLORS["error"]}">●</span> не корректный')
        legend.setObjectName("field_note")
        legend.setWordWrap(True)
        left.addWidget(legend)
        split.addLayout(left)

        form_scroll = QScrollArea()
        form_scroll.setWidgetResizable(True)
        form_scroll.setFrameShape(QFrame.NoFrame)
        form_card = QFrame()
        form_card.setObjectName("form_card")
        form = QVBoxLayout(form_card)
        form.setContentsMargins(20, 16, 20, 16)
        form.setSpacing(12)

        self.manual_form_title = QLabel("Выберите пункт слева")
        self.manual_form_title.setObjectName("form_title")
        form.addWidget(self.manual_form_title)

        f = {}
        f["purpose"] = QLineEdit()
        f["purpose"].setPlaceholderText("транспортировки …")
        form.addWidget(self._form_field(
            "Предназначение трубопровода (B14)", f["purpose"], required=True,
            note="Начните с «транспортировки», далее — конкретное вещество "
                 "(напр. «транспортировки рафината растворителя»)."))

        f["commission"] = QLineEdit()
        f["commission"].setValidator(QIntValidator(1000, 2100, self))
        f["commission"].setMaxLength(4)
        f["commission"].setPlaceholderText("гггг")
        f["commission"].setMaximumWidth(120)
        form.addWidget(self._form_field(
            "Год ввода в эксплуатацию (B19)", f["commission"], required=False,
            note="Подставляется из паспорта; если там года нет — впишите сам. "
                 "Этот год печатается в отчёте: «введен в эксплуатацию с … г»."))

        import user_data
        f["specialist"] = QComboBox()
        f["specialist"].addItem("— выберите —", "")
        for name in user_data.get_list("specialists"):
            f["specialist"].addItem(name, name)
        form.addWidget(self._form_field("Специалист (B22)", f["specialist"], required=True))

        f["date"] = self._make_date_edit()
        form.addWidget(self._form_field("Дата (B33)", f["date"], required=True,
                                        note="Введите дату цифрами в формате дд.мм.гггг."))

        sep = QLabel("Данные предыдущей ЭПБ")
        sep.setObjectName("section_label")
        form.addWidget(sep)

        f["org"] = QComboBox()
        f["org"].addItem("— не выбрано —", "")
        for org in user_data.get_list("epb_orgs"):
            f["org"].addItem(org, org)
        f["org"].currentIndexChanged.connect(self._update_epb_done_label)
        form.addWidget(self._form_field(
            "Организация (B52)", f["org"], required=False,
            note="Необязательно. Если экспертиза ранее не проводилась — оставьте пустым."))

        self.manual_epb_done = QLabel("Проводилась ранее (B51): нет")
        self.manual_epb_done.setObjectName("field_note")
        form.addWidget(self.manual_epb_done)

        f["reg_num"] = QLineEdit()
        f["reg_num"].setPlaceholderText("Рег. номер Ростехнадзора или внутренний")
        form.addWidget(self._form_field("Регистрационный номер (B53)", f["reg_num"], required=False))

        f["epb_date"] = self._make_date_edit()
        form.addWidget(self._form_field(
            "Дата предыдущей ЭПБ (B54)", f["epb_date"], required=False,
            note="Формат дд.мм.гггг. Если даты нет — оставьте поле пустым."))

        f["sheets"] = QLineEdit()
        f["sheets"].setValidator(QIntValidator(0, 100000, self))
        f["sheets"].setPlaceholderText("Число листов заключения")
        form.addWidget(self._form_field(
            "Количество листов (B55)", f["sheets"], required=False,
            note="Лист уведомления Ростехнадзора не учитывается."))

        # ── Раздел «Ремонт» (появляется у пунктов с ремонтом) ──
        self.manual_repair_banner = QFrame()
        self.manual_repair_banner.setObjectName("settings_panel")
        rb = QHBoxLayout(self.manual_repair_banner)
        rb.setContentsMargins(14, 10, 14, 10)
        self.manual_repair_hint = QLabel("")
        self.manual_repair_hint.setObjectName("hint")
        self.manual_repair_hint.setWordWrap(True)
        rb.addWidget(self.manual_repair_hint, 1)
        self.manual_btn_repair = QPushButton("Внести данные по ремонту")
        self.manual_btn_repair.setObjectName("card_button")
        self.manual_btn_repair.setCursor(Qt.PointingHandCursor)
        self.manual_btn_repair.clicked.connect(self._open_repair_for_current)
        rb.addWidget(self.manual_btn_repair, 0)
        # Модуль — не единственный путь: пункт можно оформить как раньше,
        # руками в Excel, Word и графическом редакторе. Кнопка ничего не пишет
        # в артефакты, она снимает пункт с модуля и открывает его папку.
        self.manual_btn_repair_old = QPushButton("Оформить по старому")
        self.manual_btn_repair_old.setObjectName("ghost_button")
        self.manual_btn_repair_old.setCursor(Qt.PointingHandCursor)
        self.manual_btn_repair_old.clicked.connect(self._repair_old_way)
        rb.addWidget(self.manual_btn_repair_old, 0)
        form.addWidget(self.manual_repair_banner)
        self.manual_repair_banner.setVisible(False)

        form.addStretch()
        self.manual_fields = f
        form_scroll.setWidget(form_card)
        split.addWidget(form_scroll, 1)

        # Обычный вид (список + форма) и полноширинный вид ремонта — во
        # внутреннем стеке, чтобы ремонт открывался на всю ширину.
        normal_view = QWidget()
        nv = QVBoxLayout(normal_view)
        nv.setContentsMargins(0, 0, 0, 0)
        nv.addLayout(split, 1)

        repair_view = QWidget()
        rv = QVBoxLayout(repair_view)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(6)
        rvh = QHBoxLayout()
        btn_rep_back = QPushButton("← Назад к паспорту")
        btn_rep_back.setObjectName("ghost_button")
        btn_rep_back.setCursor(Qt.PointingHandCursor)
        btn_rep_back.clicked.connect(self._close_repair_view)
        rvh.addWidget(btn_rep_back, 0, Qt.AlignLeft)
        self.manual_repair_title = QLabel("")
        self.manual_repair_title.setObjectName("form_title")
        rvh.addWidget(self.manual_repair_title, 1)
        rv.addLayout(rvh)
        self.manual_repair_panel = RepairPage()
        self.manual_repair_panel.set_embedded(True)
        rv.addWidget(self.manual_repair_panel, 1)

        self.manual_stack = QStackedWidget()
        self.manual_stack.addWidget(normal_view)     # индекс 0
        self.manual_stack.addWidget(repair_view)     # индекс 1
        outer.addWidget(self.manual_stack, 1)

        # Кнопки действий
        btn_row = QHBoxLayout()
        self.manual_btn_write = QPushButton("Записать всё в Excel")
        self.manual_btn_write.setObjectName("card_button")
        self.manual_btn_write.setMinimumWidth(200)
        self.manual_btn_write.setCursor(Qt.PointingHandCursor)
        self.manual_btn_write.clicked.connect(self._start_manual_stage3)

        btn_skip = QPushButton("Пропустить пункт")
        btn_skip.setObjectName("ghost_button")
        btn_skip.setCursor(Qt.PointingHandCursor)
        btn_skip.clicked.connect(self._skip_current_point)

        btn_excel = QPushButton("Открыть в Excel")
        btn_excel.setObjectName("ghost_button")
        btn_excel.setCursor(Qt.PointingHandCursor)
        btn_excel.clicked.connect(self._open_current_in_excel)

        self.manual_btn_stop = QPushButton("Остановить")
        self.manual_btn_stop.setObjectName("stop_button")
        self.manual_btn_stop.setCursor(Qt.PointingHandCursor)
        self.manual_btn_stop.setEnabled(False)
        self.manual_btn_stop.clicked.connect(self._force_stop_active_stage)

        btn_row.addWidget(self.manual_btn_write)
        btn_row.addWidget(btn_skip)
        btn_row.addWidget(btn_excel)
        btn_row.addStretch()
        btn_row.addWidget(self.manual_btn_stop)
        outer.addLayout(btn_row)

        self.manual_progress = QProgressBar()
        self.manual_progress.setValue(0)
        self.manual_progress.setFixedHeight(8)
        self.manual_progress.setTextVisible(False)
        outer.addWidget(self.manual_progress)

    def _refresh_form_choices(self):
        """Пересобирает выпадающие списки специалистов/организаций из user_data
        (могли измениться в настройках), сохраняя текущий выбор."""
        import user_data
        for key, list_name in (("specialist", "specialists"), ("org", "epb_orgs")):
            combo = self.manual_fields[key]
            current = combo.currentData() or ""
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("— выберите —" if key == "specialist" else "— не выбрано —", "")
            for val in user_data.get_list(list_name):
                combo.addItem(val, val)
            idx = combo.findData(current)
            combo.setCurrentIndex(idx if idx >= 0 else 0)
            combo.blockSignals(False)

    def _load_manual_points(self):
        self._refresh_form_choices()
        if hasattr(self, "manual_stack"):
            self.manual_stack.setCurrentIndex(0)   # всегда возвращаемся к списку
        points = self._scan_points()
        self.manual_point_list.blockSignals(True)
        self.manual_point_list.clear()
        state_colors = {"invalid": STATUS_COLORS["error"], "skipped": STATUS_COLORS["wait"],
                        "done": STATUS_COLORS["done"], "filled": STATUS_COLORS["done"],
                        "empty": STATUS_COLORS["idle"]}
        for pt in points:
            folder = os.path.join(self._get_local_root(), pt["name"])
            _, state = self._manual_point_state(folder, pt)
            label = pt["name"]
            has_repair = not pt["invalid"] and self._point_has_repair(folder)
            вручную = has_repair and self._repair_is_manual(folder)
            if has_repair:
                label += "   • ремонт (вручную)" if вручную else "   • ремонт"
            item = QListWidgetItem(label)
            item.setIcon(dot_icon(state_colors.get(state, STATUS_COLORS["idle"]), 12))
            item.setData(Qt.UserRole, folder)
            item.setData(Qt.UserRole + 1, state)
            if has_repair:
                item.setToolTip(
                    "Пункт оформляется по старому — ремонт вносится вручную"
                    if вручную else
                    "У пункта есть документы ремонта — при заполнении внесите "
                    "данные ремонта")
            if pt["invalid"]:
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
            self.manual_point_list.addItem(item)
        self.manual_point_list.blockSignals(False)

        # Строка итога на карточке шага 3
        filled = sum(1 for i in range(self.manual_point_list.count())
                     if self.manual_point_list.item(i).data(Qt.UserRole + 1) in ("done", "filled"))
        skipped = sum(1 for i in range(self.manual_point_list.count())
                      if self.manual_point_list.item(i).data(Qt.UserRole + 1) == "skipped")
        total = self.manual_point_list.count()
        if total:
            self.step_cards[3].set_summary(
                f"Заполнено {filled} из {total}" + (f" · пропущено {skipped}" if skipped else ""))

        target = None
        for i in range(self.manual_point_list.count()):
            it = self.manual_point_list.item(i)
            if it.flags() & Qt.ItemIsEnabled:
                if target is None:
                    target = i
                if it.data(Qt.UserRole + 1) == "empty":
                    target = i
                    break
        if target is not None:
            self.manual_point_list.setCurrentRow(target)
        else:
            self.manual_current = None
            self._clear_manual_form()

    def _manual_point_state(self, folder, pt):
        if pt["invalid"]:
            return None, "invalid"
        from status_manager import get_status_value
        if pt.get("skipped") or get_status_value(folder, "skipped", False):
            return None, "skipped"
        if folder in self.manual_drafts:
            values = self.manual_drafts[folder]
        else:
            values = get_status_value(folder, "manual3_draft", None)
        if get_status_value(folder, "manual3", False):
            return None, "done"
        if values and self._required_ok(values):
            return None, "filled"
        return None, "empty"

    def _required_ok(self, values: dict) -> bool:
        purpose = (values.get("purpose") or "").strip()
        specialist = (values.get("specialist") or "").strip()
        date = (values.get("date") or "").strip()
        # B33 обязательна и должна быть существующей датой;
        # B54 необязательна, но если введена — тоже должна быть корректной.
        if not (purpose and specialist and date):
            return False
        if not self._date_ok_or_empty(date):
            return False
        if not self._date_ok_or_empty(values.get("epb_date", "")):
            return False
        return True

    def _on_manual_point_selected(self, current, previous):
        if self._manual_loading:
            return
        if self.manual_current is not None:
            self._save_current_draft()
        if current is None:
            self.manual_current = None
            self._clear_manual_form()
            return
        folder = current.data(Qt.UserRole)
        self.manual_current = folder
        self._load_form_for(folder)

    def _point_has_repair(self, folder) -> bool:
        try:
            from repair.detect import find_repair_docs
            return find_repair_docs(folder).has_repair
        except Exception:
            return False

    @staticmethod
    def _repair_is_manual(folder) -> bool:
        """Пункт отмечен «оформляется по старому» (руками, без модуля)."""
        from status_manager import get_status_value
        return bool(get_status_value(folder, "repair_manual", False))

    def _update_repair_banner(self, folder):
        """Баннер ремонта: приглашение в модуль либо отметка «по старому»."""
        if not folder or not self._point_has_repair(folder):
            self.manual_repair_banner.setVisible(False)
            return
        по_старому = self._repair_is_manual(folder)
        if по_старому:
            self.manual_repair_hint.setText(
                "Ремонт этого пункта оформляется по старому — вручную в Excel, "
                "Word и графическом редакторе. Модуль в файлы пункта ничего не "
                "пишет.")
            self.manual_btn_repair.setText("Всё-таки внести через модуль")
            self.manual_btn_repair_old.setText("Открыть папку пункта")
        else:
            self.manual_repair_hint.setText(
                "У этого пункта есть ремонт (акты отбраковки / исполнительная). "
                "Внесите данные ремонта — они запишутся вместе с паспортом. "
                "Либо оформите пункт по старому, вручную.")
            self.manual_btn_repair.setText("Внести данные по ремонту")
            self.manual_btn_repair_old.setText("Оформить по старому")
        self.manual_repair_banner.setVisible(True)

    def _repair_old_way(self):
        """
        Отметить пункт как оформляемый вручную и открыть его папку.

        Ничего не записывает в книгу, отчёт и карту: смысл кнопки ровно в том,
        чтобы модуль к пункту не притрагивался. Отметка живёт в
        `.pnos_status.json` (ключ `repair_manual`) и снимается той же кнопкой.
        """
        folder = self.manual_current
        if not folder:
            return
        from status_manager import set_stage_status

        if self._repair_is_manual(folder):
            os.startfile(folder)                      # уже отмечен — просто папка
            return

        ответ = QMessageBox.question(
            self, "Оформить по старому",
            f"Пункт {os.path.basename(folder)} будет оформляться вручную.\n\n"
            f"Модуль «Ремонт» не тронет ни лист «Импорт», ни лист «Ремонт», "
            f"ни строку «Сведения о ремонте» в отчёте, ни коррозионную карту — "
            f"всё это правите сами, как раньше.\n\n"
            f"Отметка снимается этой же кнопкой. Продолжить?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if ответ != QMessageBox.Yes:
            return

        set_stage_status(folder, "repair_manual", True)
        try:
            from repair.safety import append_journal
            append_journal(folder, "Ремонт оформляется по старому (вручную)", [
                "модуль «Ремонт» в файлы пункта не пишет",
                "вручную: лист «Импорт» (замеры заменённых сечений), "
                "строка «Сведения о ремонте» в таблице 1 отчёта, "
                "коррозионная карта (Первичка)",
            ])
        except Exception as e:
            self._add_activity(3, f"Журнал ремонта не дописан: {e}", "warn")

        self._add_activity(
            3, f"Пункт {os.path.basename(folder)}: ремонт оформляется по "
               f"старому — модуль его не трогает", "info")
        self._update_repair_banner(folder)
        self._refresh_repair_marks()
        os.startfile(folder)

    def _refresh_repair_marks(self):
        """Обновить пометки ремонта в списке пунктов без полной перезагрузки."""
        for i in range(self.manual_point_list.count()):
            it = self.manual_point_list.item(i)
            folder = it.data(Qt.UserRole)
            if not folder or "• ремонт" not in it.text():
                continue
            имя = it.text().split("   •")[0]
            it.setText(имя + ("   • ремонт (вручную)"
                              if self._repair_is_manual(folder)
                              else "   • ремонт"))

    def _open_repair_for_current(self):
        folder = self.manual_current
        if not folder:
            return
        if self._repair_is_manual(folder):
            # Возврат в модуль: отметка снимается, дальше обычный путь.
            from status_manager import set_stage_status
            set_stage_status(folder, "repair_manual", False)
            self._add_activity(
                3, f"Пункт {os.path.basename(folder)}: ремонт снова ведётся "
                   f"через модуль", "info")
            self._update_repair_banner(folder)
            self._refresh_repair_marks()
        self.manual_repair_title.setText(
            f"Ремонт · пункт {os.path.basename(folder)}")
        self.manual_repair_panel.load_point(folder)
        self.manual_stack.setCurrentIndex(1)

    def _close_repair_view(self):
        self.manual_stack.setCurrentIndex(0)

    def _load_form_for(self, folder):
        from status_manager import get_status_value
        self._manual_loading = True
        try:
            self.manual_form_title.setText(f"Пункт {os.path.basename(folder)}")
            self._update_repair_banner(folder)
            if folder in self.manual_drafts:
                values = self.manual_drafts[folder]
            else:
                values = get_status_value(folder, "manual3_draft", None) or {}
            prefix = CONFIG["MANUAL_FORM"].get("PURPOSE_PREFIX", "транспортировки ")
            f = self.manual_fields
            f["purpose"].setText(values.get("purpose") or prefix)
            # Года в черновике ещё нет — подставляем из паспорта (ключ отсутствует).
            # Пустая строка в черновике означает, что поле очистили руками, и
            # подставлять заново нельзя.
            year = values.get("commission_year")
            if year is None:
                year = self._passport_commission_year(folder)
            f["commission"].setText(year or "")
            self._set_combo(f["specialist"], values.get("specialist", ""))
            self._set_date_edit(f["date"], values.get("date", ""))
            self._set_combo(f["org"], values.get("epb_org", ""))
            f["reg_num"].setText(values.get("reg_num", ""))
            self._set_date_edit(f["epb_date"], values.get("epb_date", ""))
            f["sheets"].setText(str(values.get("sheets", "") or ""))
            self._update_epb_done_label()
        finally:
            self._manual_loading = False

    def _clear_manual_form(self):
        self._manual_loading = True
        try:
            self.manual_form_title.setText("Выберите пункт слева")
            self.manual_repair_banner.setVisible(False)
            f = self.manual_fields
            prefix = CONFIG["MANUAL_FORM"].get("PURPOSE_PREFIX", "транспортировки ")
            f["purpose"].setText(prefix)
            f["commission"].setText("")
            f["specialist"].setCurrentIndex(0)
            self._set_date_edit(f["date"], "")
            f["org"].setCurrentIndex(0)
            f["reg_num"].setText("")
            self._set_date_edit(f["epb_date"], "")
            f["sheets"].setText("")
            self._update_epb_done_label()
        finally:
            self._manual_loading = False

    def _passport_commission_year(self, folder: str) -> str:
        """Год ввода в эксплуатацию из паспорта пункта (DocVariable F051).

        Читается zip-ом из .docx — без Word и COM, поэтому переключение пунктов
        в списке не тормозит. Результат кэшируется на сессию: паспорт за время
        работы не меняется.
        """
        cache = getattr(self, "_commission_cache", None)
        if cache is None:
            cache = self._commission_cache = {}
        if folder not in cache:
            try:
                import passport_data
                cache[folder] = passport_data.year_of(
                    passport_data.commission_date(folder))
            except Exception:
                cache[folder] = ""
        return cache[folder]

    def _set_combo(self, combo, value: str):
        idx = combo.findData(value)
        combo.setCurrentIndex(idx if idx >= 0 else 0)

    def _collect_form_values(self) -> dict:
        f = self.manual_fields
        return {
            "purpose": f["purpose"].text().strip(),
            "commission_year": f["commission"].text().strip(),
            "specialist": f["specialist"].currentData() or "",
            "date": self._date_edit_value(f["date"]),
            "epb_org": f["org"].currentData() or "",
            "reg_num": f["reg_num"].text().strip(),
            "epb_date": self._date_edit_value(f["epb_date"]),
            "sheets": f["sheets"].text().strip(),
        }

    def _save_current_draft(self):
        if not self.manual_current:
            return
        values = self._collect_form_values()
        self.manual_drafts[self.manual_current] = values
        try:
            from status_manager import set_stage_status
            set_stage_status(self.manual_current, "manual3_draft", values)
        except Exception:
            pass

    def _update_epb_done_label(self, *args):
        org = self.manual_fields["org"].currentData() or ""
        self.manual_epb_done.setText(
            f"Проводилась ранее (B51): {'да' if org.strip() else 'нет'}")
        combo = self.manual_fields["org"]
        combo.setProperty("attention", not bool(org.strip()))
        combo.style().unpolish(combo)
        combo.style().polish(combo)

    def _skip_current_point(self):
        if not self.manual_current:
            return
        from status_manager import set_stage_status
        folder = self.manual_current
        set_stage_status(folder, "skipped", True)
        set_stage_status(folder, "manual3", False)
        self.manual_drafts.pop(folder, None)
        self.manual_current = None
        self._add_activity(3, f"{os.path.basename(folder)} — пропущен", "warn")
        self._load_manual_points()

    def _open_current_in_excel(self):
        if not self.manual_current:
            return
        from workers.stage2_worker import find_file_exact_ext
        path = find_file_exact_ext(self.manual_current, (".xls", ".xlsx", ".xlsm"))
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось открыть файл:\n{e}")
        else:
            QMessageBox.information(self, "Файл не найден",
                                    "ПНОС_ТТП.xlsm в папке пункта не найден.")

    def _start_manual_stage3(self):
        """Валидация формы всех пунктов и запись данных в Excel."""
        from status_manager import get_status_value

        self._save_current_draft()

        points = self._scan_points()
        entries = []
        not_ready = []
        root = self._get_local_root()

        for pt in points:
            if pt["invalid"]:
                continue
            folder = os.path.join(root, pt["name"])
            if pt.get("skipped") or get_status_value(folder, "skipped", False):
                continue
            values = self.manual_drafts.get(folder) or get_status_value(folder, "manual3_draft", None)
            if not values or not self._required_ok(values):
                not_ready.append(pt["name"])
                continue
            entries.append({"folder": folder, "num": pt["name"].replace("п.", ""), "values": values})

        if not_ready:
            QMessageBox.warning(
                self, "Заполнены не все пункты",
                "Не заполнены обязательные поля (Предназначение, Специалист, Дата) "
                "или даты введены некорректно у пунктов:\n\n" + ", ".join(not_ready) +
                "\n\nПроверьте поля или нажмите «Пропустить пункт».")
            self._add_activity(3, f"Не заполнены пункты: {', '.join(not_ready)}", "warn")
            return

        if not entries:
            QMessageBox.information(
                self, "Нет данных для записи",
                "Нет заполненных пунктов для записи. Заполните хотя бы один пункт "
                "или пропустите все.")
            return

        self.manual_btn_write.setText("Запись…")
        self.manual_btn_write.setEnabled(False)
        self.manual_btn_stop.setEnabled(True)
        self.manual_progress.setValue(0)
        self._begin_step(3)

        from workers.manual3_worker import Manual3Worker
        worker = Manual3Worker(entries, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(3, msg, cat))
        worker.progress.connect(self.manual_progress.setValue)
        worker.progress.connect(self.step_cards[3].set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(3, ["№ Пункта", "Записано", "Результат"], d))
        worker.finished_ok.connect(self._on_manual_written)
        worker.start()
        self._workers["manual3"] = worker

    def _on_manual_written(self, success: bool):
        self.manual_btn_write.setText("Записать всё в Excel")
        self.manual_btn_write.setEnabled(True)
        self.manual_btn_stop.setEnabled(False)
        self._load_manual_points()
        if success:
            self._switch_page(self.PAGE_PROCESS)
        self._on_step_finished(3, success)
