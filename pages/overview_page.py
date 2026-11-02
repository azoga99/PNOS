# -*- coding: utf-8 -*-
"""
Страница «Обзор и отчёт»: матрица статусов пунктов и «Отчёт о прогоне».

Примесь к MainWindow: методы работают с тем же self, что и окно. Общее
для всех страниц (_get_local_root, _add_activity, _switch_page, _workers)
живёт в main_window.py.
"""

import os
import time

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame,
    QMessageBox, QSizePolicy, QStackedWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QScrollArea,
)

from widgets import STATUS_COLORS


class OverviewPageMixin:
    """Страница «Обзор и отчёт»: матрица статусов пунктов и «Отчёт о прогоне»."""

    # ═══════════════════════════════════════════════════════════════
    # Страница «Обзор»
    # ═══════════════════════════════════════════════════════════════

    def _build_chip(self, caption: str, value_obj: str = "chip_value"):
        """Компактный чип-сводка. Возвращает (frame, value_label)."""
        box = QFrame()
        box.setObjectName("stat_chip")
        v = QVBoxLayout(box)
        v.setContentsMargins(16, 7, 16, 7)
        v.setSpacing(0)
        val = QLabel("0")
        val.setObjectName(value_obj)
        val.setAlignment(Qt.AlignCenter)
        cap = QLabel(caption)
        cap.setObjectName("chip_caption")
        cap.setAlignment(Qt.AlignCenter)
        v.addWidget(val)
        v.addWidget(cap)
        return box, val

    def _setup_overview_page(self):
        """Страница «Обзор и отчёт» с вкладками: Обзор | Отчёт о прогоне."""
        outer = QVBoxLayout(self.page_overview)
        outer.setContentsMargins(2, 2, 6, 2)
        outer.setSpacing(12)

        # ── Шапка: заголовок + переключатель вкладок ──
        head = QHBoxLayout()
        head.addWidget(self._page_title("Обзор и отчёт"))
        head.addStretch()
        self.btn_tab_overview = QPushButton("Обзор")
        self.btn_tab_report = QPushButton("Отчёт о прогоне")
        for b in (self.btn_tab_overview, self.btn_tab_report):
            b.setObjectName("segment")
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
        self.btn_tab_overview.setChecked(True)
        self.btn_tab_overview.clicked.connect(lambda: self._set_overview_tab(0))
        self.btn_tab_report.clicked.connect(lambda: self._set_overview_tab(1))
        head.addWidget(self.btn_tab_overview)
        head.addWidget(self.btn_tab_report)
        outer.addLayout(head)

        self.overview_tabs = QStackedWidget()
        outer.addWidget(self.overview_tabs, 1)

        # ═══ Вкладка 0: Обзор ═══
        tab_ov = QWidget()
        ov = QVBoxLayout(tab_ov)
        ov.setContentsMargins(0, 0, 0, 0)
        ov.setSpacing(10)

        chips_row = QHBoxLayout()
        chips_row.setSpacing(8)
        chip_total, self.chip_total = self._build_chip("ПУНКТОВ")
        chip_done, self.chip_done = self._build_chip("ГОТОВЫ", "chip_value_good")
        chip_bad, self.chip_bad = self._build_chip("НЕ КОРРЕКТНЫХ", "chip_value_bad")
        chip_skip, self.chip_skip = self._build_chip("ПРОПУЩЕНО", "chip_value_warn")
        for c in (chip_total, chip_done, chip_bad, chip_skip):
            chips_row.addWidget(c)
        chips_row.addStretch()

        btn_refresh = QPushButton("Обновить")
        btn_refresh.setObjectName("ghost_button")
        btn_refresh.setCursor(Qt.PointingHandCursor)
        btn_refresh.clicked.connect(self._refresh_overview)
        btn_reset = QPushButton("Сбросить статус…")
        btn_reset.setObjectName("ghost_button")
        btn_reset.setCursor(Qt.PointingHandCursor)
        btn_reset.clicked.connect(self._reset_point_dialog)
        chips_row.addWidget(btn_refresh, 0, Qt.AlignBottom)
        chips_row.addWidget(btn_reset, 0, Qt.AlignBottom)
        ov.addLayout(chips_row)

        self.overview_table = QTableWidget()
        headers = ["Пункт"] + [str(i) for i in range(1, 8)]
        self.overview_table.setColumnCount(len(headers))
        self.overview_table.setHorizontalHeaderLabels(headers)
        self.overview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.overview_table.setSelectionMode(QTableWidget.NoSelection)
        self.overview_table.verticalHeader().setVisible(False)
        hh = self.overview_table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.Stretch)
        for c in range(1, len(headers)):
            hh.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        ov.addWidget(self.overview_table, 1)

        self.overview_tabs.addWidget(tab_ov)

        # ═══ Вкладка 1: Отчёт о прогоне ═══
        tab_rep = QWidget()
        rp = QVBoxLayout(tab_rep)
        rp.setContentsMargins(0, 0, 0, 0)
        rp.setSpacing(10)

        rep_head = QHBoxLayout()
        self.lbl_report_summary = QLabel("Отчёт появится после выполнения этапов.")
        self.lbl_report_summary.setObjectName("card_description")
        rep_head.addWidget(self.lbl_report_summary)
        rep_head.addStretch()
        btn_open = QPushButton("📂 Папка отчётов")
        btn_open.setObjectName("ghost_button")
        btn_open.setCursor(Qt.PointingHandCursor)
        btn_open.clicked.connect(self._open_reports_folder)
        rep_head.addWidget(btn_open)
        rp.addLayout(rep_head)

        rep_scroll = QScrollArea()
        rep_scroll.setWidgetResizable(True)
        rep_scroll.setFrameShape(QFrame.NoFrame)
        rep_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        rep_body = QWidget()
        rep_v = QVBoxLayout(rep_body)
        rep_v.setContentsMargins(0, 0, 6, 0)
        rep_v.setSpacing(10)
        self._report_sections = QVBoxLayout()
        self._report_sections.setSpacing(10)
        rep_v.addLayout(self._report_sections)
        rep_v.addStretch()
        rep_scroll.setWidget(rep_body)
        rp.addWidget(rep_scroll, 1)

        self.overview_tabs.addWidget(tab_rep)

    def _set_overview_tab(self, idx: int):
        """Переключает вкладку страницы «Обзор и отчёт»."""
        self.overview_tabs.setCurrentIndex(idx)
        self.btn_tab_overview.setChecked(idx == 0)
        self.btn_tab_report.setChecked(idx == 1)

    def _scan_points(self):
        """Сканирует локальную папку: [{name, invalid, missing, skipped, steps{}}]."""
        import json
        root = self._get_local_root()
        result = []
        if not root or not os.path.isdir(root):
            return result
        for name in sorted(os.listdir(root)):
            folder = os.path.join(root, name)
            if not os.path.isdir(folder):
                continue
            low = name.lower()
            if not (low.startswith("п.") or low.startswith("п ") or low.startswith("пункт")):
                continue
            status_path = os.path.join(folder, ".pnos_status.json")
            data = {}
            if os.path.exists(status_path):
                try:
                    with open(status_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception:
                    data = {}
            steps = {ui: bool(data.get(key, False))
                     for ui, key in self.STEP_STATUS_KEYS.items()}
            result.append({
                "name": name,
                "invalid": bool(data.get("invalid", False)) or ("не корректн" in low),
                "missing": data.get("missing_folders", []),
                "skipped": bool(data.get("skipped", False)),
                "steps": steps,
            })
        return result

    def _refresh_overview(self):
        points = self._scan_points()
        self.overview_table.setRowCount(len(points))

        c_done = QColor(STATUS_COLORS["done"])
        c_wait = QColor(STATUS_COLORS["idle"])
        c_skip = QColor(STATUS_COLORS["wait"])
        c_bad = QColor(STATUS_COLORS["error"])
        bg_bad = QColor(239, 68, 68, 26)     # полупрозрачная подсветка строки
        bg_skip = QColor(245, 158, 11, 20)

        done_all = invalid_n = skipped_n = 0
        for r, pt in enumerate(points):
            invalid, skipped = pt["invalid"], pt["skipped"]

            # Имя пункта с причиной прямо в строке
            name = pt["name"]
            if invalid:
                reason = ", ".join(pt["missing"]) if pt["missing"] else "нет обязательных папок"
                base = name.split(" (")[0]  # без суффикса «(не корректный)»
                name_item = QTableWidgetItem(f"{base} — {reason}")
                name_item.setForeground(c_bad)
                invalid_n += 1
            elif skipped:
                name_item = QTableWidgetItem(f"{name} — пропущен")
                name_item.setForeground(QColor("#b45309"))
                skipped_n += 1
            else:
                name_item = QTableWidgetItem(name)
            self.overview_table.setItem(r, 0, name_item)

            all_done = True
            for ui_step in range(1, 8):
                if invalid:
                    glyph, color = "✕", c_bad
                    all_done = False
                elif skipped and ui_step >= 3:
                    glyph, color = "›", c_skip
                elif pt["steps"].get(ui_step):
                    glyph, color = "✓", c_done
                else:
                    glyph, color = "—", c_wait
                    all_done = False
                cell = QTableWidgetItem(glyph)
                cell.setTextAlignment(Qt.AlignCenter)
                cell.setForeground(color)
                self.overview_table.setItem(r, ui_step, cell)

            # Подсветка всей строки для проблемных/пропущенных
            if invalid or skipped:
                bg = bg_bad if invalid else bg_skip
                for c in range(self.overview_table.columnCount()):
                    it = self.overview_table.item(r, c)
                    if it:
                        it.setBackground(bg)

            if not invalid and not skipped and all_done:
                done_all += 1

        # Чипы-сводка
        self.chip_total.setText(str(len(points)))
        self.chip_done.setText(str(done_all))
        self.chip_bad.setText(str(invalid_n))
        self.chip_skip.setText(str(skipped_n))

    def _reset_point_dialog(self):
        """Диалог: выбрать пункт и что сбросить (весь пункт / конкретный шаг)."""
        from PySide6.QtWidgets import QDialog, QDialogButtonBox
        from widgets import NoWheelComboBox as QComboBox
        points = [pt["name"] for pt in self._scan_points()]
        if not points:
            QMessageBox.information(self, "Нет пунктов", "В локальной папке нет пунктов.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Сброс статуса")
        v = QVBoxLayout(dlg)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)

        v.addWidget(QLabel("Пункт:"))
        cb_point = QComboBox()
        cb_point.addItem("— все пункты —")
        cb_point.addItems(points)
        v.addWidget(cb_point)

        v.addWidget(QLabel("Что сбросить:"))
        cb_scope = QComboBox()
        cb_scope.addItem("Весь пункт (обработать заново с шага 1)", "all")
        for step in range(1, 8):
            cb_scope.addItem(f"Только шаг {step}: {self.STEP_TITLES[step]}",
                             self.STEP_STATUS_KEYS[step])
        v.addWidget(cb_scope)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Сбросить")
        buttons.button(QDialogButtonBox.Cancel).setText("Отмена")
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        v.addWidget(buttons)

        if dlg.exec() != QDialog.Accepted:
            return

        from status_manager import reset_point_status, reset_stage_status
        scope = cb_scope.currentData()
        root = self._get_local_root()

        # Список целей: конкретный пункт или все
        if cb_point.currentIndex() == 0:
            if QMessageBox.question(
                    self, "Сброс всех пунктов",
                    f"Сбросить {'все статусы' if scope == 'all' else 'этот шаг'} "
                    f"у всех пунктов ({len(points)})?\n\nСкачанные файлы не удаляются.",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
                return
            targets = points
        else:
            targets = [cb_point.currentText()]

        ok_n = 0
        for name in targets:
            folder = os.path.join(root, name)
            if scope == "all":
                ok_n += 1 if reset_point_status(folder) else 0
            else:
                ok_n += 1 if reset_stage_status(folder, scope) else 0
        what = "статус полностью" if scope == "all" else f"шаг «{scope}»"
        self._add_activity(0, f"Сброшен {what}: {ok_n} из {len(targets)} пунктов", "done")
        self._refresh_overview()


    # ═══════════════════════════════════════════════════════════════
    # Страница «Итоговый отчёт»
    # ═══════════════════════════════════════════════════════════════

    def _open_reports_folder(self):
        root = self._get_local_root()
        if root and os.path.isdir(root):
            try:
                os.startfile(root)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось открыть папку:\n{e}")
        else:
            QMessageBox.information(self, "Папка не найдена", "Локальная папка отчётов не найдена.")

    def _refresh_report_page(self):
        """Пересобирает секции отчёта из накопленных результатов этапов."""
        # Очистка старых секций
        while self._report_sections.count():
            item = self._report_sections.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        if not self._stage_results:
            self.lbl_report_summary.setText("Отчёт появится после выполнения этапов.")
            return

        # Сводка
        total_errors = 0
        for step, res in self._stage_results.items():
            _, errors, _ = self._summarize_details(res.get("details", []))
            total_errors += errors
        elapsed = self._fmt_duration(time.time() - self._start_time) if self._start_time else "—"
        skipped_pts = sum(1 for pt in self._scan_points() if pt["skipped"])
        invalid_pts = sum(1 for pt in self._scan_points() if pt["invalid"])
        summary = f"Время: {elapsed} · Ошибок: {total_errors}"
        if skipped_pts:
            summary += f" · Пропущено пунктов: {skipped_pts}"
        if invalid_pts:
            summary += f" · Не корректных: {invalid_pts}"
        self.lbl_report_summary.setText(summary)

        # Секции по шагам
        for step in sorted(self._stage_results.keys()):
            res = self._stage_results[step]
            columns, details = res.get("columns", []), res.get("details", [])
            if not details:
                continue

            box = QFrame()
            box.setObjectName("form_card")
            box.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            bv = QVBoxLayout(box)
            bv.setContentsMargins(14, 12, 14, 12)
            bv.setSpacing(8)

            title = QLabel(f"Шаг {step}: {self.STEP_TITLES.get(step, '')}")
            title.setObjectName("form_field_label")
            bv.addWidget(title)

            table = QTableWidget()
            table.setColumnCount(len(columns))
            table.setHorizontalHeaderLabels(columns)
            table.setRowCount(len(details))
            table.setEditTriggers(QTableWidget.NoEditTriggers)
            table.setSelectionMode(QTableWidget.NoSelection)
            table.verticalHeader().setVisible(False)
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.setAlternatingRowColors(True)

            c_ok = QColor(STATUS_COLORS["done"])
            c_bad = QColor(STATUS_COLORS["error"])
            for r, d in enumerate(details):
                for c, col in enumerate(columns):
                    val = d.get(col, "")
                    if isinstance(val, bool):
                        item = QTableWidgetItem("✓" if val else "—")
                        item.setForeground(c_ok if val else QColor(STATUS_COLORS["idle"]))
                    else:
                        text = str(val)
                        item = QTableWidgetItem(text)
                        if col == "Результат":
                            low = text.lower()
                            good = "успешно" in low or "✅" in low or "частично" in low or "пропущен" in low
                            item.setForeground(c_ok if good else c_bad)
                    item.setTextAlignment(Qt.AlignCenter)
                    table.setItem(r, c, item)

            # Высота таблицы под содержимое (без внутренней прокрутки)
            h = table.horizontalHeader().height() + 2
            for r in range(table.rowCount()):
                h += table.rowHeight(r)
            table.setFixedHeight(min(h + 4, 420))
            bv.addWidget(table)
            self._report_sections.addWidget(box)
