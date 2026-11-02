# -*- coding: utf-8 -*-
"""
Вкладка «Отчёт о прогоне» рабочего места и диалог сброса статуса пунктов.

Примесь к MainWindow: методы работают с тем же self, что и окно. Общее
для всех страниц (_get_local_root, _scan_points, _add_activity) живёт в
main_window.py.
"""

import os
import time

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame,
    QMessageBox, QSizePolicy, QTableWidget, QTableWidgetItem, QHeaderView,
    QScrollArea,
)

from widgets import STATUS_COLORS
from workspace_model import point_number


class RunReportMixin:
    """Вкладка «Отчёт о прогоне» и сброс статуса пунктов."""

    # ═══════════════════════════════════════════════════════════════
    # Вкладка «Отчёт о прогоне»
    # ═══════════════════════════════════════════════════════════════

    def _setup_report_tab(self, container: QWidget):
        rp = container.layout()

        rep_head = QHBoxLayout()
        self.lbl_report_summary = QLabel("Отчёт появится после выполнения шагов.")
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
        """Пересобирает секции отчёта из накопленных результатов шагов."""
        while self._report_sections.count():
            item = self._report_sections.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        if not self._stage_results:
            self.lbl_report_summary.setText("Отчёт появится после выполнения шагов.")
            return

        # Сводка
        total_errors = 0
        for step, res in self._stage_results.items():
            _, errors, _ = self._summarize_details(res.get("details", []))
            total_errors += errors
        elapsed = self._fmt_duration(time.time() - self._start_time) if self._start_time else "—"
        points = self._scan_points()
        skipped_pts = sum(1 for pt in points if pt["skipped"])
        invalid_pts = sum(1 for pt in points if pt["invalid"])
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

    # ═══════════════════════════════════════════════════════════════
    # Сброс статуса
    # ═══════════════════════════════════════════════════════════════

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
        # Отмеченный в таблице пункт — сразу выбран
        selected = self.workspace.selected_points()
        if len(selected) == 1:
            for i, name in enumerate(points, start=1):
                if point_number(name) == selected[0]:
                    cb_point.setCurrentIndex(i)
                    break
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
        self._refresh_workspace()
        self._update_nav_badges()
