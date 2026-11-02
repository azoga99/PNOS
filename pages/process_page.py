# -*- coding: utf-8 -*-
"""
Страница «Процесс»: карточки шагов, запуск этапов, принудительная остановка.

Примесь к MainWindow: методы работают с тем же self, что и окно. Общее
для всех страниц (_get_local_root, _add_activity, _switch_page, _workers)
живёт в main_window.py.
"""

import os
import subprocess
import time

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGridLayout, QLabel, QPushButton, QFrame,
    QLineEdit, QTextEdit, QMessageBox, QApplication, QListWidget,
    QScrollArea,
)

from widgets import StepCard


class ProcessPageMixin:
    """Страница «Процесс»: карточки шагов, запуск этапов, принудительная остановка."""

    # ═══════════════════════════════════════════════════════════════
    # Страница «Процесс» — конвейер шагов
    # ═══════════════════════════════════════════════════════════════

    def _setup_process_page(self):
        outer = QVBoxLayout(self.page_process)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        body = QWidget()
        v = QVBoxLayout(body)
        v.setContentsMargins(2, 2, 8, 2)
        v.setSpacing(10)
        scroll.setWidget(body)
        outer.addWidget(scroll)

        v.addWidget(self._page_title("Процесс"))

        # ── Параметры запуска ──
        params = QFrame()
        params.setObjectName("settings_panel")
        g = QGridLayout(params)
        g.setContentsMargins(14, 12, 14, 12)
        g.setHorizontalSpacing(10)
        g.setVerticalSpacing(8)

        lbl_pts = QLabel("Пункты")
        lbl_pts.setAlignment(Qt.AlignTop)
        g.addWidget(lbl_pts, 0, 0)
        self.text_manual = QTextEdit()
        self.text_manual.setObjectName("manual_input")
        self.text_manual.setPlaceholderText(
            "Номера пунктов через пробел, запятую или с новой строки. Пример: 2758, 2759, 2760")
        self.text_manual.setFixedHeight(56)
        g.addWidget(self.text_manual, 0, 1, 1, 2)

        g.addWidget(QLabel("Локальная папка"), 1, 0)
        self.entry_local = QLineEdit()
        desktop = os.path.join(os.path.expanduser("~"), "Desktop", "ПНОС")
        self.entry_local.setText(desktop)
        g.addWidget(self.entry_local, 1, 1)
        btn_browse = QPushButton("📂")
        btn_browse.setObjectName("browse_btn")
        btn_browse.setFixedWidth(40)
        btn_browse.setCursor(Qt.PointingHandCursor)
        btn_browse.clicked.connect(self._browse_local)
        g.addWidget(btn_browse, 1, 2)
        v.addWidget(params)

        # ── Карточки шагов ──
        self.step_cards = {}
        actions = {
            # Лямбды: сигнал кнопки не должен попасть в параметр points
            1: lambda: self._start_stage1(), 2: lambda: self._start_stage2(),
            3: lambda: self._switch_page(self.PAGE_MANUAL),
            4: lambda: self._start_stage3(), 5: lambda: self._start_stage4(),
            6: lambda: self._start_stage5(), 7: lambda: self._start_stage6(),
        }
        for step in range(1, 8):
            card = StepCard(
                step, self.STEP_TITLES[step], self.STEP_DESCS[step],
                action_text="Заполнить форму" if step == 3 else "Запустить")
            card.action_clicked.connect(actions[step])
            card.stop_clicked.connect(self._force_stop_active_stage)
            self.step_cards[step] = card
            v.addWidget(card)

        # ── Единый журнал ──
        lbl_log = QLabel("Журнал")
        lbl_log.setObjectName("section_label")
        v.addWidget(lbl_log)
        self.activity_list = QListWidget()
        self.activity_list.setObjectName("activity_feed")
        self.activity_list.setMinimumHeight(140)
        v.addWidget(self.activity_list, 1)

    # ═══════════════════════════════════════════════════════════════
    # Управление шагами (карточками)
    # ═══════════════════════════════════════════════════════════════

    def _begin_step(self, step: int):
        """Переводит карточку шага в состояние «выполняется»."""
        card = self.step_cards[step]
        card.set_status("run")
        card.set_progress(0)
        card.set_summary("Выполняется…")
        if step == 4:
            card.set_summary("Выполняется… Не открывайте другие файлы Excel и не перехватывайте фокус мыши")
        self._step_started[step] = time.time()

    def _summarize_details(self, details: list):
        """Считает (готово, ошибки, пропущено) по списку строк-результатов."""
        done = errors = skipped = 0
        for d in details:
            res = str(d.get("Результат", d.get("status", ""))).lower()
            if "пропущен" in res or "⚪" in res:
                skipped += 1
            elif "успешно" in res or "✅" in res or "✔" in res or "частично" in res:
                done += 1
            elif res:
                errors += 1
            else:
                done += 1
        return done, errors, skipped

    def _store_stage_result(self, step: int, columns: list, details: list):
        """Сохраняет результат этапа для страницы «Итоговый отчёт» и карточки."""
        self._stage_results[step] = {"columns": columns, "details": details}
        done, errors, skipped = self._summarize_details(details)
        parts = [f"Готово {done} из {len(details)}"]
        if errors:
            parts.append(f"ошибок {errors}")
        if skipped:
            parts.append(f"пропущено {skipped}")
        started = self._step_started.get(step)
        if started:
            parts.append(self._fmt_duration(time.time() - started))
        self.step_cards[step].set_summary(" · ".join(parts))

    def _on_step_finished(self, step: int, success: bool):
        """Общее завершение шага: карточка, авто-цепочка, обзор."""
        card = self.step_cards[step]
        card.set_status("done" if success else "error")
        if success:
            self._completed_stages.add(step)
            self._add_activity(step, "Этап завершён", "done")
        else:
            self._add_activity(step, "Этап завершён с ошибками", "error")
        self._update_global_progress()

        if success:
            if step >= 7:
                # Конвейер завершён — показываем вкладку «Отчёт о прогоне»
                self._session_timer.stop()
                self._switch_page(self.PAGE_OVERVIEW)
                self._set_overview_tab(1)
                return
            next_step = step + 1
            if self.page_settings.chk_auto.isChecked():
                if next_step == 3:
                    # Ручной шаг: подсветить и открыть форму
                    self.step_cards[3].set_status("wait")
                    self.step_cards[3].set_summary("Заполните форму и запишите данные в Excel")
                    self._switch_page(self.PAGE_MANUAL)
                else:
                    QTimer.singleShot(800, self.step_cards[next_step].btn_action.click)

    # ═══════════════════════════════════════════════════════════════
    # Запуск этапов
    # ═══════════════════════════════════════════════════════════════

    def _start_stage1(self, points: list[str] | None = None):
        """points — номера пунктов; None — берутся из поля ввода номеров."""
        local_root = self._get_local_root()

        if points is None:
            import re
            raw_points = re.findall(r"\d+", self.text_manual.toPlainText())
        else:
            raw_points = [str(p).strip() for p in points]
        manual_points = []
        for p in raw_points:
            if p not in manual_points:
                manual_points.append(p)
        if not manual_points:
            QMessageBox.warning(self, "Ошибка", "Укажите хотя бы один номер пункта.")
            return

        self._start_time = time.time()
        self._session_timer.start()
        self._stage_results = {}
        self._begin_step(1)
        self._add_activity(1, f"Запуск: {len(manual_points)} пунктов", "wait")

        card = self.step_cards[1]
        from workers.stage1_worker import Stage1Worker
        # «Стар. ЭПБ» ищется всегда (её отсутствие не блокирует пункт),
        # скорость сети всегда максимальная.
        worker = Stage1Worker(
            excel_path=None, local_root=local_root, manual_points=manual_points,
            need_epb=True, max_threads=10, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(1, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(self._on_stage1_report)
        worker.finished_ok.connect(lambda ok: self._on_step_finished(1, ok))
        worker.start()
        self._workers["stage1"] = worker

    def _on_stage1_report(self, report: dict):
        """Результат этапа 1 (dict) — приводим к общему виду."""
        details = []
        for d in report.get("details", []):
            folders = d.get("folders", {})
            optional = d.get("optional", {})
            details.append({
                "№ Пункта": str(d.get("point", "")),
                "Паспорт": folders.get("Паспорт", False),
                "Первичка": folders.get("Первичка", False),
                "Стар. ЭПБ": folders.get("Стар. ЭПБ", False),
                "Исполнит.": optional.get("Исполнительная", False),
                "Акт отбр.": optional.get("Акт отбраковки", False),
                "Результат": d.get("status", ""),
            })
        columns = ["№ Пункта", "Паспорт", "Первичка", "Стар. ЭПБ",
                   "Исполнит.", "Акт отбр.", "Результат"]
        self._store_stage_result(1, columns, details)

    def _start_stage2(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(2)
        card = self.step_cards[2]
        from workers.stage2_worker import Stage2Worker
        worker = Stage2Worker(local_root, parent=self, points=points)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(2, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(2, ["№ Пункта", "Акт найден", "Таблица 3 найдена", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(2, ok))
        worker.start()
        self._workers["stage2"] = worker

    def _start_stage3(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(4)
        card = self.step_cards[4]
        macro_path = self.page_settings.edit_macro_path.text().strip() or None
        from workers.stage3_worker import Stage3Worker
        worker = Stage3Worker(local_root, macro_master_path=macro_path, parent=self,
                              points=points)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(4, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(4, ["№ Пункта", "Паспорт", "Excel", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(4, ok))
        worker.start()
        self._workers["stage3"] = worker

    def _start_stage4(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(5)
        card = self.step_cards[5]
        from workers.stage4_worker import Stage4Worker
        worker = Stage4Worker(local_root, parent=self, points=points)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(5, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(5, ["№ Пункта", "Макрос Excel", "Слияние Word", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(5, ok))
        worker.start()
        self._workers["stage4"] = worker

    def _start_stage5(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(6)
        card = self.step_cards[6]
        from workers.stage5_worker import Stage5Worker
        worker = Stage5Worker(local_root, parent=self, points=points)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(6, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(6, ["№ Пункта", "Фото", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(6, ok))
        worker.start()
        self._workers["stage5"] = worker

    def _start_stage6(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(7)
        card = self.step_cards[7]
        from workers.stage6_worker import Stage6Worker
        worker = Stage6Worker(local_root, parent=self, points=points)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(7, msg, cat))
        worker.progress.connect(card.set_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(7, ["№ Пункта", "Акты ГИ", "Результат"], d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(7, ok))
        worker.start()
        self._workers["stage6"] = worker

    # ═══════════════════════════════════════════════════════════════
    # Принудительная остановка
    # ═══════════════════════════════════════════════════════════════

    def _force_stop_active_stage(self):
        if self._is_stopping:
            return
        self._is_stopping = True
        try:
            self._append_log("\n🛑 ПРИНУДИТЕЛЬНАЯ ОСТАНОВКА ПОЛЬЗОВАТЕЛЕМ...")
            self._add_activity(0, "Принудительная остановка", "warn")

            for worker in list(self._workers.values()):
                if worker:
                    try:
                        worker.disconnect()
                        if worker.isRunning():
                            worker.cancel()
                            worker.quit()
                            for _ in range(15):
                                if worker.wait(100):
                                    break
                                QApplication.processEvents()
                    except Exception as e:
                        print(f"Ошибка при отключении воркера: {e}")
            self._workers = {}

            for proc in ["EXCEL.EXE", "WINWORD.EXE"]:
                try:
                    subprocess.Popen(f"taskkill /F /IM {proc}", shell=True,
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception:
                    pass

            self._session_timer.stop()
            for card in self.step_cards.values():
                if card.status() == "run":
                    card.set_status("idle")
                    card.reset_summary()
            # Кнопка записи на форме шага 3
            if hasattr(self, "manual_btn_write"):
                self.manual_btn_write.setEnabled(True)
                self.manual_btn_write.setText("Записать всё в Excel")
            # Кнопки страницы «Правка отчётов»
            if hasattr(self, "fix_btn_start"):
                self.fix_btn_start.setEnabled(True)
                self.fix_btn_stop.setEnabled(False)

            self._add_activity(0, "Система готова к новому запуску", "done")
        except Exception as e:
            msg = f"Ошибка в процессе остановки: {e}"
            print(msg)
            self._log_error(msg)
        finally:
            self._is_stopping = False
