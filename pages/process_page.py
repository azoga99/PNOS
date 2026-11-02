# -*- coding: utf-8 -*-
"""
Страница «Пункты» (рабочее место): запуск шагов, цепочка «До конца»,
ход прогона, принудительная остановка.

Сам виджет — pages/workspace_page.py: он только показывает и отдаёт
намерения сигналами, а здесь они превращаются в запуск воркеров.

Примесь к MainWindow: методы работают с тем же self, что и окно. Общее
для всех страниц (_get_local_root, _add_activity, _switch_page, _workers)
живёт в main_window.py.
"""

import os
import subprocess
import time

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import QVBoxLayout, QMessageBox, QApplication

from pages.workspace_page import WorkspacePage, plural_points
from workspace_model import FORM_STEP, build_rows, next_step, point_number


class ProcessPageMixin:
    """Страница «Пункты»: запуск шагов, цепочка «До конца», остановка."""

    # ═══════════════════════════════════════════════════════════════
    # Страница «Пункты» — рабочее место
    # ═══════════════════════════════════════════════════════════════

    def _setup_process_page(self):
        outer = QVBoxLayout(self.page_process)
        outer.setContentsMargins(0, 0, 0, 0)

        self.workspace = ws = WorkspacePage()
        outer.addWidget(ws)
        ws.add_points_requested.connect(self._ws_add_points)
        ws.run_step_requested.connect(self._ws_run_step)
        ws.run_to_end_requested.connect(self._ws_run_to_end)
        ws.open_form_requested.connect(self._open_form_for)
        ws.stop_requested.connect(self._force_stop_active_stage)
        ws.refresh_requested.connect(self._refresh_workspace)
        ws.reset_requested.connect(self._reset_point_dialog)
        ws.folder_requested.connect(self._browse_local)
        ws.btn_tab_report.clicked.connect(self._refresh_report_page)
        self._setup_report_tab(ws.report_container)

        self._run = None            # (UI-шаг, номера пунктов | None) — что идёт сейчас
        self._run_progress = 0
        self._step_summary = {}     # UI-шаг -> «Готово 3 из 4 · …» для журнала
        self._chain = None          # пункты цепочки «До конца»; None — цепочки нет
        self._chain_paused = []     # пункты, на которых цепочка ждёт формы
        # Таблица перечитывается во время прогона, но не чаще раза в 1,5 с
        self._ws_refresh_timer = QTimer(self)
        self._ws_refresh_timer.setSingleShot(True)
        self._ws_refresh_timer.setInterval(1500)
        self._ws_refresh_timer.timeout.connect(self._refresh_workspace)

    def _refresh_workspace(self):
        """Перечитывает пункты из папки; идущий шаг помечает точкой «идёт»."""
        root = self._get_local_root()
        rows = build_rows(root)
        if self._run is not None:
            step, points = self._run
            for row in rows:
                in_scope = points is None or row["point"] in points
                if in_scope and row["steps"].get(step) not in ("done", "error"):
                    row["steps"][step] = "run"
        self.workspace.set_points(rows)
        self.workspace.set_folder_info(root, len(rows))
        self._update_run_status()

    def _update_run_status(self):
        if self._run is None:
            self.workspace.set_run_status(None, None, 0, "")
            return
        step, points = self._run
        point = points[0] if points and len(points) == 1 else None
        elapsed = self._fmt_duration(time.time() - self._start_time) if self._start_time else ""
        self.workspace.set_run_status(step, point, self._run_progress, elapsed)

    def _on_run_progress(self, value: int):
        self._run_progress = value
        self._update_run_status()
        if not self._ws_refresh_timer.isActive():
            self._ws_refresh_timer.start()

    def _start_session_clock(self):
        self._start_time = time.time()
        self._session_timer.start()

    # ═══════════════════════════════════════════════════════════════
    # Намерения рабочего места
    # ═══════════════════════════════════════════════════════════════

    def _is_running(self) -> bool:
        return self._run is not None or self._chain is not None

    def _ws_add_points(self, numbers: list):
        if self._is_running():
            return
        self._stage_results = {}
        self._start_session_clock()
        self._start_stage1(points=numbers)

    def _ws_run_step(self, step: int, points: list):
        if self._is_running():
            return
        if step == FORM_STEP:
            rows = {r["point"]: r for r in build_rows(self._get_local_root())}
            waiting = [p for p in points
                       if p in rows and rows[p]["steps"].get(FORM_STEP) != "done"]
            self._open_form_for((waiting or points)[0])
            return
        self._start_session_clock()
        self._launch_step(step, points)

    def _launch_step(self, step: int, points: list):
        starters = {1: self._start_stage1, 2: self._start_stage2,
                    4: self._start_stage3, 5: self._start_stage4,
                    6: self._start_stage5, 7: self._start_stage6}
        starters[step](points=points)

    def _open_form_for(self, point: str):
        """Форма шага 3, сразу на нужном пункте."""
        self._switch_page(self.PAGE_MANUAL)
        lst = self.manual_point_list
        for i in range(lst.count()):
            item = lst.item(i)
            folder = item.data(Qt.UserRole) or ""
            if point_number(os.path.basename(os.path.normpath(folder))) == str(point):
                if item.flags() & Qt.ItemIsEnabled:
                    lst.setCurrentRow(i)
                break

    # ═══════════════════════════════════════════════════════════════
    # Цепочка «До конца»
    # ═══════════════════════════════════════════════════════════════

    def _ws_run_to_end(self, points: list, resume: bool = False):
        """
        Шаг за шагом до седьмого для отмеченных пунктов. На ручном шаге 3
        цепочка встаёт и открывает форму; после записи формы продолжается
        сама. Пункт, у которого шаг не выполнился, выбывает из цепочки —
        остальные идут дальше.
        """
        if self._is_running():
            return
        self._chain = [str(p) for p in points]
        self._chain_paused = []
        if not resume:
            self._stage_results = {}
        self._start_session_clock()
        self.workspace.set_busy(True)
        what = "продолжение" if resume else "запуск"
        self._add_activity(0, f"«До конца», {what}: {plural_points(len(self._chain))}", "wait")
        self._chain_next()

    def _chain_rows(self) -> dict:
        return {r["point"]: r for r in build_rows(self._get_local_root())}

    def _chain_drop(self, point: str, step: int, reason: str):
        if point in self._chain:
            self._chain.remove(point)
        self._add_activity(step, f"п.{point} выбыл из «До конца»: {reason}", "warn")

    def _chain_next(self):
        if self._chain is None:     # остановлено
            return
        rows = self._chain_rows()
        pending = {}                # пункт -> его следующий шаг
        for point in list(self._chain):
            row = rows.get(point)
            if row is None:
                self._chain_drop(point, 0, "папка пункта не найдена")
                continue
            nxt = next_step(row["steps"])
            if nxt is None:
                self._chain.remove(point)
            elif row["steps"].get(nxt) == "error":
                state = row["state"] or "Пункт заблокирован"
                self._chain_drop(point, nxt, state[:1].lower() + state[1:])
            else:
                pending[point] = nxt

        # Пункты дальше формы не ждут тех, кто на форме стоит
        movable = {p: s for p, s in pending.items() if s != FORM_STEP}
        if movable:
            step = min(movable.values())
            self._launch_step(step, [p for p, s in movable.items() if s == step])
            return

        waiting = [p for p, s in pending.items() if s == FORM_STEP]
        self._chain = None
        self._session_timer.stop()
        self.workspace.set_busy(False)
        self._refresh_workspace()
        if waiting:
            self._chain_paused = waiting
            names = ", ".join(f"п.{p}" for p in waiting)
            self._add_activity(
                FORM_STEP, f"Ждут формы: {names}. После записи формы «До конца» "
                           "продолжится само", "wait")
            self._open_form_for(waiting[0])
        else:
            self._add_activity(0, "«До конца» завершено", "done")
            self.workspace.show_tab("report")
            self._refresh_report_page()

    def _chain_after_step(self, step: int, participants: list):
        if self._chain is None:
            return
        rows = self._chain_rows()
        for point in participants or []:
            row = rows.get(point)
            if row is None or row["steps"].get(step) != "done":
                self._chain_drop(point, step, f"шаг {step} не выполнен, причина в журнале")
        self._chain_next()

    # ═══════════════════════════════════════════════════════════════
    # Ход шага
    # ═══════════════════════════════════════════════════════════════

    def _begin_step(self, step: int, points: list | None = None):
        """Отмечает начало шага: блок прогона, точки «идёт», блокировка кнопок."""
        self._run = (step, [str(p) for p in points] if points else None)
        self._run_progress = 0
        self._step_started[step] = time.time()
        if not self._session_timer.isActive():
            self._start_session_clock()
        self.workspace.set_busy(True)
        self._refresh_workspace()
        if step == 4:
            self._add_activity(4, "Не открывайте другие файлы Excel и не перехватывайте "
                                  "фокус мыши, пока идут макросы", "warn")

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
        """Сохраняет результат этапа для вкладки «Отчёт о прогоне» и журнала."""
        self._stage_results[step] = {"columns": columns, "details": details}
        done, errors, skipped = self._summarize_details(details)
        parts = [f"готово {done} из {len(details)}"]
        if errors:
            parts.append(f"ошибок {errors}")
        if skipped:
            parts.append(f"пропущено {skipped}")
        started = self._step_started.get(step)
        if started:
            parts.append(self._fmt_duration(time.time() - started))
        self._step_summary[step] = " · ".join(parts)

    def _on_step_finished(self, step: int, success: bool):
        """Общее завершение шага: журнал, таблица, цепочка «До конца»."""
        run, self._run = self._run, None
        summary = self._step_summary.pop(step, "")
        if success:
            self._add_activity(step, "Шаг завершён" + (f": {summary}" if summary else ""), "done")
        else:
            self._add_activity(step, "Шаг завершён с ошибками"
                               + (f": {summary}" if summary else ""), "error")
        self._ws_refresh_timer.stop()
        self._update_nav_badges()

        if self._chain is not None:
            self._refresh_workspace()
            participants = run[1] if run else []
            QTimer.singleShot(800, lambda: self._chain_after_step(step, participants))
            return

        self._session_timer.stop()
        self.workspace.set_busy(False)
        self._refresh_workspace()

        if step == FORM_STEP and success and self._chain_paused:
            points, self._chain_paused = self._chain_paused, []
            self._switch_page(self.PAGE_WORKSPACE)
            QTimer.singleShot(800, lambda: self._ws_run_to_end(points, resume=True))

    # ═══════════════════════════════════════════════════════════════
    # Запуск этапов
    # ═══════════════════════════════════════════════════════════════

    def _start_stage1(self, points: list[str] | None = None):
        """points — номера пунктов (строки из поля «Добавить»)."""
        local_root = self._get_local_root()

        manual_points = []
        for p in points or []:
            p = str(p).strip()
            if p and p not in manual_points:
                manual_points.append(p)
        if not manual_points:
            QMessageBox.warning(self, "Ошибка", "Укажите хотя бы один номер пункта.")
            return

        self._begin_step(1, manual_points)
        self._add_activity(1, f"Запуск: {plural_points(len(manual_points))}", "wait")

        from workers.stage1_worker import Stage1Worker
        # «Стар. ЭПБ» ищется всегда (её отсутствие не блокирует пункт),
        # скорость сети всегда максимальная.
        worker = Stage1Worker(
            excel_path=None, local_root=local_root, manual_points=manual_points,
            need_epb=True, max_threads=10, parent=self)
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(1, msg, cat))
        worker.progress.connect(self._on_run_progress)
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

    def _start_worker(self, ui_step: int, key: str, worker, columns: list):
        """Общая обвязка воркеров этапов 2, 4–7."""
        worker.log.connect(self._append_log)
        worker.info.connect(lambda msg, cat: self._add_activity(ui_step, msg, cat))
        worker.progress.connect(self._on_run_progress)
        worker.report_ready.connect(
            lambda d: self._store_stage_result(ui_step, columns, d))
        worker.finished_ok.connect(lambda ok: self._on_step_finished(ui_step, ok))
        worker.start()
        self._workers[key] = worker

    def _start_stage2(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(2, points)
        from workers.stage2_worker import Stage2Worker
        worker = Stage2Worker(local_root, parent=self, points=points)
        self._start_worker(2, "stage2", worker,
                           ["№ Пункта", "Акт найден", "Таблица 3 найдена", "Результат"])

    def _start_stage3(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(4, points)
        macro_path = self.page_settings.edit_macro_path.text().strip() or None
        from workers.stage3_worker import Stage3Worker
        worker = Stage3Worker(local_root, macro_master_path=macro_path, parent=self,
                              points=points)
        self._start_worker(4, "stage3", worker, ["№ Пункта", "Паспорт", "Excel", "Результат"])

    def _start_stage4(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(5, points)
        from workers.stage4_worker import Stage4Worker
        worker = Stage4Worker(local_root, parent=self, points=points)
        self._start_worker(5, "stage4", worker,
                           ["№ Пункта", "Макрос Excel", "Слияние Word", "Результат"])

    def _start_stage5(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(6, points)
        from workers.stage5_worker import Stage5Worker
        worker = Stage5Worker(local_root, parent=self, points=points)
        self._start_worker(6, "stage5", worker, ["№ Пункта", "Фото", "Результат"])

    def _start_stage6(self, points: list[str] | None = None):
        local_root = self._get_local_root()
        self._begin_step(7, points)
        from workers.stage6_worker import Stage6Worker
        worker = Stage6Worker(local_root, parent=self, points=points)
        self._start_worker(7, "stage6", worker, ["№ Пункта", "Акты ГИ", "Результат"])

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
            # Цепочка не должна подхватить «завершение» остановленного шага
            self._chain = None
            self._chain_paused = []

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
            self._ws_refresh_timer.stop()
            self._run = None
            self.workspace.set_busy(False)
            self._refresh_workspace()
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
