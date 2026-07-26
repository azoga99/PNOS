# -*- coding: utf-8 -*-
"""
Воркеры страницы «Ремонт» (модуль repair, ветка repair-module).

RepairScanWorker  — фоновое сканирование пункта: детекция документов ремонта,
                    привязка по именам файлов, чтение реестра (лист «Ремонт»)
                    и сечений листа «Импорт» через openpyxl. Без COM.
RepairApplyWorker — фоновое применение: бэкап → COM-запись в книгу
                    («Импорт», «Импорт (ПР)», «Ремонт») → строка «Сведения
                    о ремонте» в отчёт Word → журнал; при ошибке — автооткат
                    (см. repair.orchestrator).

Тяжёлые импорты (openpyxl, docx, win32com) — только внутри run(), чтобы не
ломать быстрый старт приложения.
"""

from __future__ import annotations

import os

from PySide6.QtCore import QThread, Signal


def find_point_report(point_folder: str) -> str | None:
    """
    Отчёт ЗЭПБ в корне пункта: .docx, не шаблон ПНОС_ТТП, предпочтительно
    без «ПЕЧАТЬ» (как в find_report_file этапа 5, без импорта воркера).
    """
    cands = []
    for f in os.listdir(point_folder):
        p = os.path.join(point_folder, f)
        low = f.lower()
        if (os.path.isfile(p) and low.endswith(".docx")
                and not f.startswith("~$") and "пнос_ттп" not in low):
            cands.append(p)
    if not cands:
        return None
    non_print = [p for p in cands if "печать" not in os.path.basename(p).lower()]
    return sorted(non_print or cands)[0]


def find_point_xlsm(point_folder: str) -> str | None:
    for f in sorted(os.listdir(point_folder)):
        p = os.path.join(point_folder, f)
        if os.path.isfile(p) and f.lower().endswith((".xlsm", ".xlsx", ".xls")) \
                and not f.startswith("~$"):
            return p
    return None


class RepairScanWorker(QThread):
    """Сканирование пункта для страницы «Ремонт» (только чтение)."""

    log = Signal(str)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    # результат: dict с ключами docs, verdicts, records, sections, unit, line,
    # xlsm, report, svedenia_now
    result = Signal(dict)

    def __init__(self, point_folder: str, parent=None):
        super().__init__(parent)
        self.point_folder = point_folder

    def run(self):
        try:
            from repair.detect import (find_repair_docs, parse_doc_filename,
                                       match_point, build_records_from_docs)
            from repair.registry import read_registry
            from repair.imports import read_import_sections
            from repair.word_writer import read_svedenia

            folder = self.point_folder
            out: dict = {"folder": folder}

            docs = find_repair_docs(folder)
            out["docs"] = docs
            if not docs.has_repair:
                self.info.emit("Документов ремонта не найдено — состояние С0", "info")
                self.result.emit(out)
                self.finished_ok.emit(True)
                return

            xlsm = find_point_xlsm(folder)
            report = find_point_report(folder)
            out["xlsm"], out["report"] = xlsm, report

            # Карточка пункта для светофора: установка (паспорт!B12),
            # № трубопровода (паспорт!B17). Пусто — вердикт unknown.
            unit = line = ""
            if xlsm:
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(xlsm, read_only=True, data_only=True)
                    ws = wb["паспорт"]
                    unit = str(ws["B12"].value or "").strip()
                    line = str(ws["B17"].value or "").strip()
                    wb.close()
                except Exception as e:
                    self.log.emit(f"паспорт не прочитан: {e}")
            out["unit"], out["line"] = unit, line

            verdicts = []
            for f in docs.all_files:
                p = parse_doc_filename(f)
                verdicts.append((f, p, match_point(p, unit, line)))
            out["verdicts"] = verdicts

            # Реестр: из листа «Ремонт», если он уже есть; иначе черновик
            # из имён документов.
            records = read_registry(xlsm) if xlsm else []
            out["registry_existed"] = bool(records)
            if not records:
                records = build_records_from_docs(docs)
            out["records"] = records

            out["sections"] = read_import_sections(xlsm) if xlsm else []
            out["svedenia_now"] = read_svedenia(report) if report else None

            n_ok = sum(1 for _, _, v in verdicts if v == "ok")
            self.info.emit(
                f"Документов: {len(verdicts)} (совпало {n_ok}) · "
                f"сечений: {len(out['sections'])}", "done")
            self.result.emit(out)
            self.finished_ok.emit(True)

        except Exception as e:
            import traceback
            self.log.emit(traceback.format_exc())
            self.info.emit(f"Ошибка сканирования: {str(e)[:80]}", "error")
            self.finished_ok.emit(False)


class RepairVerifyWorker(QThread):
    """Верификация пункта (этап Э2): только чтение, без COM."""

    log = Signal(str)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    result = Signal(list)            # list[Finding]

    def __init__(self, point_folder: str, parent=None):
        super().__init__(parent)
        self.point_folder = point_folder

    def run(self):
        try:
            from repair.verifier import verify_point
            findings = verify_point(self.point_folder)
            self.result.emit(findings)
            n_err = sum(1 for f in findings if f.level == "error")
            n_warn = sum(1 for f in findings if f.level == "warn")
            if n_err:
                self.info.emit(f"Верификация: ошибок {n_err}, "
                               f"предупреждений {n_warn}", "error")
            elif n_warn:
                self.info.emit(f"Верификация: предупреждений {n_warn}", "warn")
            else:
                self.info.emit("Верификация: пункт согласован", "done")
            self.finished_ok.emit(True)
        except Exception as e:
            import traceback
            self.log.emit(traceback.format_exc())
            self.info.emit(f"Ошибка верификации: {str(e)[:80]}", "error")
            self.finished_ok.emit(False)


class RepairOcrWorker(QThread):
    """
    Распознавание документа (Windows OCR) для предзаполнения реестра.

    Для комплекта ИД сначала строится карта разделов (`repair.id_map`), и
    читаются только нужные страницы: копия акта и удостоверение о качестве
    ремонта. Прежний способ — «первые шесть страниц» — на комплекте в 131
    страницу почти всегда промахивался мимо обоих документов.

    Для отдельного акта (2–4 страницы) читается весь файл: первый лист даёт
    номер и дату акта, последний — таблицу подлежащих замене элементов.
    """

    log = Signal(str)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    result = Signal(object)          # DocInfo
    map_ready = Signal(object)       # карта разделов ИД (dict) или None
    # Фрагмент документа, который сотруднику нужно увидеть глазами:
    # {"page", "полоса", "поворот", "зачем"}. Рукописные даты OCR не читает.
    zone_hint = Signal(object)
    # Полоса этапов: список названий один раз в начале, затем номер текущего.
    # Обычный процент здесь врал бы — сколько страниц придётся прочитать,
    # заранее неизвестно, комплект читается адресно.
    phases = Signal(list)
    phase = Signal(int, str)

    MAX_PAGES = 6                    # запасной путь, когда карта не построена
    SHORT_DOC = 6                    # столько страниц и меньше — читаем целиком

    # Этапы распознавания — по ним заполняется полоса.
    PHASES_ACT = ("Листы акта", "Таблица замены")
    PHASES_ID = ("Реестр ИД", "Карта разделов", "Копия акта",
                 "Таблица замены", "Блок подписей")

    # Зоны интереса на листе (доли высоты).
    ZONE_REPLACEMENTS = (0.70, 1.0)   # таблица «подлежащих замене»
    ZONE_SIGNATURES = (0.62, 1.0)     # блок подписей реестра ИД
    ZONE_GRIF = (0.0, 0.28)           # гриф УТВЕРЖДАЮ на листе акта

    def __init__(self, pdf_path: str, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self._date_done = ""
        self._reg_act_date = ""      # дата акта из печатной шапки реестра ИД
        self._rotations: dict[int, int] = {}
        self._phases: tuple = ()
        # Кроп для сотрудника показываем в самом конце: к этому времени реестр
        # уже дозаполнен распознанным, и курсор можно ставить прямо в ячейку.
        self._pending_hint = None

    def _set_phase(self, name: str):
        """Отметить текущий этап на полосе (по названию из набора этапов)."""
        if name in self._phases:
            self.phase.emit(self._phases.index(name), name)

    # ── чтение страниц ───────────────────────────────────────────────

    def _read(self, page: int, total: int, what: str = "") -> list[str]:
        from repair.ocr import ocr_page, plain_text

        подпись = f" ({what})" if what else ""
        self.info.emit(f"Распознавание: стр. {page}/{total}{подпись}...", "wait")
        # Ориентацию подбирает ocr_page: в комплектах ИД повёрнута каждая
        # восьмая страница, а при 90/270° OCR молча возвращает пустоту.
        res = ocr_page(self.pdf_path, page, try_all=True)
        self._rotations[page] = res["rotation"]
        if res["rotation"]:
            self.log.emit(f"стр. {page}: страница повёрнута на "
                          f"{res['rotation']}°, развёрнута для чтения")
        return plain_text(res).splitlines()

    def _rotation_of(self, page: int) -> int:
        """Поворот страницы, если она уже распознавалась."""
        if page in self._rotations:
            return self._rotations[page]
        from repair.ocr import ocr_page
        rot = ocr_page(self.pdf_path, page, try_all=True)["rotation"]
        self._rotations[page] = rot
        return rot

    def run(self):
        try:
            from repair.ocr import available
            from repair.pdf_preview import page_count

            if not available():
                self.info.emit("Windows OCR недоступен (нет русского "
                               "языкового пакета)", "error")
                self.finished_ok.emit(False)
                return

            total = page_count(self.pdf_path)
            self._date_done = self._reg_act_date = ""
            self._pending_hint = None
            короткий = total <= self.SHORT_DOC
            self._phases = self.PHASES_ACT if короткий else self.PHASES_ID
            self.phases.emit(list(self._phases))
            lines = (self._read_short(total) if короткий
                     else self._read_id_set(total))

            from repair.doc_extract import extract_doc_info
            from repair.dates import act_date
            info = extract_doc_info(lines)
            # Правило Д2: полная дата грифа УТВЕРЖДАЮ → дата из имени файла
            # акта → печатная шапка реестра ИД (день грифа вписан от руки и
            # часто не читается, а в имени комплекта ИД даты обычно нет).
            info.act_date = act_date(info.act_date, self.pdf_path,
                                     self._reg_act_date)
            if self._date_done:
                info.date_done = self._date_done
            self.result.emit(info)
            self.phase.emit(len(self._phases), "")
            # Кроп — последним: реестр уже заполнен, и страница может сразу
            # поставить курсор в ячейку даты.
            if self._pending_hint:
                self.zone_hint.emit(self._pending_hint)

            got = []
            if info.act_numbers:
                got.append("акты " + ", ".join(info.act_numbers))
            if info.act_date:
                got.append(f"дата {info.act_date}")
            if info.unit:
                got.append(f"уст. {info.unit}")
            if info.material:
                got.append(info.material)
            self.info.emit("Распознано: " + ("; ".join(got) if got else
                           "полезных полей не найдено"), "done")
            self.finished_ok.emit(True)

        except Exception as e:
            import traceback
            self.log.emit(traceback.format_exc())
            self.info.emit(f"Ошибка распознавания: {str(e)[:80]}", "error")
            self.finished_ok.emit(False)

    # ── короткий документ: отдельный акт ─────────────────────────────

    def _read_short(self, total: int) -> list[str]:
        self.map_ready.emit(None)
        self._set_phase("Листы акта")
        lines: list[str] = []
        for page in range(1, total + 1):
            lines += self._read(page, total)
        # Таблица «подлежащих замене» — на последнем листе. Мелкие ячейки
        # (типоразмер) OCR читает нестабильно, поэтому зона берётся в
        # несколько проходов с разными масштабами.
        self._set_phase("Таблица замены")
        lines += self._read_replacement_zone(total)
        return lines

    def _read_zone(self, page: int, полоса: tuple,
                   rotation: int = 0) -> list[str]:
        """
        Зона листа, прочитанная в несколько масштабов.

        Мелкие ячейки (типоразмер) Windows OCR читает нестабильно: «108х6-89х6»
        верно при 2.2/3.0/3.5 и с ошибкой при 2.5. Несколько проходов дают
        устойчивый результат.
        """
        from repair.ocr import ocr_zone_multi, rows_from_lines

        multi = ocr_zone_multi(self.pdf_path, page, полоса[0], полоса[1],
                               rotation=rotation)
        out: list[str] = []
        for v in multi["варианты"]:
            # И ячейки по отдельности, и склеенные строки таблицы: в акте
            # «Переход» и «108х6-89х6» распознаются разными ячейками, и
            # запись о замене видна только после склейки.
            out += [ln["text"] for ln in v["lines"]]
            out += rows_from_lines(v["lines"])
        return out

    def _read_replacement_zone(self, page: int, rotation: int = 0) -> list[str]:
        return self._read_zone(page, self.ZONE_REPLACEMENTS, rotation)

    def _read_signatures(self, page: int, rotation: int = 0) -> None:
        """
        Блок подписей реестра ИД: правило Д1 плюс подсказка сотруднику.

        Даты там почти всегда рукописные или штампованные, и OCR их не читает.
        Поэтому: если дата нашлась — подставляем, если нет — просим показать
        сотруднику именно этот фрагмент, чтобы он вписал её одним взглядом.
        """
        from repair.dates import latest_date

        строки = self._read_zone(page, self.ZONE_SIGNATURES, rotation)
        дата = latest_date(строки)
        if дата:
            self._date_done = дата
            self.info.emit(f"Дата выполнения работ: {дата} "
                           f"(самая поздняя в блоке подписей реестра)", "done")
        else:
            self.info.emit("Дата выполнения работ в блоке подписей рукописная — "
                           "OCR её не читает, впишите по фрагменту справа", "warn")
        self._pending_hint = {
            "page": page, "полоса": list(self.ZONE_SIGNATURES),
            "поворот": rotation, "зачем": "дата выполнения работ",
            "найдено": дата,
        }

    # ── комплект ИД: адресное чтение по карте ────────────────────────

    def _read_id_set(self, total: int) -> list[str]:
        from repair.id_map import (build_page_map, quality_cert_section,
                                   section_pages, SEC_ACT)

        self.info.emit("Разбор реестра исполнительной документации...", "wait")
        # Этапы «Реестр ИД» и «Карта разделов» отмечает сам build_page_map —
        # он единственный знает, когда реестр найден.
        page_map = build_page_map(self.pdf_path, log=self.log.emit,
                                  on_phase=self._set_phase)
        self.map_ready.emit(page_map)

        if not page_map or not page_map.get("разделы"):
            self.info.emit("Карта разделов не построена — читаю первые "
                           f"{self.MAX_PAGES} страниц", "warn")
            self._set_phase("Копия акта")
            lines = []
            for page in range(1, min(total, self.MAX_PAGES) + 1):
                lines += self._read(page, total)
            return lines

        self.info.emit(
            f"Реестр разобран: разделов {len(page_map['разделы'])}, "
            f"смещение {page_map['смещение']} "
            f"({page_map.get('смещение_источник', '?')})", "done")

        lines: list[str] = []
        # Шапка реестра даёт номер и дату акта печатным текстом — у части
        # подрядчиков («ТП Инжиниринг»: «№222-1-309 от 07.07.2023г.») это
        # единственный автоматический источник даты акта. Строки собираются в
        # том же виде, в каком они встречаются в документах, чтобы работали
        # обычные правила разбора.
        шапка = page_map.get("шапка") or {}
        акты = шапка.get("акты") or []
        дата = шапка.get("дата_акта")
        if дата:
            # Номер акта в шапке OCR иногда искажает («222-1-ЗО9» кириллицей),
            # и связать дату с номером по тексту не получается. Дату запоминаем
            # отдельно — она уйдёт в правило Д2 третьим источником.
            self._reg_act_date = дата
        for номер in акты:
            lines.append(f"Акт ОТН №{номер}"
                         + (f" от {дата}г." if дата else ""))
        if шапка.get("линия"):
            lines.append(f"Линия №{шапка['линия']}")
        if шапка.get("установка"):
            lines.append(f"уст.{шапка['установка']}")

        цели: list[tuple[int, str]] = []
        for page in section_pages(page_map, SEC_ACT, limit=2):
            цели.append((page, "копия акта"))

        # Лист удостоверения и лист с таблицей замены — не всегда один и тот
        # же: у «ТП Инжиниринг» таблица вынесена в приложение.
        удост = page_map.get("удостоверение_страница")
        if удост:
            цели.append((удост, "удостоверение о качестве"))
        замена = page_map.get("замена_страница")
        if замена and замена != удост:
            цели.append((замена, "таблица замены"))
        if not удост and not замена:
            self.log.emit("Ни удостоверение о качестве, ни таблица замены не "
                          "найдены — данные о замене придётся внести вручную")

        if not цели:
            # Названия разделов в реестре могли слипнуться при распознавании.
            # Но страница копии акта известна: по ней вычислялось смещение.
            акт = (page_map.get("проверка_смещения") or {}).get("страница_акта")
            if акт:
                self.log.emit("Разделы в реестре не опознаны — читаю копию "
                              f"акта, найденную по заголовку (стр. {акт})")
                цели = [(акт, "копия акта"), (акт + 1, "копия акта")]
            else:
                self.info.emit("Ни один раздел не опознан — читаю первые "
                               f"{self.MAX_PAGES} страниц", "warn")
                цели = [(p, "") for p in
                        range(1, min(total, self.MAX_PAGES) + 1)]
            цели = [(p, w) for p, w in цели if 1 <= p <= total]

        self._set_phase("Копия акта")
        for page, what in цели:
            lines += self._read(page, total, what)

        лист_замены = замена or удост
        self._set_phase("Таблица замены")
        if лист_замены:
            # Таблица «Ремонту подвергались следующие участки» печатная —
            # лучший источник типоразмеров и марки стали. Лист бывает
            # повёрнутым, поэтому зону читаем с учётом найденного поворота.
            поворот = self._rotation_of(лист_замены)
            lines += self._read_replacement_zone(лист_замены, поворот)

        # Блок подписей реестра: дата выполнения работ (правило Д1).
        рег = page_map.get("страница_реестра")
        self._set_phase("Блок подписей")
        if рег:
            self._read_signatures(рег, self._rotation_of(рег))
        return lines


class MapLabelsWorker(QThread):
    """
    Поиск подписей сечений на коррозионной карте (фоном).

    Распознавание карты целиком занимает несколько секунд, поэтому выполняется
    в потоке: интерфейс не должен подвисать после загрузки карты.
    """

    log = Signal(str)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    result = Signal(object)          # list[MapLabel]

    def __init__(self, arr, parent=None):
        super().__init__(parent)
        self._arr = arr

    def run(self):
        try:
            from repair.map_labels import find_labels

            подписи = find_labels(self._arr, log=self.log.emit)
            self.result.emit(подписи)
            if подписи:
                номера = sorted({l.number for l in подписи})
                self.info.emit(
                    f"На карте распознано сечений: {len(номера)} "
                    f"(от {номера[0]} до {номера[-1]})", "done")
            else:
                self.info.emit("Подписи сечений на карте распознать не удалось "
                               "— правьте кликом по подписи", "warn")
            self.finished_ok.emit(True)
        except Exception as e:
            import traceback
            self.log.emit(traceback.format_exc())
            self.info.emit(f"Ошибка разбора карты: {str(e)[:80]}", "error")
            self.finished_ok.emit(False)


class RepairApplyWorker(QThread):
    """Применение согласованных изменений (COM Excel + python-docx Word)."""

    log = Signal(str)
    info = Signal(str, str)
    finished_ok = Signal(bool)
    result = Signal(object)          # ApplyResult

    def __init__(self, point_folder: str, xlsm_path: str,
                 report_path: str | None, ops: list, records: list,
                 svedenia_text: str | None, parent=None):
        super().__init__(parent)
        self.point_folder = point_folder
        self.xlsm_path = xlsm_path
        self.report_path = report_path
        self.ops = ops
        self.records = records
        self.svedenia_text = svedenia_text

    def run(self):
        try:
            from repair.orchestrator import apply_repair
            from repair.com_writer import apply_to_workbook
            from repair.word_writer import write_svedenia

            report = self.report_path
            sved = self.svedenia_text

            def writer(xlsm, ops, records, log=None):
                counts = apply_to_workbook(xlsm, ops, records, log=log)
                if report and sved is not None:
                    if write_svedenia(report, sved):
                        counts["word"] = 1
                        if log:
                            log("Отчёт Word: строка «Сведения о ремонте» записана")
                    else:
                        raise RuntimeError(
                            "в отчёте не найдена строка «Сведения о ремонте»")
                return counts

            extra = [report] if report else []
            res = apply_repair(self.point_folder, self.xlsm_path, self.ops,
                               self.records, writer, extra_files=extra,
                               log=self.log.emit)
            self.result.emit(res)
            if res.ok:
                self.info.emit("Применено. Бэкап и журнал созданы.", "done")
            else:
                self.info.emit(f"Ошибка, выполнен откат: {res.error[:80]}", "error")
            self.finished_ok.emit(res.ok)

        except Exception as e:
            import traceback
            self.log.emit(traceback.format_exc())
            self.info.emit(f"Ошибка применения: {str(e)[:80]}", "error")
            self.finished_ok.emit(False)
