# -*- coding: utf-8 -*-
"""
Верификатор пункта (этап Э2 спецификации).

Отдельная, независимая от применения проверка согласованности артефактов.
Принципиально НЕ знает и не спрашивает, кто делал пункт — программа или
человек по-старому: можно работать как привык и просто нажимать «Проверить».

Проверки (без OCR карты — она в этапе Э3):
    R1  есть документы ремонта, но лист «Ремонт» не заведён / пуст;
    R2  строка «Сведения о ремонте» в отчёте не соответствует реестру;
    R3  листы «Импорт» и «Импорт (ПР)» разошлись (они должны быть зеркальны);
    R4  ячейки с ошибками #REF! на листах книги (мусор от ручных правок);
    R5  задвижки в сводной таблице 3 отчёта (их там быть не должно).

Результат — список находок (уровень ok/warn/error, где, что). Пустой список
ошибок = «пункт согласован».
"""

from __future__ import annotations

from dataclasses import dataclass

from repair.model import Section


@dataclass
class Finding:
    level: str      # "ok" / "warn" / "error"
    code: str       # R1…R5
    where: str
    message: str


def _norm_text(s: str) -> str:
    """Нормализация текста для сравнения строк Word: пробелы и «;»-хвосты."""
    return " ".join((s or "").replace(";", "; ").split()).strip(" .")


# ─── R2: строка «Сведения о ремонте» ──────────────────────────────────

def check_svedenia(records, svedenia_now: str | None) -> list[Finding]:
    from repair.svedenia import build_svedenia

    if svedenia_now is None:
        return [Finding("warn", "R2", "Word",
                        "строка «Сведения о ремонте» в отчёте не найдена")]
    want = build_svedenia(records)
    if _norm_text(want) == _norm_text(svedenia_now):
        return [Finding("ok", "R2", "Word", "строка «Сведения о ремонте» "
                        "соответствует реестру")]
    return [Finding("error", "R2", "Word",
                    f"строка «Сведения о ремонте» не соответствует реестру. "
                    f"В отчёте: «{svedenia_now[:120]}…» · по реестру: "
                    f"«{want[:120]}…»")]


# ─── R3: зеркальность «Импорт» ↔ «Импорт (ПР)» ────────────────────────

def check_mirror(a: list[Section], b: list[Section], limit: int = 10) -> list[Finding]:
    by_row_b = {s.row: s for s in b}
    diffs: list[str] = []
    for s in a:
        o = by_row_b.get(s.row)
        if o is None:
            diffs.append(f"строка {s.row}: есть в «Импорт», нет в «Импорт (ПР)»")
            continue
        if (s.number, s.element, s.thickness_nom, s.diameter,
                tuple(s.measures), s.tmin) != \
           (o.number, o.element, o.thickness_nom, o.diameter,
                tuple(o.measures), o.tmin):
            diffs.append(f"строка {s.row} (сечение {s.number}): значения расходятся")
    extra = set(by_row_b) - {s.row for s in a}
    for r in sorted(extra):
        diffs.append(f"строка {r}: есть в «Импорт (ПР)», нет в «Импорт»")

    if not diffs:
        return [Finding("ok", "R3", "Excel",
                        "листы «Импорт» и «Импорт (ПР)» зеркальны")]
    shown = "; ".join(diffs[:limit])
    more = f" (и ещё {len(diffs) - limit})" if len(diffs) > limit else ""
    return [Finding("error", "R3", "Excel",
                    f"листы «Импорт» и «Импорт (ПР)» разошлись: {shown}{more}")]


# ─── R4: #REF! в книге ────────────────────────────────────────────────

def check_ref_errors(xlsm_path: str) -> list[Finding]:
    import openpyxl

    wb = openpyxl.load_workbook(xlsm_path, read_only=True, data_only=True)
    per_sheet: dict[str, int] = {}
    try:
        for ws in wb.worksheets:
            n = 0
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str) and "#REF!" in v:
                        n += 1
            if n:
                per_sheet[ws.title] = n
    finally:
        wb.close()

    if not per_sheet:
        return [Finding("ok", "R4", "Excel", "ошибок #REF! в книге нет")]
    detail = ", ".join(f"«{s}»: {n}" for s, n in per_sheet.items())
    return [Finding("warn", "R4", "Excel",
                    f"ячейки с #REF! (мусор от ручных правок): {detail}")]


# ─── R5: задвижки в сводной таблице 3 отчёта ──────────────────────────

_T3_HEADER = "Наименование элемента трубопровода"


def check_valves_in_table3(docx_path: str, excluded: list[str]) -> list[Finding]:
    import docx

    doc = docx.Document(docx_path)
    excluded_low = [e.lower() for e in excluded]
    for table in doc.tables:
        try:
            head = table.rows[0].cells[0].text
        except IndexError:
            continue
        if _T3_HEADER.lower() not in head.replace("\n", " ").lower():
            continue
        bad = []
        for row in table.rows[1:]:
            first = row.cells[0].text.strip().lower()
            if any(first.startswith(e) for e in excluded_low):
                bad.append(row.cells[0].text.strip())
        if bad:
            return [Finding("warn", "R5", "Word",
                            f"в сводной таблице 3 присутствуют: "
                            f"{', '.join(sorted(set(bad)))} — по правилам их "
                            f"там быть не должно")]
        return [Finding("ok", "R5", "Word", "задвижек в таблице 3 нет")]
    return [Finding("warn", "R5", "Word", "сводная таблица 3 в отчёте не найдена")]


# ─── Сводная проверка пункта ──────────────────────────────────────────

def verify_point(point_folder: str) -> list[Finding]:
    """Полная проверка пункта. Возвращает находки всех уровней."""
    from repair.config import REPAIR_CONFIG
    from repair.detect import find_repair_docs
    from repair.registry import read_registry
    from repair.imports import read_import_sections, SHEET_IMPORT_PR
    from repair.word_writer import read_svedenia
    from workers.repair_worker import find_point_xlsm, find_point_report

    findings: list[Finding] = []

    docs = find_repair_docs(point_folder)
    xlsm = find_point_xlsm(point_folder)
    report = find_point_report(point_folder)

    if not xlsm:
        findings.append(Finding("error", "R0", "Excel",
                                "в пункте не найден ПНОС_ТТП.xlsm"))
        return findings

    records = read_registry(xlsm)

    # R1: документы есть, реестра нет
    if docs.has_repair and not records:
        findings.append(Finding(
            "warn", "R1", "Excel",
            f"найдены документы ремонта ({len(docs.all_files)} шт.), но лист "
            f"«Ремонт» не заведён — строка в отчёте могла быть внесена вручную"))
    elif docs.has_repair:
        findings.append(Finding("ok", "R1", "Excel",
                                f"лист «Ремонт»: записей {len(records)}"))

    # R2: строка Word
    if report:
        findings.extend(check_svedenia(records, read_svedenia(report)))
    else:
        findings.append(Finding("warn", "R2", "Word",
                                "отчёт ЗЭПБ в пункте не найден"))

    # R3: зеркальность
    try:
        a = read_import_sections(xlsm)
        b = read_import_sections(xlsm, sheet=SHEET_IMPORT_PR)
        findings.extend(check_mirror(a, b))
    except Exception as e:
        findings.append(Finding("error", "R3", "Excel",
                                f"не удалось сравнить листы «Импорт»: {e}"))

    # R4: #REF!
    findings.extend(check_ref_errors(xlsm))

    # R5: задвижки в таблице 3
    if report:
        try:
            findings.extend(check_valves_in_table3(
                report, REPAIR_CONFIG["EXCLUDE_FROM_TABLE3"]))
        except Exception as e:
            findings.append(Finding("warn", "R5", "Word",
                                    f"таблица 3 не проверена: {e}"))

    return findings
