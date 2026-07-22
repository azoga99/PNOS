# -*- coding: utf-8 -*-
"""
Обнаружение ремонта у пункта и привязка документов.

Детекция: в папке пункта есть непустая подпапка «Акт отбраковки» и/или
«Исполнительная» → у пункта есть ремонт, модуль включается. Иначе — состояние
С0, новый код пункта не касается.

Привязка (первый, самый дешёвый уровень — имя файла, читается всегда, даже
когда скан ужасный):
    «36-30, №6, 222-1-84.pdf»                     → уст. 36-30, линия 6, акт 222-1-84
    «222-1-121 №122 39-40 24.03.2026.pdf»          → акт 222-1-121, линия 122,
                                                     уст. 39-40, дата 24.03.2026
    «ИД по Акту ОТН №222-1-143,121. Линия №122…»   → акты 222-1-143 и 222-1-121
                                                     (сокращение через запятую)

Итог привязки — светофор: ok (совпало с пунктом) / mismatch (противоречит —
похоже, документ другого пункта) / unknown (в имени данных нет — решает
сотрудник по превью). Ничего не блокируется: последнее слово за человеком.
"""

from __future__ import annotations

import difflib
import os
import re
from dataclasses import dataclass, field

from repair.config import REPAIR_CONFIG

# Номер акта ревизии и отбраковки: «222-1-84», «222-5-04-028» и т.п.
_ACT_RE = re.compile(r"\b(\d{3}-\d(?:-\d+)+)\b")
# Сокращение «222-1-143,121» → второй номер наследует префикс.
_ACT_TAIL_RE = re.compile(r",\s*(\d{2,4})\b")
# Установка: «36-30», «39-40» (2 цифры - 2 цифры, не часть номера акта).
_UNIT_RE = re.compile(r"\b(\d{2}-\d{2})\b")
# Линия/трубопровод: «№6», «№ 122»; приоритетно — после слова «Линия».
_LINE_RE = re.compile(r"№\s*(\d+)")
_LINE_WORD_RE = re.compile(r"[Лл]ини\w*\s*№\s*(\d+)")
_DATE_RE = re.compile(r"\b(\d{2}\.\d{2}\.\d{4})\b")

MATCH_OK = "ok"
MATCH_MISMATCH = "mismatch"
MATCH_UNKNOWN = "unknown"


@dataclass
class ParsedName:
    """Что удалось извлечь из имени файла документа."""
    acts: list[str] = field(default_factory=list)
    unit: str = ""      # установка, напр. «36-30»
    line: str = ""      # № трубопровода, напр. «6»
    date: str = ""      # ДД.ММ.ГГГГ (у актов часто дата акта в имени)
    is_id: bool = False  # похоже на комплект ИД (а не отдельный акт)


@dataclass
class RepairDocs:
    """Результат обнаружения ремонта в папке пункта."""
    has_repair: bool = False
    act_files: list[str] = field(default_factory=list)   # PDF из «Акт отбраковки»
    id_files: list[str] = field(default_factory=list)    # PDF из «Исполнительная»

    @property
    def all_files(self) -> list[str]:
        return self.act_files + self.id_files


def parse_doc_filename(filename: str) -> ParsedName:
    """Разобрать имя файла документа ремонта (без пути)."""
    name = os.path.basename(filename)
    stem = os.path.splitext(name)[0]

    p = ParsedName()
    p.is_id = bool(re.search(r"\bИД\b|исполнительн", stem, re.IGNORECASE))

    # Акты: полные номера + сокращения через запятую («…-143,121»).
    rest = stem
    for m in _ACT_RE.finditer(stem):
        full = m.group(1)
        p.acts.append(full)
        # хвосты «,121» сразу после номера
        tail_from = m.end()
        tail_str = stem[tail_from:]
        prefix = full.rsplit("-", 1)[0]  # «222-1-143» → «222-1»
        for tm in _ACT_TAIL_RE.finditer(tail_str):
            # хвост должен идти подряд (через запятые), не через слова
            if tail_str[:tm.start()].strip(" ,").replace(",", "").strip():
                break
            p.acts.append(f"{prefix}-{tm.group(1)}")
        rest = rest.replace(full, " ")

    # Установка ищется в остатке строки, чтобы не зацепить кусок номера акта.
    um = _UNIT_RE.search(rest)
    if um:
        p.unit = um.group(1)

    # Линия: сперва явное «Линия №NNN», иначе «№NNN» в остатке без номеров
    # актов (чтобы «№222-1-143» не принялся за линию «222»).
    lwm = _LINE_WORD_RE.search(stem)
    if lwm:
        p.line = lwm.group(1)
    else:
        lm = _LINE_RE.search(rest)
        if lm:
            p.line = lm.group(1)

    dm = _DATE_RE.search(stem)
    if dm:
        p.date = dm.group(1)

    return p


def match_point(parsed: ParsedName, unit: str, line: str) -> str:
    """
    Светофор привязки документа к пункту по данным из имени файла.
    unit / line — установка и № трубопровода из карточки пункта.
    """
    checks = []
    if parsed.unit and unit:
        checks.append(parsed.unit == str(unit).strip())
    if parsed.line and line:
        checks.append(parsed.line == str(line).strip())

    if not checks:
        return MATCH_UNKNOWN
    return MATCH_OK if all(checks) else MATCH_MISMATCH


def _find_subfolder(base: str, target: str, cutoff: float = 0.7) -> str | None:
    """Нечёткий поиск подпапки (как в stage2_worker, без импорта воркера)."""
    if not os.path.isdir(base):
        return None
    dirs = {d.lower(): d for d in os.listdir(base)
            if os.path.isdir(os.path.join(base, d))}
    if target.lower() in dirs:
        return os.path.join(base, dirs[target.lower()])
    m = difflib.get_close_matches(target.lower(), dirs.keys(), n=1, cutoff=cutoff)
    return os.path.join(base, dirs[m[0]]) if m else None


def _pdfs_in(folder: str | None) -> list[str]:
    if not folder or not os.path.isdir(folder):
        return []
    out = []
    for f in sorted(os.listdir(folder)):
        p = os.path.join(folder, f)
        if os.path.isfile(p) and f.lower().endswith(".pdf"):
            out.append(p)
    return out


def find_repair_docs(point_folder: str, config: dict | None = None) -> RepairDocs:
    """
    Обнаружить документы ремонта в папке пункта.
    Папки ищутся по именам из config["DOC_FOLDERS"] (нечётко).
    """
    cfg = config or REPAIR_CONFIG
    names = cfg["DOC_FOLDERS"]  # ["Акт отбраковки", "Исполнительная"]

    docs = RepairDocs()
    docs.act_files = _pdfs_in(_find_subfolder(point_folder, names[0]))
    docs.id_files = _pdfs_in(_find_subfolder(point_folder, names[1]))
    docs.has_repair = bool(docs.act_files or docs.id_files)
    return docs


def build_records_from_docs(docs: RepairDocs) -> list["RepairRecord"]:
    """
    Предзаполнить реестр ремонтов из обнаруженных документов (по именам
    файлов). Правила:
      - каждый файл ИД → своя запись (статус «есть»; дату сдачи сотрудник
        впишет с превью реестра ИД — из имени файла её не взять);
      - отдельный файл акта с номером, уже входящим в какую-то ИД, —
        прикладывается к той записи (дедупликация, дата акта из имени);
      - акт без ИД → своя запись со статусом «ожидается».
    Всё это только черновик для формы — сотрудник правит и подтверждает.
    """
    from repair.model import ActRef, RepairRecord, ID_PRESENT, ID_EXPECTED

    records: list[RepairRecord] = []
    act_index: dict[str, ActRef] = {}   # номер акта → его ActRef в записях

    for f in docs.id_files:
        p = parse_doc_filename(f)
        rec = RepairRecord(id_status=ID_PRESENT,
                           source_files=[os.path.basename(f)])
        for num in dedupe_acts(p.acts):
            ref = ActRef(number=num, date="")
            rec.acts.append(ref)
            act_index.setdefault(num, ref)
        records.append(rec)

    for f in docs.act_files:
        p = parse_doc_filename(f)
        nums = dedupe_acts(p.acts)
        if not nums:
            # имя не разобралось — пустая запись-заготовка с файлом
            records.append(RepairRecord(id_status=ID_EXPECTED,
                                        source_files=[os.path.basename(f)]))
            continue
        for num in nums:
            if num in act_index:
                # копия акта из ИД: дополняем дату акта, файл — к той записи
                ref = act_index[num]
                if p.date and not ref.date:
                    ref.date = p.date
                for rec in records:
                    if ref in rec.acts and os.path.basename(f) not in rec.source_files:
                        rec.source_files.append(os.path.basename(f))
            else:
                ref = ActRef(number=num, date=p.date)
                act_index[num] = ref
                records.append(RepairRecord(
                    id_status=ID_EXPECTED, acts=[ref],
                    source_files=[os.path.basename(f)]))
    return records


def dedupe_acts(act_numbers: list[str]) -> list[str]:
    """
    Дедупликация номеров актов с сохранением порядка: отдельный файл акта
    с тем же номером, что копия внутри ИД, — один и тот же ремонт.
    """
    seen: set[str] = set()
    out: list[str] = []
    for n in act_numbers:
        key = n.strip()
        if key and key not in seen:
            seen.add(key)
            out.append(key)
    return out
