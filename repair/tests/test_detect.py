# -*- coding: utf-8 -*-
"""
Тесты обнаружения ремонта и привязки документов по именам файлов.
Имена — реальные, из эталонных пунктов п.2701 и п.2775.
"""

import os

from repair.detect import (
    parse_doc_filename, match_point, find_repair_docs, dedupe_acts,
    MATCH_OK, MATCH_MISMATCH, MATCH_UNKNOWN,
)


# ─── Разбор имён файлов (реальные примеры) ────────────────────────────

def test_parse_2701_id_file():
    p = parse_doc_filename("36-30, №6, 222-1-84.pdf")
    assert p.acts == ["222-1-84"]
    assert p.unit == "36-30"
    assert p.line == "6"


def test_parse_2775_act_file():
    p = parse_doc_filename("222-1-121 №122 39-40 24.03.2026.pdf")
    assert p.acts == ["222-1-121"]
    assert p.unit == "39-40"
    assert p.line == "122"
    assert p.date == "24.03.2026"


def test_parse_2775_id_two_acts_comma():
    """«…№222-1-143,121…» — сокращение: второй акт наследует префикс."""
    p = parse_doc_filename(
        "ИД по Акту ОТН №222-1-143,121. Линия №122.pdf_v00.100001268600097.pdf")
    assert "222-1-143" in p.acts
    assert "222-1-121" in p.acts
    assert p.is_id
    assert p.line == "122"


def test_parse_line_number_is_line():
    """В имени два «№…»: «№222-1-143» (акт) и «Линия №122». Линия — 122,
    номер акта не должен приняться за линию «222»."""
    p = parse_doc_filename("ИД по Акту ОТН №222-1-143,121. Линия №122.pdf")
    assert p.line == "122"


def test_parse_no_data():
    p = parse_doc_filename("скан 001.pdf")
    assert p.acts == []
    assert p.unit == ""
    assert p.line == ""


# ─── Светофор привязки ────────────────────────────────────────────────

def test_match_ok():
    p = parse_doc_filename("36-30, №6, 222-1-84.pdf")
    assert match_point(p, "36-30", "6") == MATCH_OK


def test_match_mismatch_other_point():
    """Акт другого пункта (линия 122, уст. 39-40) против пункта 36-30/№6."""
    p = parse_doc_filename("222-1-121 №122 39-40 24.03.2026.pdf")
    assert match_point(p, "36-30", "6") == MATCH_MISMATCH


def test_match_unknown_when_name_empty():
    p = parse_doc_filename("скан 001.pdf")
    assert match_point(p, "36-30", "6") == MATCH_UNKNOWN


# ─── Обнаружение по папкам ────────────────────────────────────────────

def _mk(tmp_path, folder, files=()):
    d = tmp_path / folder
    d.mkdir(parents=True, exist_ok=True)
    for f in files:
        (d / f).write_bytes(b"%PDF-1.4 test")
    return d


def test_find_repair_docs_both(tmp_path):
    _mk(tmp_path, "Акт отбраковки", ["a1.pdf"])
    _mk(tmp_path, "Исполнительная", ["id1.pdf", "заметка.txt"])
    docs = find_repair_docs(str(tmp_path))
    assert docs.has_repair
    assert len(docs.act_files) == 1
    assert len(docs.id_files) == 1          # .txt не считается
    assert all(f.lower().endswith(".pdf") for f in docs.all_files)


def test_find_repair_docs_empty_folder_is_no_repair(tmp_path):
    _mk(tmp_path, "Акт отбраковки", [])     # папка есть, но пустая
    docs = find_repair_docs(str(tmp_path))
    assert not docs.has_repair


def test_find_repair_docs_absent(tmp_path):
    _mk(tmp_path, "Паспорт", ["п.pdf"])
    docs = find_repair_docs(str(tmp_path))
    assert not docs.has_repair


def test_find_repair_docs_fuzzy_folder_name(tmp_path):
    """Папку с опечаткой находим нечётким поиском (как в остальном ПНОС)."""
    _mk(tmp_path, "Акты отбраковки", ["a.pdf"])   # опечатка в имени папки
    docs = find_repair_docs(str(tmp_path))
    assert docs.has_repair


# ─── Дедупликация актов ───────────────────────────────────────────────

def test_dedupe_keeps_order():
    acts = ["222-1-84", "222-1-309", "222-1-84", " 222-1-309 "]
    assert dedupe_acts(acts) == ["222-1-84", "222-1-309"]


# ─── Предзаполнение реестра из документов ─────────────────────────────

from repair.detect import build_records_from_docs, RepairDocs
from repair.model import ID_PRESENT, ID_EXPECTED


def test_prefill_2701_two_ids():
    """п.2701: два комплекта ИД → две записи «есть», по акту в каждой."""
    docs = RepairDocs(has_repair=True, act_files=[], id_files=[
        r"C:\x\Исполнительная\36-30, №6, 222-1-84.pdf",
        r"C:\x\Исполнительная\36-30, №6, 222-1-309.pdf",
    ])
    recs = build_records_from_docs(docs)
    assert len(recs) == 2
    assert all(r.id_status == ID_PRESENT for r in recs)
    assert [r.acts[0].number for r in recs] == ["222-1-84", "222-1-309"]
    assert recs[0].date_done == ""              # дату сдачи впишет сотрудник


def test_prefill_2775_one_id_two_acts():
    """п.2775: ИД на оба акта + отдельные чёткие акты → ОДНА запись,
    даты актов подтянуты из имён файлов актов, файлы приложены."""
    docs = RepairDocs(has_repair=True, act_files=[
        r"C:\x\Акт отбраковки\222-1-121 №122 39-40 24.03.2026.pdf",
        r"C:\x\Акт отбраковки\222-1-143 № 122 39-40 30.03.2026.pdf",
    ], id_files=[
        r"C:\x\Исполнительная\ИД по Акту ОТН №222-1-143,121. Линия №122.pdf",
    ])
    recs = build_records_from_docs(docs)
    assert len(recs) == 1
    rec = recs[0]
    assert rec.id_status == ID_PRESENT
    got = {a.number: a.date for a in rec.acts}
    assert got == {"222-1-143": "30.03.2026", "222-1-121": "24.03.2026"}
    assert len(rec.source_files) == 3           # ИД + оба акта


def test_prefill_act_only():
    docs = RepairDocs(has_repair=True, act_files=[
        r"C:\x\Акт отбраковки\222-1-121 №122 39-40 24.03.2026.pdf",
    ], id_files=[])
    recs = build_records_from_docs(docs)
    assert len(recs) == 1
    assert recs[0].id_status == ID_EXPECTED
    assert recs[0].acts[0].date == "24.03.2026"


def test_prefill_unparsed_name_kept():
    """Нечитаемое имя не теряется — пустая запись-заготовка с файлом."""
    docs = RepairDocs(has_repair=True,
                      act_files=[r"C:\x\Акт отбраковки\скан 001.pdf"], id_files=[])
    recs = build_records_from_docs(docs)
    assert len(recs) == 1
    assert recs[0].acts == []
    assert recs[0].source_files == ["скан 001.pdf"]
