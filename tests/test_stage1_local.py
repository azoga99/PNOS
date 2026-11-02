# -*- coding: utf-8 -*-
"""
Этап 1: комплект, который сотрудник сам сложил в папку ПНОС.

С Яндекс.Диска многое перестало скачиваться, и пункт проще собрать вручную.
Раньше такой пункт оставался заблокированным: прошлый прогон пометил его
«не корректный» (папки ещё не было), а повторный этап 1 такие пункты
пропускал не глядя. Теперь полный комплект берётся как есть — без Яндекса,
пометка снимается, досылаются только шаблоны.
"""

import asyncio
import json
import os

import pytest

from config import CONFIG
from status_manager import get_stage_status, get_status_value, set_stage_status
from workers import stage1_worker
from workers.stage1_worker import (Stage1Worker, TEMPLATE_CACHE, clear_invalid_mark,
                                   find_local_point, missing_required,
                                   normalize_point_folder, point_number_of)

SUFFIX = CONFIG["INVALID_SUFFIX"]


def _point(root, name, passport=True, primary=True, templates=True, status=None):
    folder = root / name
    folder.mkdir()
    if passport:
        (folder / "Паспорт").mkdir()
        (folder / "Паспорт" / "Паспорт трубопровода.docx").write_bytes(b"p")
    if primary:
        (folder / "Первичка").mkdir()
        (folder / "Первичка" / "3186._1.png").write_bytes(b"m")
    if templates:
        for n in CONFIG["TEMPLATE_FILES"]:
            (folder / n).write_bytes(b"own")
    for k, v in (status or {}).items():
        set_stage_status(str(folder), k, v)
    return folder


@pytest.mark.parametrize("name, num", [
    ("п.3186", "3186"), ("п. 3186", "3186"), ("П.3186", "3186"), ("3186", "3186"),
    ("пункт 3186", "3186"), ("п.3186 (не корректный)", "3186"), ("п.3186(ШАР)", "3186"),
    ("_шаблоны", None), ("_gi_cache", None), ("Отчёты", None),
])
def test_point_number_of(name, num):
    assert point_number_of(name) == num


def test_number_is_not_a_prefix():
    # п.31860 — другой пункт, не 3186
    assert point_number_of("п.31860") == "31860"


def test_find_prefers_complete_folder(tmp_path):
    _point(tmp_path, "п.3186", passport=False, primary=False)
    full = _point(tmp_path, "п.3186" + SUFFIX)
    assert find_local_point(str(tmp_path), "3186") == str(full)


def test_find_prefers_exact_name_among_complete(tmp_path):
    exact = _point(tmp_path, "п.3186")
    _point(tmp_path, "п. 3186")
    assert find_local_point(str(tmp_path), "3186") == str(exact)


def test_find_nothing(tmp_path):
    _point(tmp_path, "п.3187")
    assert find_local_point(str(tmp_path), "3186") is None


def test_missing_required_counts_empty_folders(tmp_path):
    folder = _point(tmp_path, "п.3186", primary=False)
    (folder / "Первичка").mkdir()
    assert missing_required(str(folder)) == ["Первичка"]


def test_normalize_renames(tmp_path):
    folder = _point(tmp_path, "п. 3186" + SUFFIX)
    new, err = normalize_point_folder(str(tmp_path), str(folder), "3186")
    assert err is None and new == str(tmp_path / "п.3186") and os.path.isdir(new)


def test_normalize_does_not_overwrite(tmp_path):
    _point(tmp_path, "п.3186")
    other = _point(tmp_path, "п.3186" + SUFFIX)
    new, err = normalize_point_folder(str(tmp_path), str(other), "3186")
    assert new == str(other) and "уже есть" in err


def test_clear_invalid_keeps_other_keys(tmp_path):
    folder = _point(tmp_path, "п.3186", status={
        "stage1": True, "invalid": True, "missing_folders": ["Паспорт"],
        "manual3_draft": {"B14": "транспортировки нефти"}})
    assert clear_invalid_mark(str(folder)) is True
    assert not get_stage_status(str(folder), "invalid")
    assert get_status_value(str(folder), "missing_folders") is None
    assert get_status_value(str(folder), "manual3_draft") == {"B14": "транспортировки нефти"}
    assert get_stage_status(str(folder), "stage1") is True
    assert clear_invalid_mark(str(folder)) is False


# ─── Весь этап 1 на своих комплектах ────────────────────────────────

def _run(tmp_path, monkeypatch, points, connection=None):
    """Этап 1 без сети: check_connection подменён, остальные обращения к диску — ошибка."""
    calls = []

    async def fake_check(self, session):
        calls.append("check")
        return connection

    async def no_network(*a, **k):
        raise AssertionError("этап 1 полез на Яндекс.Диск")

    monkeypatch.setattr(stage1_worker.YandexDiskAPI, "check_connection", fake_check)
    monkeypatch.setattr(stage1_worker.YandexDiskAPI, "get_folder_contents", no_network)
    monkeypatch.setattr(stage1_worker.YandexDiskAPI, "preload_search_paths", no_network)

    w = Stage1Worker(excel_path=None, local_root=str(tmp_path), manual_points=points)
    out = {"info": []}
    w.report_ready.connect(lambda r: out.__setitem__("report", r))
    w.finished_ok.connect(lambda ok: out.__setitem__("ok", ok))
    w.info.connect(lambda m, c: out["info"].append((c, m)))
    asyncio.run(w.async_run())
    out["calls"] = calls
    return out


def test_own_set_marked_invalid_is_accepted_without_yandex(tmp_path, monkeypatch):
    # Ровно случай п.3186: прошлый прогон пометил пункт, потом сотрудник
    # положил Паспорт и Первичку сам
    _point(tmp_path, "п.3186" + SUFFIX, status={
        "stage1": True, "invalid": True, "missing_folders": ["Паспорт", "Первичка"]})
    out = _run(tmp_path, monkeypatch, ["3186"])

    folder = tmp_path / "п.3186"
    assert folder.is_dir() and not (tmp_path / ("п.3186" + SUFFIX)).exists()
    data = json.loads((folder / ".pnos_status.json").read_text(encoding="utf-8"))
    assert data == {"stage1": True}
    assert out["ok"] is True and out["calls"] == []  # шаблоны на месте — сеть не нужна
    (detail,) = out["report"]["details"]
    assert detail["status"] == "✓ Свой комплект"
    assert detail["folders"]["Паспорт"] and detail["folders"]["Первичка"]


def test_own_templates_are_not_overwritten(tmp_path, monkeypatch):
    folder = _point(tmp_path, "п.3186")
    cache = tmp_path / TEMPLATE_CACHE
    cache.mkdir()
    for n in CONFIG["TEMPLATE_FILES"]:
        (cache / n).write_bytes(b"cache")
    _run(tmp_path, monkeypatch, ["3186"])
    for n in CONFIG["TEMPLATE_FILES"]:
        assert (folder / n).read_bytes() == b"own"


def test_templates_from_saved_copy_when_offline(tmp_path, monkeypatch):
    folder = _point(tmp_path, "п. 3186", templates=False)
    cache = tmp_path / TEMPLATE_CACHE
    cache.mkdir()
    for n in CONFIG["TEMPLATE_FILES"]:
        (cache / n).write_bytes(b"cache")
    out = _run(tmp_path, monkeypatch, ["3186"], connection="нет связи")

    folder = tmp_path / "п.3186"
    for n in CONFIG["TEMPLATE_FILES"]:
        assert (folder / n).read_bytes() == b"cache"
    assert out["ok"] is True
    assert out["report"]["details"][0]["status"] == "✓ Свой комплект"


def test_no_templates_anywhere_is_reported(tmp_path, monkeypatch):
    _point(tmp_path, "п.3186", templates=False)
    out = _run(tmp_path, monkeypatch, ["3186"], connection="нет связи")
    assert out["report"]["details"][0]["status"].startswith("⚠ Нет шаблонов")
    assert any(c == "warn" and "шаблонов" in m for c, m in out["info"])


def test_incomplete_point_offline_is_not_downloaded(tmp_path, monkeypatch):
    _point(tmp_path, "п.3186")
    _point(tmp_path, "п.3187", primary=False)
    out = _run(tmp_path, monkeypatch, ["3186", "3187"], connection="нет связи")
    statuses = {d["point"]: d["status"] for d in out["report"]["details"]}
    assert statuses == {"3186": "✓ Свой комплект",
                        "3187": "Не скачан — нет связи с Яндекс.Диском"}
    assert out["ok"] is True  # свой комплект подготовлен, хотя второй пункт не скачан
    assert any(c == "error" and "Нет связи" in m for c, m in out["info"])
