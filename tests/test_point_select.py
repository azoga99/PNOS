# -*- coding: utf-8 -*-
"""
Выбор пунктов для этапов 2–7.

Этапы перебирают папки «п.*» в ПНОС; с параметром points — только папки
с этими номерами. Без выбора всё как раньше: все пункты, тот же порядок.
"""

import os

import pytest

from status_manager import (missing_points, no_selected_points_message, point_number_of,
                            select_point_folders, set_stage_status)
from workers.stage2_worker import Stage2Worker
from workers.stage3_worker import Stage3Worker
from workers.stage4_worker import Stage4Worker
from workers.stage5_worker import Stage5Worker
from workers.stage6_worker import Stage6Worker

WORKERS = [Stage2Worker, Stage3Worker, Stage4Worker, Stage5Worker, Stage6Worker]


@pytest.mark.parametrize("name, num", [
    ("п.3186", "3186"), ("п. 3186", "3186"), ("П.3186", "3186"), ("3186", "3186"),
    ("пункт 3186", "3186"), ("п.3186 (не корректный)", "3186"), ("п.31860", "31860"),
    ("_шаблоны", None), ("_gi_cache", None),
])
def test_point_number_of(name, num):
    assert point_number_of(name) == num


def test_stage1_still_exports_point_number_of():
    from workers.stage1_worker import point_number_of as from_stage1
    assert from_stage1 is point_number_of


@pytest.fixture
def root(tmp_path):
    for name in ["п.3186", "п. 2758", "п.2704 (не корректный)", "п.31860",
                 "_шаблоны", "_gi_cache", "3186", "П.1111", "Отчёты"]:
        (tmp_path / name).mkdir()
    (tmp_path / "п.9999.txt").write_text("не папка", encoding="utf-8")
    return tmp_path


def _old_listing(local_root):
    """Как этапы 2–7 перебирали пункты до появления выбора."""
    return [os.path.join(local_root, d) for d in os.listdir(local_root)
            if os.path.isdir(os.path.join(local_root, d)) and d.startswith("п.")]


def _names(folders):
    return [os.path.basename(f) for f in folders]


def test_all_points_as_before(root):
    assert select_point_folders(str(root)) == _old_listing(str(root))
    assert sorted(_names(select_point_folders(str(root)))) == sorted(
        ["п.3186", "п. 2758", "п.2704 (не корректный)", "п.31860"])


def test_filter_keeps_listing_order(root):
    got = select_point_folders(str(root), ["2758", "3186", "2704"])
    expected = [f for f in _old_listing(str(root))
                if os.path.basename(f) in ("п.3186", "п. 2758", "п.2704 (не корректный)")]
    assert got == expected


def test_filter_by_exact_number(root):
    # п.31860 — другой пункт; «3186» и «П.1111» этапы не брали и не берут
    assert _names(select_point_folders(str(root), ["3186"])) == ["п.3186"]
    assert select_point_folders(str(root), ["1111"]) == []


def test_points_are_normalized(root):
    assert _names(select_point_folders(str(root), [" п.3186 ", 3186])) == ["п.3186"]


def test_blocked_point_is_returned(root):
    folder = root / "п.3186"
    set_stage_status(str(folder), "skipped", True)
    assert _names(select_point_folders(str(root), ["3186"])) == ["п.3186"]
    assert _names(select_point_folders(str(root), ["2704"])) == ["п.2704 (не корректный)"]


def test_empty_selection(root):
    assert select_point_folders(str(root), []) == []
    assert select_point_folders(str(root), ["5555"]) == []


def test_missing_root(tmp_path):
    assert select_point_folders(str(tmp_path / "нет такой")) == []
    assert select_point_folders(str(tmp_path / "нет такой"), ["3186"]) == []


def test_missing_points(root):
    folders = select_point_folders(str(root), ["3186", "5555"])
    assert missing_points(folders, ["3186", "5555"]) == ["5555"]


def test_no_selected_message():
    assert no_selected_points_message(["3186", "2758"]) == \
        "Выбранных пунктов нет в папке ПНОС (п.3186, п.2758)"


@pytest.mark.parametrize("cls", WORKERS)
def test_worker_accepts_points(cls, tmp_path):
    assert cls(str(tmp_path)).points is None
    w = cls(str(tmp_path), points=("3186", "2758"))
    assert w.points == ["3186", "2758"]


def _run(cls, root, tmp_path, points):
    kwargs = {"points": points}
    if cls is Stage3Worker:
        master = tmp_path / "master.xlsm"
        master.write_bytes(b"x")
        kwargs["macro_master_path"] = str(master)
    w = cls(str(root), **kwargs)
    infos, logs, results = [], [], []
    w.info.connect(lambda m, c: infos.append((m, c)))
    w.log.connect(logs.append)
    w.finished_ok.connect(results.append)
    w.run()   # синхронно: до COM и сети дело не доходит
    return infos, logs, results


@pytest.mark.parametrize("cls", WORKERS)
def test_worker_warns_when_nothing_selected(cls, root, tmp_path):
    infos, logs, results = _run(cls, root, tmp_path, ["5555"])
    msg = "Выбранных пунктов нет в папке ПНОС (п.5555)"
    assert infos == [(msg, "warn")]
    assert f"⚠ {msg}" in logs
    assert results == [False]


@pytest.mark.parametrize("cls", WORKERS)
def test_worker_without_points_keeps_old_message(cls, tmp_path):
    empty = tmp_path / "ПНОС"
    empty.mkdir()
    infos, logs, results = _run(cls, empty, tmp_path, None)
    assert infos == []
    assert any(line.startswith("⚠ ") and "пункт" in line.lower() for line in logs)
    assert results == [False]
