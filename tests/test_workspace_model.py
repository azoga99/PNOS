# -*- coding: utf-8 -*-
"""Строки «Рабочего места» по папкам пунктов и их .pnos_status.json."""

import json

from workspace_model import STEP_KEYS, build_row, build_rows, next_step


def _point(root, name, **status):
    folder = root / name
    folder.mkdir()
    if status:
        (folder / ".pnos_status.json").write_text(
            json.dumps(status, ensure_ascii=False), encoding="utf-8")
    return folder


def _done(*steps):
    return {STEP_KEYS[s]: True for s in steps}


def _by_point(rows):
    return {r["point"]: r for r in rows}


def test_missing_root_gives_no_rows(tmp_path):
    assert build_rows(str(tmp_path / "нет")) == []
    assert build_rows("") == []


def test_only_point_folders_sorted_by_number(tmp_path):
    _point(tmp_path, "п.3186")
    _point(tmp_path, "п.759")
    _point(tmp_path, "п 2758")
    _point(tmp_path, "_шаблоны")
    _point(tmp_path, "_gi_cache")
    _point(tmp_path, "п.без номера")
    (tmp_path / "п.1.txt").write_text("")
    assert [r["point"] for r in build_rows(str(tmp_path))] == ["759", "2758", "3186"]


def test_fresh_point_waits_for_step_one(tmp_path):
    row = build_row(str(_point(tmp_path, "п.3186")))
    assert row["point"] == "3186"
    assert row["steps"] == {s: "idle" for s in range(1, 8)}
    assert row["state"] == "Следующий шаг: 1 · Структура"
    assert row["state_kind"] == "info"


def test_done_steps_and_next_step(tmp_path):
    row = build_row(str(_point(tmp_path, "п.1", **_done(1, 2, 3, 4))))
    assert [row["steps"][s] for s in range(1, 8)] == ["done"] * 4 + ["idle"] * 3
    assert row["state"] == "Следующий шаг: 5 · Отчёт"
    assert next_step(row["steps"]) == 5


def test_step_two_done_waits_for_form(tmp_path):
    row = build_row(str(_point(tmp_path, "п.3186", **_done(1, 2))))
    assert row["steps"][3] == "wait"
    assert row["steps"][4] == "idle"
    assert row["state"] == "Ждёт формы"
    assert row["state_kind"] == "warn"


def test_all_seven_done_is_ready(tmp_path):
    row = build_row(str(_point(tmp_path, "п.2758", **_done(*range(1, 8)))))
    assert set(row["steps"].values()) == {"done"}
    assert (row["state"], row["state_kind"]) == ("Готов", "done")


def test_invalid_point_lists_missing_folders(tmp_path):
    folder = _point(tmp_path, "п.3189 (не корректный)", invalid=True,
                    missing_folders=["Паспорт", "Первичка"])
    row = build_row(str(folder))
    assert row["point"] == "3189"
    assert row["steps"][1] == "error"
    assert row["state"] == "Не корректный: нет Паспорт, Первичка"
    assert row["state_kind"] == "error"


def test_invalid_by_name_only(tmp_path):
    row = build_row(str(_point(tmp_path, "п.3189 (не корректный)")))
    assert row["steps"][1] == "error"
    assert row["state"] == "Не корректный"


def test_skipped_point_marks_first_unfinished_step(tmp_path):
    row = build_row(str(_point(tmp_path, "п.2760", skipped=True, **_done(1, 2))))
    assert row["steps"][3] == "error"
    assert row["steps"][4] == "idle"
    assert (row["state"], row["state_kind"]) == ("Пропущен", "warn")


def test_finished_point_is_ready_even_if_skipped_later(tmp_path):
    row = build_row(str(_point(tmp_path, "п.1", skipped=True, **_done(*range(1, 8)))))
    assert row["state"] == "Готов"


def test_broken_status_file_counts_as_empty(tmp_path):
    folder = _point(tmp_path, "п.5")
    (folder / ".pnos_status.json").write_text("{не json", encoding="utf-8")
    row = build_row(str(folder))
    assert row["steps"][1] == "idle"


def test_values_other_than_true_count_as_done(tmp_path):
    # Ключи шагов хранят True, но set_stage_status допускает любое значение
    row = build_row(str(_point(tmp_path, "п.5", stage1="да", stage2=0)))
    assert row["steps"][1] == "done"
    assert row["steps"][2] == "idle"


def test_plain_folder_wins_over_invalid_twin(tmp_path):
    _point(tmp_path, "п.3189 (не корректный)", invalid=True)
    plain = _point(tmp_path, "п.3189", **_done(1))
    rows = build_rows(str(tmp_path))
    assert len(rows) == 1
    assert rows[0]["folder"] == str(plain)
    assert rows[0]["steps"][1] == "done"


def test_step_seven_without_acts_waits_with_reason(tmp_path):
    folder = _point(tmp_path, "п.3186", **_done(1, 2, 3, 4, 5, 6),
                    stage6_wait="Нет актов ГИ (нет папки для 21-20)")
    row = build_row(str(folder))
    assert row["steps"][7] == "wait"
    assert row["state"].startswith("Нет актов ГИ (нет папки для 21-20)")
    assert row["state_kind"] == "warn"


def test_step_seven_done_ignores_stale_wait(tmp_path):
    folder = _point(tmp_path, "п.3186", **_done(*range(1, 8)),
                    stage6_wait="Нет актов ГИ")
    row = build_row(str(folder))
    assert row["steps"][7] == "done"
    assert row["state"] == "Готов"
