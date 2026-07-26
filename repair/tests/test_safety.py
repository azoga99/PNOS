# -*- coding: utf-8 -*-
"""Тесты бэкапа, отката и журнала."""

import os
from datetime import datetime

from repair.safety import (
    make_backup, restore_backup, latest_backup, append_journal,
    BACKUP_DIR, JOURNAL_NAME,
)


def _point(tmp_path):
    xlsm = tmp_path / "ПНОС_ТТП.xlsm"
    xlsm.write_bytes(b"original-excel")
    png = tmp_path / "2701.png"
    png.write_bytes(b"original-map")
    return str(tmp_path), str(xlsm), str(png)


def test_backup_and_restore(tmp_path):
    folder, xlsm, png = _point(tmp_path)

    b = make_backup(folder, [xlsm, png], now=datetime(2026, 7, 22, 14, 0, 0))
    assert os.path.isdir(b)
    assert sorted(os.listdir(b)) == ["2701.png", "ПНОС_ТТП.xlsm"]

    # «запись» испортила файлы
    open(xlsm, "wb").write(b"broken")
    open(png, "wb").write(b"broken")

    restored = restore_backup(b, folder)
    assert len(restored) == 2
    assert open(xlsm, "rb").read() == b"original-excel"
    assert open(png, "rb").read() == b"original-map"


def test_backup_skips_missing_files(tmp_path):
    folder, xlsm, _ = _point(tmp_path)
    b = make_backup(folder, [xlsm, str(tmp_path / "нет_такого.docx")])
    assert os.listdir(b) == ["ПНОС_ТТП.xlsm"]


def test_latest_backup_orders_by_time(tmp_path):
    folder, xlsm, _ = _point(tmp_path)
    make_backup(folder, [xlsm], now=datetime(2026, 7, 22, 10, 0, 0))
    b2 = make_backup(folder, [xlsm], now=datetime(2026, 7, 22, 12, 0, 0))
    assert latest_backup(folder) == b2


def test_latest_backup_none(tmp_path):
    assert latest_backup(str(tmp_path)) is None


def test_prune_keeps_last(tmp_path):
    folder, xlsm, _ = _point(tmp_path)
    for h in range(13):
        make_backup(folder, [xlsm], now=datetime(2026, 7, 22, 0, h % 24, h))
    root = os.path.join(folder, BACKUP_DIR)
    assert len(os.listdir(root)) == 10          # KEEP_BACKUPS


def test_journal_appends(tmp_path):
    folder, _, _ = _point(tmp_path)
    p1 = append_journal(folder, "Применение ремонта",
                        ["Импорт: 43 сечения, 257 ячеек"],
                        now=datetime(2026, 7, 22, 14, 5))
    append_journal(folder, "Откат", ["восстановлено 2 файла"],
                   now=datetime(2026, 7, 22, 14, 30))

    assert os.path.basename(p1) == JOURNAL_NAME
    text = open(p1, encoding="utf-8").read()
    assert "22.07.2026 14:05 · Применение ремонта" in text
    assert "Импорт: 43 сечения, 257 ячеек" in text
    assert "22.07.2026 14:30 · Откат" in text
    # записи разделены пустой строкой, порядок хронологический
    assert text.index("Применение") < text.index("Откат")
