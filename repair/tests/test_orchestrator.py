# -*- coding: utf-8 -*-
"""Тесты оркестратора применения: успех, ошибка с автооткатом."""

import os

from repair.imports import WriteOp
from repair.model import ActRef, RepairRecord, ID_PRESENT
from repair.orchestrator import apply_repair
from repair.safety import JOURNAL_NAME


def _setup(tmp_path):
    xlsm = tmp_path / "ПНОС_ТТП.xlsm"
    xlsm.write_bytes(b"original")
    records = [RepairRecord(id_status=ID_PRESENT, date_done="13.04.2026",
                            acts=[ActRef("222-1-84", "24.02.2026")])]
    ops = [WriteOp(9, 4, 8.0), WriteOp(9, 6, 8.0)]
    return str(tmp_path), str(xlsm), ops, records


def test_apply_success(tmp_path):
    folder, xlsm, ops, records = _setup(tmp_path)

    def writer(path, ops_, records_, log=None):
        assert path == xlsm and len(ops_) == 2
        open(path, "wb").write(b"written")
        return {"import": 2, "import_pr": 2, "registry": 1}

    res = apply_repair(folder, xlsm, ops, records, writer)

    assert res.ok
    assert res.counts["import"] == 2
    assert os.path.isdir(res.backup_dir)
    # бэкап содержит исходник, файл — новый
    backed = open(os.path.join(res.backup_dir, "ПНОС_ТТП.xlsm"), "rb").read()
    assert backed == b"original"
    assert open(xlsm, "rb").read() == b"written"
    # журнал: источник и счётчики
    text = open(res.journal_path, encoding="utf-8").read()
    assert "Применение ремонта" in text
    assert "222-1-84" in text and "13.04.2026" in text
    assert "«Импорт»: ячеек 2" in text


def test_apply_failure_rolls_back(tmp_path):
    folder, xlsm, ops, records = _setup(tmp_path)

    def writer(path, ops_, records_, log=None):
        open(path, "wb").write(b"corrupted-half-written")
        raise RuntimeError("Excel умер на середине")

    res = apply_repair(folder, xlsm, ops, records, writer)

    assert not res.ok
    assert "Excel умер" in res.error
    # файл восстановлен из бэкапа
    assert open(xlsm, "rb").read() == b"original"
    assert len(res.restored) == 1
    # журнал фиксирует ошибку и откат
    text = open(os.path.join(folder, JOURNAL_NAME), encoding="utf-8").read()
    assert "ОШИБКА" in text and "откат" in text.lower()
    assert "восстановлено файлов из бэкапа: 1" in text


def test_apply_backs_up_extra_files(tmp_path):
    folder, xlsm, ops, records = _setup(tmp_path)
    png = tmp_path / "2701.png"
    png.write_bytes(b"map")

    def writer(path, ops_, records_, log=None):
        return {}

    res = apply_repair(folder, xlsm, ops, records, writer,
                       extra_files=[str(png)])
    assert sorted(os.listdir(res.backup_dir)) == ["2701.png", "ПНОС_ТТП.xlsm"]
