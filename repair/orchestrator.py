# -*- coding: utf-8 -*-
"""
Оркестратор применения ремонта (принципы П-2/П-4).

Жёсткая последовательность:
    бэкап затрагиваемых файлов → запись → журнал.
Ошибка записи → автооткат из только что сделанного бэкапа → журнал об откате.

Сам оркестратор ничего не знает про Excel: писатель передаётся снаружи
(боевой — repair.com_writer.apply_to_workbook, в тестах — фейк). Список
операций к этому моменту уже согласован сотрудником на экране предпросмотра —
сюда приходит только то, что осталось с включёнными галочками.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from repair.imports import WriteOp
from repair.model import RepairRecord
from repair.safety import make_backup, restore_backup, append_journal


@dataclass
class ApplyResult:
    ok: bool
    backup_dir: str = ""
    journal_path: str = ""
    error: str = ""
    counts: dict = field(default_factory=dict)
    restored: list = field(default_factory=list)


def _records_summary(records: list[RepairRecord]) -> list[str]:
    lines = []
    for i, r in enumerate(records, start=1):
        acts = ", ".join(f"№ {a.number} от {a.date}".strip() for a in r.acts if a.is_valid())
        tail = f" (ИД сдана {r.date_done})" if r.has_id else f" (ИД: {r.id_status})"
        lines.append(f"источник {i}: акт {acts}{tail}")
    return lines


def apply_repair(
    point_folder: str,
    xlsm_path: str,
    ops: list[WriteOp],
    records: list[RepairRecord],
    writer,
    extra_files: list[str] | None = None,
    log=None,
) -> ApplyResult:
    """
    point_folder — папка пункта (там живут бэкапы и журнал);
    xlsm_path    — ПНОС_ТТП.xlsm пункта;
    ops          — согласованные операции записи (после галочек предпросмотра);
    records      — реестр ремонтов для листа «Ремонт»;
    writer       — callable(xlsm_path, ops, records, log=...) -> dict счётчиков;
    extra_files  — дополнительные файлы в бэкап (карта PNG, отчёт Word — на
                   будущие шаги; сейчас обычно только xlsm).
    """
    def _log(msg: str):
        if log:
            log(msg)

    result = ApplyResult(ok=False)

    files = [xlsm_path] + list(extra_files or [])
    result.backup_dir = make_backup(point_folder, files)
    _log(f"Бэкап: {result.backup_dir}")

    try:
        result.counts = writer(xlsm_path, ops, records, log=log) or {}
    except Exception as e:
        result.error = str(e)
        _log(f"Ошибка записи: {e}")
        result.restored = restore_backup(result.backup_dir, point_folder)
        _log(f"Откат: восстановлено файлов {len(result.restored)}")
        result.journal_path = append_journal(
            point_folder, "Применение ремонта — ОШИБКА, выполнен откат",
            _records_summary(records) + [
                f"ошибка: {result.error}",
                f"восстановлено файлов из бэкапа: {len(result.restored)}",
                f"бэкап: {result.backup_dir}",
            ])
        return result

    result.ok = True
    c = result.counts
    result.journal_path = append_journal(
        point_folder, "Применение ремонта",
        _records_summary(records) + [
            f"«Импорт»: ячеек {c.get('import', 0)}; "
            f"«Импорт (ПР)»: ячеек {c.get('import_pr', 0)}; "
            f"«Ремонт»: записей {c.get('registry', 0)}",
            f"операций согласовано: {len(ops)}",
            f"бэкап: {result.backup_dir}",
        ])
    _log("Применение завершено, журнал дописан.")
    return result
