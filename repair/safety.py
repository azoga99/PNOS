# -*- coding: utf-8 -*-
"""
Безопасность применения (принцип П-4): бэкап перед записью, откат, журнал.

Последовательность применения жёсткая:
    бэкап затрагиваемых файлов → запись → журнал операции.
Любая ошибка на записи → автооткат из бэкапа (на вызывающей стороне).

Бэкапы лежат рядом с пунктом в подпапке «_backup/ГГГГ-ММ-ДД_ЧЧ-ММ-СС/»
с сохранением имён файлов. Журнал — обычный текстовый файл
«_журнал_ремонта.txt» в папке пункта: человекочитаемый, объясняет коллеге,
что и из чего было изменено.
"""

from __future__ import annotations

import os
import shutil
from datetime import datetime

BACKUP_DIR = "_backup"
JOURNAL_NAME = "_журнал_ремонта.txt"
KEEP_BACKUPS = 10   # хранить последних операций (настройка)


# ─── Бэкап и откат ────────────────────────────────────────────────────

def make_backup(point_folder: str, files: list[str], now: datetime | None = None) -> str:
    """
    Скопировать файлы в новую папку бэкапа. Возвращает путь к ней.
    Файлы, которых нет (ещё не созданы), молча пропускаются.
    """
    stamp = (now or datetime.now()).strftime("%Y-%m-%d_%H-%M-%S")
    dest = os.path.join(point_folder, BACKUP_DIR, stamp)
    os.makedirs(dest, exist_ok=True)
    for f in files:
        if os.path.isfile(f):
            shutil.copy2(f, os.path.join(dest, os.path.basename(f)))
    _prune_backups(os.path.join(point_folder, BACKUP_DIR))
    return dest


def restore_backup(backup_dir: str, point_folder: str) -> list[str]:
    """
    Вернуть файлы из папки бэкапа на исходные места (в корень пункта).
    Возвращает список восстановленных файлов.
    """
    restored = []
    if not os.path.isdir(backup_dir):
        return restored
    for name in os.listdir(backup_dir):
        src = os.path.join(backup_dir, name)
        if os.path.isfile(src):
            dst = os.path.join(point_folder, name)
            shutil.copy2(src, dst)
            restored.append(dst)
    return restored


def latest_backup(point_folder: str) -> str | None:
    """Путь к последнему бэкапу пункта (для кнопки «Откатить»)."""
    root = os.path.join(point_folder, BACKUP_DIR)
    if not os.path.isdir(root):
        return None
    dirs = sorted(d for d in os.listdir(root)
                  if os.path.isdir(os.path.join(root, d)))
    return os.path.join(root, dirs[-1]) if dirs else None


def _prune_backups(root: str, keep: int = KEEP_BACKUPS) -> None:
    if not os.path.isdir(root):
        return
    dirs = sorted(d for d in os.listdir(root)
                  if os.path.isdir(os.path.join(root, d)))
    for d in dirs[:-keep] if len(dirs) > keep else []:
        shutil.rmtree(os.path.join(root, d), ignore_errors=True)


# ─── Журнал ───────────────────────────────────────────────────────────

def append_journal(point_folder: str, title: str, lines: list[str],
                   now: datetime | None = None) -> str:
    """
    Дописать запись в журнал пункта. Формат — простой текст:

        ── 22.07.2026 14:05 · Применение ремонта ──
        источник: 222-1-84 (ИД сдана 13.04.2026)
        Импорт: 43 сечения, 257 ячеек
        ...

    Возвращает путь к журналу.
    """
    stamp = (now or datetime.now()).strftime("%d.%m.%Y %H:%M")
    path = os.path.join(point_folder, JOURNAL_NAME)
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"── {stamp} · {title} ──\n")
        for line in lines:
            f.write(f"{line}\n")
        f.write("\n")
    return path
