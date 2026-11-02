# -*- coding: utf-8 -*-
"""
Модель «Рабочего места»: строки таблицы пунктов по папкам ПНОС.

Чистый модуль без Qt: читает папки `п.*` в локальной папке и их скрытые
`.pnos_status.json` и превращает в строки для `WorkspacePage.set_points`:

    {"point": "3186", "folder": путь,
     "steps": {1..7: "done" | "run" | "wait" | "error" | "idle"},
     "state": текст, "state_kind": "info" | "warn" | "error" | "done"}

Состояние «run» модель не ставит — о том, что шаг идёт прямо сейчас, знает
только окно, которое его запустило.
"""

import json
import os
import re

from status_manager import GI_WAIT_KEY, STATUS_FILE_NAME, is_point_blocked

# UI-шаг -> ключ в .pnos_status.json (как MainWindow.STEP_STATUS_KEYS:
# шаг 3 ручной, поэтому воркеры этапов 4–7 пишут stage3–stage6)
STEP_KEYS = {1: "stage1", 2: "stage2", 3: "manual3",
             4: "stage3", 5: "stage4", 6: "stage5", 7: "stage6"}

# Короткие названия шагов для строки шагов и подсказок
STEP_NAMES = {1: "Структура", 2: "Таблицы", 3: "Форма", 4: "Макросы",
              5: "Отчёт", 6: "Фото", 7: "Акты ГИ"}

FORM_STEP = 3
GI_STEP = 7


def point_number(name: str) -> str:
    """Номер пункта из имени папки: «п.3186 (не корректный)» → «3186»."""
    m = re.search(r"\d+", name or "")
    return m.group(0) if m else ""


def is_point_folder(name: str) -> bool:
    """Папка пункта — «п.N», «п N» или «пункт N» (как на странице «Обзор»)."""
    low = (name or "").lower()
    return (low.startswith(("п.", "п ", "пункт"))
            and bool(point_number(name)))


def _read_status(folder: str) -> dict:
    try:
        with open(os.path.join(folder, STATUS_FILE_NAME), "r",
                  encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def next_step(steps: dict) -> int | None:
    """Первый невыполненный шаг; None — выполнены все семь."""
    for step in range(1, 8):
        if steps.get(step) != "done":
            return step
    return None


def _blocked_text(data: dict) -> tuple[str, str]:
    """Причина блокировки для колонки «Состояние»: (текст, категория)."""
    if data.get("skipped") and not data.get("invalid"):
        return "Пропущен", "warn"
    missing = [str(m) for m in (data.get("missing_folders") or [])]
    if missing:
        return f"Не корректный: нет {', '.join(missing)}", "error"
    return "Не корректный", "error"


def build_row(folder: str) -> dict:
    """Строка таблицы для одной папки пункта."""
    data = _read_status(folder)
    done = {step: bool(data.get(key)) for step, key in STEP_KEYS.items()}
    steps = {step: ("done" if ok else "idle") for step, ok in done.items()}
    nxt = next_step(steps)

    blocked, _ = is_point_blocked(folder)
    if nxt is None:
        state, kind = "Готов", "done"
    elif blocked:
        steps[nxt] = "error"
        state, kind = _blocked_text(data)
    elif done[2] and not done[FORM_STEP]:
        steps[FORM_STEP] = "wait"
        state, kind = "Ждёт формы", "warn"
    elif nxt == GI_STEP and data.get(GI_WAIT_KEY):
        # Шаг 7 прошёл без актов: не «не начат», а ждёт, пока акты появятся
        steps[GI_STEP] = "wait"
        state, kind = f"{data[GI_WAIT_KEY]} — повторите шаг 7, когда появятся", "warn"
    else:
        state, kind = f"Следующий шаг: {nxt} · {STEP_NAMES[nxt]}", "info"

    return {
        "point": point_number(os.path.basename(os.path.normpath(folder))),
        "folder": folder,
        "steps": steps,
        "state": state,
        "state_kind": kind,
    }


def build_rows(local_root: str) -> list[dict]:
    """
    Строки по всем папкам пунктов в local_root, по возрастанию номера.

    Если рядом лежат «п.N» и «п.N (не корректный)» (этап 1 не удаляет вторую,
    чтобы освободить имя), в таблицу идёт «п.N»: этапы 2–7 работают с ней.
    """
    if not local_root or not os.path.isdir(local_root):
        return []
    by_number: dict[str, str] = {}
    for name in sorted(os.listdir(local_root)):
        folder = os.path.join(local_root, name)
        if not os.path.isdir(folder) or not is_point_folder(name):
            continue
        num = point_number(name)
        prev = by_number.get(num)
        if prev is None or "не корректн" in os.path.basename(prev).lower():
            by_number[num] = folder
    return [build_row(by_number[n]) for n in sorted(by_number, key=int)]
