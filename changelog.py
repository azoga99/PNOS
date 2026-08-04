# -*- coding: utf-8 -*-
"""
Чтение журнала изменений (CHANGELOG.md) для окна «Что нового».

Единый источник правды — CHANGELOG.md в корне репозитория. Он же кладётся
в поставку (см. datas в PNOS.spec) и читается через resource_path, поэтому
журнал доступен офлайн и всегда соответствует установленной версии.

Формат секции (важен для парсера):

    ## X.Y.Z — ДД.ММ.ГГГГ
    **Новое**
    - пункт
    - пункт
    **Исправлено**
    - пункт

Заголовок версии — строго `## X.Y.Z`, дата после `—`/`-` необязательна.
Внутри — группы `**Новое**` / `**Улучшено**` / `**Исправлено**` / `**Изменено**`
со списком пунктов `- ...` (перенос строки внутри пункта допускается).

Модуль намеренно лёгкий (только re/os) — его импортируют settings_page и
main, где важен быстрый старт.
"""

from __future__ import annotations

import re

from utils import resource_path

# Заголовок версии: «## 2.1.0 — 04.08.2026» (дата необязательна; допускаем [ ]).
_VERSION_RE = re.compile(
    r"^##\s+\[?(\d+\.\d+\.\d+)\]?\s*(?:[—–-]\s*(.+?))?\s*$", re.MULTILINE)
# Заголовок группы: «**Новое**».
_GROUP_RE = re.compile(r"^\*\*(.+?)\*\*\s*$")

# Порядок и человекочитаемые метки категорий. Ключ — как в CHANGELOG.md
# (в нижнем регистре), значение — (метка чипа, код категории для стиля).
CATEGORIES = {
    "новое": ("Новое", "new"),
    "улучшено": ("Улучшено", "improved"),
    "исправлено": ("Исправлено", "fixed"),
    "изменено": ("Изменено", "changed"),
}
_CATEGORY_ORDER = ["new", "improved", "fixed", "changed"]


def _changelog_path() -> str:
    return resource_path("CHANGELOG.md")


def _read_text() -> str:
    try:
        with open(_changelog_path(), encoding="utf-8") as f:
            return f.read()
    except OSError:
        return ""


def _parse_version(ver: str) -> tuple:
    """'2.1.0' → (2, 1, 0). Нечисловые части → 0."""
    parts = []
    for p in ver.split("."):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def _parse_groups(body: str) -> list[dict]:
    """Тело секции версии → список групп [{"label","code","items":[...]}].

    Пункты `- ...`; строки-продолжения (без `-`) приклеиваются к последнему
    пункту. Строки вне групп (без предшествующего `**...**`) игнорируются.
    """
    groups: list[dict] = []
    cur: dict | None = None
    for raw in body.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        gm = _GROUP_RE.match(line.strip())
        if gm:
            label, code = CATEGORIES.get(
                gm.group(1).strip().lower(),
                (gm.group(1).strip(), "changed"))
            cur = {"label": label, "code": code, "items": []}
            groups.append(cur)
            continue
        stripped = line.strip()
        if stripped.startswith("- "):
            if cur is None:
                cur = {"label": "Изменено", "code": "changed", "items": []}
                groups.append(cur)
            cur["items"].append(stripped[2:].strip())
        elif cur and cur["items"]:
            # Перенос строки внутри пункта.
            cur["items"][-1] += " " + stripped
    return [g for g in groups if g["items"]]


def load_entries() -> list[dict]:
    """
    Разбирает CHANGELOG.md на секции версий.

    Возвращает список {"version","date","groups","text"} в порядке файла
    (свежие сверху). `text` — плоский текст версии для поиска. Пустой список,
    если файла нет.
    """
    text = _read_text()
    if not text:
        return []

    matches = list(_VERSION_RE.finditer(text))
    entries = []
    for i, m in enumerate(matches):
        version = m.group(1)
        date = (m.group(2) or "").strip()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        body = text[start:end]
        groups = _parse_groups(body)
        flat = " ".join(f"{version} {date} " + g["label"] + " " +
                        " ".join(g["items"]) for g in groups)
        entries.append({"version": version, "date": date,
                        "groups": groups, "text": flat.lower()})
    return entries


def entries_since(version: str) -> list[dict]:
    """Только версии новее указанной (для автопоказа после обновления)."""
    entries = load_entries()
    if not version:
        return entries
    base = _parse_version(version)
    return [e for e in entries if _parse_version(e["version"]) > base]
