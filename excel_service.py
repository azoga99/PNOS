# -*- coding: utf-8 -*-
"""
Модуль фильтрации Excel-файла для определения пунктов отчётов.
Фильтры: столбец E содержит "ТТП", столбец F содержит "ЭПБ", столбец H пустой.
Номера пунктов берутся из столбца B.

Также — чтение графика для страницы «Анализ пунктов»: фильтрация по
закреплённому специалисту (CONFIG["ANALYTICS"]) и вычисление статуса
согласования по колонкам-датам QZ/RB/RC/RE.
"""

import re
import datetime

import pandas as pd
from config import CONFIG


def analyze_excel(file_path: str, batch_size: int = None) -> tuple[int, list[str]]:
    """
    Фильтрует Excel и возвращает список номеров пунктов из столбца B.

    Фильтры:
        - Столбец E (индекс 4): содержит "ТТП"
        - Столбец F (индекс 5): содержит "ЭПБ"
        - Столбец H (индекс 7): пустой

    Args:
        file_path: путь к Excel файлу (.xlsx / .xlsm)
        batch_size: максимальное количество пунктов (по умолчанию из CONFIG)

    Returns:
        (total_filtered, points_list) — общее кол-во подходящих строк
        и список первых batch_size номеров пунктов
    """
    if batch_size is None:
        batch_size = CONFIG.get("BATCH_SIZE", 20)

    try:
        df = pd.read_excel(file_path, engine="openpyxl")

        # Столбцы по индексу: E=4, F=5, H=7
        col_e = df.iloc[:, 4].astype(str)
        col_f = df.iloc[:, 5].astype(str)
        col_h = df.iloc[:, 7]

        # Условия фильтрации
        cond_e = col_e.str.contains("ТТП", na=False, case=False)
        cond_f = col_f.str.contains("ЭПБ", na=False, case=False)
        cond_h = col_h.isna() | (col_h.astype(str).str.strip() == "")

        filtered = df[cond_e & cond_f & cond_h]
        total_count = len(filtered)

        # Извлекаем номера пунктов из столбца B (индекс 1)
        points = []
        for _, row in filtered.iterrows():
            raw_val = str(row.iloc[1]).strip()
            # Номера пунктов — целые числа, убираем .0 если есть
            if raw_val.endswith(".0"):
                raw_val = raw_val[:-2]
            if raw_val and raw_val != "nan":
                points.append(raw_val)

        # Ограничиваем размером пакета
        batch_points = points[:batch_size]

        return total_count, batch_points

    except Exception as e:
        print(f"Ошибка чтения Excel: {e}")
        return 0, []


# ═══════════════════════════════════════════════════════════════════════
# Номер пункта → строка графика (этап 4)
# ═══════════════════════════════════════════════════════════════════════
#
# Папка пункта названа по столбцу B («№ п/п по графикам»), а макрос Restore
# ищет строку графика по столбцу A («Номер п/п»). Числа в A и B у одной строки
# почти всегда разные, и одно и то же число в A и в B принадлежит разным
# пунктам: B=3224 — АВТ-5, A=3224 — 517 ХВП. Ошибка перевода не падает, а
# тихо собирает отчёт по чужому пункту, поэтому любая неоднозначность — отказ.


class PlanLookupError(Exception):
    """Пункт нельзя однозначно найти в графике — собирать по нему отчёт нельзя."""


def plan_key(value):
    """Номер из ячейки графика: 3224, 3224.0 и «3224» дают один ключ."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        num = float(text.replace(",", "."))
    except ValueError:
        return text
    return int(num) if num.is_integer() else text


def _filled(value) -> bool:
    return value is not None and str(value).strip() != ""


def build_plan_index(rows, first_row: int = 1) -> dict:
    """Строки ТТП + ЭПБ листа «ПЛАН», сгруппированные по столбцу B.

    rows — значения столбцов A…H построчно (``Range.Value`` из COM или
    ``iter_rows(values_only=True)``). Нумерация в B своя у каждого вида
    оборудования: одно число есть и у печи, и у РВС, и у ТТП. Поэтому берутся
    только строки с тем же фильтром, по которому этап 1 выбирает пункты:
    E содержит «ТТП», F содержит «ЭПБ».
    """
    index = {}
    for row_no, row in enumerate(rows, start=first_row):
        cells = (tuple(row) + (None,) * 8)[:8]
        a, b, _, place, kind, diag, deleted, mark = cells
        key = plan_key(b)
        if key is None:
            continue
        if "ТТП" not in str(kind or "").upper() or "ЭПБ" not in str(diag or "").upper():
            continue
        index.setdefault(key, []).append({
            "row": row_no,
            "a": plan_key(a),
            "place": str(place or "").strip(),
            "deleted": _filled(deleted),
            "marked": _filled(mark),
        })
    return index


def resolve_plan_number(index: dict, point) -> dict:
    """Строка графика для пункта из папки «п.<B>».

    Возвращает словарь ``{"row", "a", "place", ...}``; ``a`` — номер п/п,
    который пишется в ``паспорт!F9``. Нет строки, строк несколько или пуст
    столбец A — ``PlanLookupError``.
    """
    found = index.get(plan_key(point), [])
    if len(found) > 1:
        # Повторы в B у ТТП — строки, уже помеченные в G (удалён) или H;
        # этап 1 такие не выбирает. Живая строка среди них — наш пункт.
        live = [r for r in found if not r["deleted"] and not r["marked"]]
        if live:
            found = live
    if not found:
        raise PlanLookupError(
            f"п.{point} нет в столбце B графика среди строк ТТП / ЭПБ")
    if len(found) > 1:
        rows = ", ".join(str(r["row"]) for r in found)
        raise PlanLookupError(
            f"п.{point} встречается в столбце B графика несколько раз "
            f"(строки {rows})")
    match = found[0]
    if match["a"] is None:
        raise PlanLookupError(
            f"у п.{point} пуст столбец A графика (строка {match['row']})")
    return match


def read_point_f9(xlsm_path: str):
    """``паспорт!F9`` книги пункта без Excel — номер п/п, по которому Restore
    взял строку графика. Нужен, чтобы найти пункты, собранные по чужой строке."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsm_path, read_only=True, data_only=True)
    try:
        return wb["паспорт"]["F9"].value
    finally:
        wb.close()


# ═══════════════════════════════════════════════════════════════════════
# Анализ пунктов (страница «Анализ пунктов»)
# ═══════════════════════════════════════════════════════════════════════

# Лестница статусов согласования: код -> подпись в UI
ANALYTICS_STATUS_LABELS = {
    "in_work": "Не оформлен",
    "done": "Оформлен, не отправлен",
    "sent_approval": "Отправлен Шарову",
    "approved": "Согласован Шаровым",
    "sent_final": "Отправлен Жиганьшину",
}

_DATE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")


def col_to_idx(letter: str) -> int:
    """Буква колонки Excel -> 0-базный индекс ("B" -> 1, "MF" -> 343)."""
    idx = 0
    for ch in letter.strip().upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _cell_is_date(value) -> bool:
    """True, если в ячейке дата (datetime или строка строго ДД.ММ.ГГГГ)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    if isinstance(value, (datetime.datetime, datetime.date, pd.Timestamp)):
        return True
    return bool(_DATE_RE.match(str(value).strip()))


def _clean_point(value) -> str:
    """Номер пункта из ячейки: "2744.0" -> "2744"; пусто/nan -> ""."""
    s = str(value).strip()
    if s.lower() in ("", "nan", "none", "nat"):
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    return s


def read_analytics_rows(file_path: str) -> list[dict]:
    """
    Читает график целиком для страницы «Анализ пунктов».

    Возвращает список словарей по каждому пункту (все специалисты сразу —
    фильтрация по фамилии/оборудованию делается в UI без повторного чтения):
        {"point", "equipment", "person", "status", "status_label"}

    Всегда исключаются строки с пустым номером пункта и строки, где
    колонка COL_DELETED непуста («удалённые» пункты).
    """
    cfg = CONFIG["ANALYTICS"]
    letters = {
        "point": cfg["COL_POINT"],
        "equipment": cfg["COL_EQUIPMENT"],
        "deleted": cfg["COL_DELETED"],
        "person": cfg["COL_PERSON"],
        "done": cfg["COL_DONE"],
        "sent_approval": cfg["COL_SENT_APPROVAL"],
        "approved": cfg["COL_APPROVED"],
        "sent_final": cfg["COL_SENT_FINAL"],
    }

    df = pd.read_excel(
        file_path,
        header=None,
        engine="openpyxl",
        skiprows=cfg["DATA_START_ROW"] - 1,
        usecols=",".join(letters.values()),
        dtype=object,
    )
    # usecols возвращает колонки в порядке следования в файле —
    # сопоставляем их нашим именам через сортировку по индексу колонки
    ordered_names = [name for name, _ in
                     sorted(letters.items(), key=lambda kv: col_to_idx(kv[1]))]
    df.columns = ordered_names

    rows = []
    for _, r in df.iterrows():
        point = _clean_point(r["point"])
        if not point:
            continue

        deleted = str(r["deleted"]).strip() if r["deleted"] is not None else ""
        if deleted and deleted.lower() not in ("nan", "none", "nat"):
            continue  # «удалённый» пункт — не показываем никогда

        has_done = _cell_is_date(r["done"])
        has_sent = _cell_is_date(r["sent_approval"])
        has_approved = _cell_is_date(r["approved"])
        has_final = _cell_is_date(r["sent_final"])
        if has_done and has_sent and has_approved and has_final:
            status = "sent_final"
        elif has_done and has_sent and has_approved:
            status = "approved"
        elif has_done and has_sent:
            status = "sent_approval"
        elif has_done:
            status = "done"
        else:
            status = "in_work"

        equipment = str(r["equipment"]).strip()
        if equipment.lower() in ("nan", "none", "nat"):
            equipment = ""
        person = str(r["person"]).strip()
        if person.lower() in ("nan", "none", "nat"):
            person = ""

        rows.append({
            "point": point,
            "equipment": equipment,
            "person": person,
            "status": status,
            "status_label": ANALYTICS_STATUS_LABELS[status],
        })
    return rows
