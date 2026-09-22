# -*- coding: utf-8 -*-
"""
Проверка: не собран ли пункт по чужой строке графика.

Этап 4 пишет в «паспорт!F9» номер п/п (столбец A графика), найденный по
номеру папки (столбец B), и макрос Restore заполняет паспорт из строки с этим
A. До 2.2.4 ненайденный номер уходил в F9 как есть, а при повторе в B
выигрывала последняя строка листа любого вида оборудования. В обоих случаях
отчёт тихо собирался по другому пункту — с его установкой и актами ГИ.

Скрипт ничего не пишет: читает F9 из книги каждого пункта и сверяет с тем A,
который должен соответствовать номеру папки.

    python tools/check_plan_numbers.py <папка с пунктами> [ещё папки…] [--schedule <график.xlsm>]

По умолчанию график — CONFIG["MACRO_MASTER_WB_PATH"]. Код выхода 1, если
нашёлся хоть один пункт по чужой строке.
"""

import argparse
import glob
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl  # noqa: E402

from config import CONFIG  # noqa: E402
from excel_service import (  # noqa: E402
    PlanLookupError, build_plan_index, plan_key, read_point_f9, resolve_plan_number,
)


def read_schedule(path):
    """Индекс ТТП / ЭПБ по B и все строки по A (чтобы назвать чужую строку)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(wb["ПЛАН"].iter_rows(min_col=1, max_col=8, values_only=True))
    finally:
        wb.close()
    by_a = {}
    for n, row in enumerate(rows, start=1):
        cells = (tuple(row) + (None,) * 8)[:8]
        key = plan_key(cells[0])
        if key is not None:
            by_a.setdefault(key, (n, str(cells[3] or "").strip(), str(cells[4] or "").strip()))
    return build_plan_index(rows), by_a


def point_workbook(folder):
    for path in sorted(glob.glob(os.path.join(folder, "*.xlsm"))):
        if not os.path.basename(path).startswith("~$"):
            return path
    return None


def point_report(folder):
    for name in sorted(os.listdir(folder)):
        low = name.lower()
        if low.endswith(".docx") and not name.startswith("~$") and (
                name.startswith("ЗЭПБ") or name.startswith("Отчет_")):
            return name
    return None


def check_folder(folder, index, by_a):
    """(статус, текст). Статусы: ok / wrong / not_run / no_book / no_plan."""
    name = os.path.basename(folder)
    # Папки после сдачи переименовывают с пометками: «п.2684(отпр Шар)»
    m = re.match(r"п\.(\d+)", name)
    number = m.group(1) if m else name[2:]
    try:
        plan = resolve_plan_number(index, number)
    except PlanLookupError as e:
        return "no_plan", str(e)

    xlsm = point_workbook(folder)
    if not xlsm:
        return "no_book", "нет книги пункта"
    try:
        f9 = plan_key(read_point_f9(xlsm))
    except Exception as e:  # битая книга, нет листа
        return "no_book", f"книга не читается: {e}"
    if f9 is None:
        return "not_run", "F9 пуст — этап 4 не запускался"

    if f9 == plan["a"]:
        return "ok", f"F9 = {f9} ({plan['place']})"

    got = by_a.get(f9)
    got_text = f"строка {got[0]}, {got[1]}, {got[2]}" if got else "такого номера п/п в графике нет"
    report = point_report(folder)
    tail = f"; отчёт «{report}» собран по чужой строке" if report else ""
    return "wrong", (f"F9 = {f9} ({got_text}), должно быть {plan['a']} "
                     f"(строка {plan['row']}, {plan['place']}){tail}")


LABELS = {
    "wrong": "ЧУЖАЯ СТРОКА",
    "ok": "ок",
    "not_run": "не собран",
    "no_book": "нет книги",
    "no_plan": "нет в графике",
}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("roots", nargs="+", help="папки, в которых лежат папки п.NNNN")
    ap.add_argument("--schedule", default=CONFIG.get("MACRO_MASTER_WB_PATH", ""),
                    help="сводный график (.xlsm)")
    args = ap.parse_args(argv)

    if not os.path.isfile(args.schedule):
        print(f"График не найден: {args.schedule}")
        return 2
    print(f"График: {args.schedule}")
    index, by_a = read_schedule(args.schedule)

    counts = {k: 0 for k in LABELS}
    for root in args.roots:
        folders = sorted(p for p in glob.glob(os.path.join(root, "п.*")) if os.path.isdir(p))
        print(f"\n{root} — пунктов: {len(folders)}")
        for folder in folders:
            status, text = check_folder(folder, index, by_a)
            counts[status] += 1
            print(f"  {LABELS[status]:<13} {os.path.basename(folder):<22} {text}")

    print("\nИтого: " + ", ".join(f"{LABELS[k]} — {v}" for k, v in counts.items() if v))
    return 1 if counts["wrong"] else 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
