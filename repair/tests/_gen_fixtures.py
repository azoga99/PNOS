# -*- coding: utf-8 -*-
"""
Одноразовый генератор фикстур для регрессионных тестов модуля «Ремонт».

Читает извлечённые тексты эталонных отчётов (до/после ремонта) из каталога
--src, парсит таблицы толщинометрии и строку «Сведения о ремонте», добавляет
выверенные вручную списки заменённых сечений и пишет компактный JSON в
fixtures/. После этого тесты самодостаточны и не зависят от внешних путей.

Запуск (интерпретатор ПНОС):
    python -m repair.tests._gen_fixtures --src <каталог с *.txt>
"""

import argparse
import json
import os

# Заменённые сечения и переходы — выверены по схемам актов отбраковки,
# совпадение с эталоном 100% (ретро-тесты сессии проектирования модуля).
REPLACED_2701 = {
    1: 12.0, 2: 12.0, 3: 10.0, 5: 12.0, 6: 10.0, 7: 8.0, 8: 10.0, 9: 10.0,
    10: 10.0, 11: 12.0, 12: 10.0, 13: 8.0, 14: 10.0, 15: 10.0, 17: 10.0,
    18: 12.0, 19: 5.0, 22: 5.0, 23: 10.0, 24: 10.0, 27: 10.0, 28: 10.0,
    29: 10.0, 30: 10.0, 32: 8.0, 33: 10.0, 34: 8.0, 35: 10.0, 36: 10.0,
    37: 12.0, 38: 5.0, 39: 5.0, 57: 12.0, 94: 10.0, 95: 10.0, 96: 10.0,
    100: 8.0, 101: 10.0, 102: 8.0, 103: 8.0, 104: 8.0, 105: 10.0, 108: 12.0,
}
PEREHOD_2701 = {11: ("12 / 10", "273 / 219")}

REPLACED_2775 = {47: 6.0, 64: 6.0, 65: 6.0, 68: 6.0, 106: 6.0}
PEREHOD_2775 = {}


def parse_table(path):
    """Таблица толщинометрии: строки вида «№ | Сечение | Элемент | Тисп | Диам | т1..т4 | Мин»."""
    rows = []
    for line in open(path, encoding="utf-8"):
        parts = [p.strip() for p in line.rstrip("\n").split(" | ")]
        if len(parts) != 10:
            continue
        if not (parts[0].isdigit() and parts[1].isdigit()):
            continue
        rows.append({
            "number": int(parts[1]),
            "element": parts[2],
            "thickness_nom": parts[3],
            "diameter": parts[4],
            "measures": [parts[5], parts[6], parts[7], parts[8]],
            "tmin": parts[9],
        })
    return rows


def find_svedenia(path):
    for line in open(path, encoding="utf-8"):
        if "Сведения о ремонте" in line:
            # ячейка таблицы: «... | Сведения о ремонте | <текст> | <стр> »
            cells = [c.strip() for c in line.split(" | ")]
            for i, c in enumerate(cells):
                if c == "Сведения о ремонте" and i + 1 < len(cells):
                    return cells[i + 1]
    return ""


def build(src):
    fx = {
        "2701": {
            "old": os.path.join(src, "zepb.txt"),
            "new": os.path.join(src, "zepb_new.txt"),
            "replaced": REPLACED_2701,
            "perehod": PEREHOD_2701,
        },
        "2775": {
            "old": os.path.join(src, "z2775_old.txt"),
            "new": os.path.join(src, "z2775_new.txt"),
            "replaced": REPLACED_2775,
            "perehod": PEREHOD_2775,
        },
    }
    out_dir = os.path.join(os.path.dirname(__file__), "fixtures")
    os.makedirs(out_dir, exist_ok=True)

    for key, spec in fx.items():
        data = {
            "point": key,
            "sections_old": parse_table(spec["old"]),
            "sections_new": parse_table(spec["new"]),
            "replaced": {str(k): v for k, v in spec["replaced"].items()},
            "perehod": {str(k): list(v) for k, v in spec["perehod"].items()},
            "svedenia_new": find_svedenia(spec["new"]),
        }
        out = os.path.join(out_dir, f"{key}.json")
        with open(out, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        print(f"{out}: old={len(data['sections_old'])} new={len(data['sections_new'])} "
              f"replaced={len(data['replaced'])}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="каталог с извлечёнными *.txt")
    build(ap.parse_args().src)
