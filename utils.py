# -*- coding: utf-8 -*-
import os
import sys

def resource_path(relative_path: str) -> str:
    """Получает абсолютный путь к ресурсу, работает в разработке и в EXE."""
    try:
        # PyInstaller создает временную папку и сохраняет путь в _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# Родительный падеж месяцев для дат вида «06 июля 2026 г.»
MONTHS_GENITIVE = {
    "01": "января", "02": "февраля", "03": "марта", "04": "апреля",
    "05": "мая", "06": "июня", "07": "июля", "08": "августа",
    "09": "сентября", "10": "октября", "11": "ноября", "12": "декабря",
}


def title_date_text(date_str: str) -> str:
    """«06.07.2026» → «06 июля 2026 г.» (формат даты на титуле отчёта)."""
    dd, mm, yyyy = date_str.strip().split(".")
    return f"{dd} {MONTHS_GENITIVE[mm]} {yyyy} г."


def find_order_scan(folder: str) -> str | None:
    """Ищет скан приказа (JPG/PNG) в корне папки «Правки отчётов».
    Если файлов несколько — берёт самый свежий по дате изменения."""
    candidates = []
    try:
        for f in os.listdir(folder):
            p = os.path.join(folder, f)
            if os.path.isfile(p) and f.lower().endswith((".jpg", ".jpeg", ".png")):
                candidates.append(p)
    except OSError:
        return None
    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)
