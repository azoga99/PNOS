# -*- coding: utf-8 -*-
"""
Чтение данных паспорта трубопровода без Word и без COM.

DocVariables (`F01`, `F051`, …), которые этап 4 обычно вычитывает через Word
COM, лежат внутри .docx как `<w:docVar w:name="F051" w:val="01.05.1975"/>` в
`word/settings.xml` — их достаёт обычный zip за миллисекунды. Это нужно шагу 3:
форма ручного контроля показывает год ввода в эксплуатацию ещё до того, как
этап 4 запишет его в книгу (до этапа 4 ячейка `паспорт!B19` пуста).

Модуль намеренно лёгкий (os / re / zipfile): его импортирует `main_window`, где
тяжёлые импорты на верхнем уровне ломают быстрый старт.
"""

import os
import re
import zipfile
import difflib

# «01.05.1975» — формат, в котором год лежит и в F051, и в паспорт!B19
_DATE_RE = re.compile(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4})\b")
_YEAR_RE = re.compile(r"\b(\d{4})\b")
_DOCVAR_RE = re.compile(r'<w:docVar w:name="([^"]+)" w:val="([^"]*)"')

COMMISSION_VAR = "F051"      # DocVariable «Год ввода в эксплуатацию»


def find_passport_docx(folder: str, target: str = "паспорт",
                       cutoff: float = 0.7) -> str | None:
    """Путь к файлу паспорта пункта (папка ищется нечётко, как на этапе 4)."""
    if not folder or not os.path.isdir(folder):
        return None
    dirs = {d.lower(): d for d in os.listdir(folder)
            if os.path.isdir(os.path.join(folder, d))}
    name = dirs.get(target)
    if name is None:
        m = difflib.get_close_matches(target, list(dirs), n=1, cutoff=cutoff)
        name = dirs[m[0]] if m else None
    if name is None:
        return None
    p_folder = os.path.join(folder, name)
    for f in sorted(os.listdir(p_folder)):
        if f.startswith("~$"):
            continue
        if f.lower().endswith((".docx", ".docm")):
            return os.path.join(p_folder, f)
    return None


def read_doc_vars(path: str) -> dict:
    """DocVariables документа. Пустой словарь, если это не .docx или файл битый."""
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/settings.xml").decode("utf-8", "replace")
    except Exception:
        return {}
    return dict(_DOCVAR_RE.findall(xml))


def commission_date(folder: str) -> str:
    """Дата ввода в эксплуатацию из паспорта пункта («01.05.1975») или ''."""
    doc = find_passport_docx(folder)
    if not doc:
        return ""
    return (read_doc_vars(doc).get(COMMISSION_VAR) or "").strip()


def year_of(value) -> str:
    """Год (4 цифры) из даты любого вида: datetime, «01.05.1975», «1975», число."""
    if value is None:
        return ""
    if hasattr(value, "year"):
        return str(value.year)
    text = str(value).strip()
    if not text:
        return ""
    m = _DATE_RE.search(text)
    if m:
        return m.group(3)
    m = _YEAR_RE.search(text)
    return m.group(1) if m else ""


def day_month_of(value) -> tuple[str, str] | None:
    """День и месяц из даты («01», «05») или None, если их не разобрать."""
    if value is None:
        return None
    if hasattr(value, "day") and hasattr(value, "month"):
        return f"{value.day:02d}", f"{value.month:02d}"
    m = _DATE_RE.search(str(value))
    if not m:
        return None
    return f"{int(m.group(1)):02d}", f"{int(m.group(2)):02d}"


def commission_value(year: str, *sources) -> str:
    """
    Значение для `паспорт!B19` по введённому году.

    День и месяц берутся из первой разобранной даты в `sources` (текущее
    значение ячейки, затем F051 паспорта) — правка года не должна терять
    остальную часть даты. Если разобрать нечего, ставится 01.01.
    Пустой год — пустая строка (вызывающий код такое значение не пишет).
    """
    year = (str(year or "")).strip()
    if not re.fullmatch(r"\d{4}", year):
        return ""
    for src in sources:
        dm = day_month_of(src)
        if dm:
            return f"{dm[0]}.{dm[1]}.{year}"
    return f"01.01.{year}"
