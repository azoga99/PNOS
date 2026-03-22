# -*- coding: utf-8 -*-
"""
Модуль фильтрации Excel-файла для определения пунктов отчётов.
Фильтры: столбец E содержит "ТТП", столбец F содержит "ЭПБ", столбец H пустой.
Номера пунктов берутся из столбца B.
"""

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
