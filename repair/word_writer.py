# -*- coding: utf-8 -*-
"""
Запись строки «Сведения о ремонте» в отчёт ЗЭПБ (.docx).

В таблице 1 отчёта («Перечень документов») есть строка с ячейкой-меткой
«Сведения о ремонте»; текст лежит в следующей ячейке той же строки. Модуль
находит эту ячейку независимо от номера таблицы/строки (структура шаблона
может меняться) и заменяет текст, сохраняя форматирование первого run
(шрифт/кегль ячейки не сбиваются).

Пишем через python-docx: правится ровно одна ячейка, остальной документ
(картинки, поля, стили) не пересобирается. Бэкап файла делает оркестратор.
"""

from __future__ import annotations

LABEL = "Сведения о ремонте"


def _iter_tables(tables):
    """Таблицы документа, включая вложенные."""
    for t in tables:
        yield t
        for row in t.rows:
            for cell in row.cells:
                if cell.tables:
                    yield from _iter_tables(cell.tables)


def _find_cell(doc):
    """
    Найти ячейку с текстом строки «Сведения о ремонте».
    Возвращает docx-ячейку с ТЕКСТОМ (следующую за меткой) или None.
    """
    for table in _iter_tables(doc.tables):
        for row in table.rows:
            cells = row.cells
            for i, cell in enumerate(cells):
                if cell.text.strip() != LABEL:
                    continue
                # следующая ячейка строки с отличным от метки содержимым
                # (при горизонтальном объединении row.cells дублируется)
                for nxt in cells[i + 1:]:
                    if nxt._tc is not cell._tc:
                        return nxt
    return None


def set_cell_text(cell, text: str) -> None:
    """
    Заменить текст ячейки, сохранив форматирование первого run.
    Лишние абзацы и run-ы удаляются.
    """
    first_para = cell.paragraphs[0]

    if first_para.runs:
        keep = first_para.runs[0]
        keep.text = text
        for run in first_para.runs[1:]:
            run._element.getparent().remove(run._element)
    else:
        first_para.add_run(text)

    for para in cell.paragraphs[1:]:
        para._element.getparent().remove(para._element)


def read_svedenia(docx_path: str) -> str | None:
    """Текущий текст строки «Сведения о ремонте» (None — метка не найдена)."""
    import docx  # локальный импорт — не тянуть при старте приложения

    cell = _find_cell(docx.Document(docx_path))
    return None if cell is None else cell.text.strip()


def write_svedenia(docx_path: str, new_text: str) -> bool:
    """
    Записать текст в строку «Сведения о ремонте» и сохранить документ.
    False — метка в документе не найдена (файл не изменён).
    """
    import docx

    doc = docx.Document(docx_path)
    cell = _find_cell(doc)
    if cell is None:
        return False
    set_cell_text(cell, new_text)
    doc.save(docx_path)
    return True
