# -*- coding: utf-8 -*-
"""Тесты записи строки «Сведения о ремонте» в .docx."""

import docx
import pytest

from repair.word_writer import read_svedenia, write_svedenia, LABEL


def _make_report(path, svedenia_text="старый текст", with_label=True):
    d = docx.Document()
    d.add_paragraph("5 СВЕДЕНИЯ О РАССМОТРЕННЫХ ДОКУМЕНТАХ")
    t = d.add_table(rows=3, cols=3)
    t.rows[0].cells[0].text = "Паспорт трубопровода"
    t.rows[0].cells[1].text = "Технический паспорт №6"
    t.rows[0].cells[2].text = "73"
    if with_label:
        t.rows[1].cells[0].text = LABEL
        # текст с форматированием: два run-а, первый — курсив
        cell = t.rows[1].cells[1]
        p = cell.paragraphs[0]
        r1 = p.add_run(svedenia_text[:10])
        r1.italic = True
        p.add_run(svedenia_text[10:])
        t.rows[1].cells[2].text = "42"
    t.rows[2].cells[0].text = "Заключение ЭПБ"
    t.rows[2].cells[1].text = "№0340-ТУ-2017"
    d.add_paragraph("6 КРАТКАЯ ХАРАКТЕРИСТИКА")
    d.save(path)


def test_read(tmp_path):
    p = str(tmp_path / "отчёт.docx")
    _make_report(p, "Исполнительная документация в наличии")
    assert read_svedenia(p) == "Исполнительная документация в наличии"


def test_write_and_preserve_format(tmp_path):
    p = str(tmp_path / "отчёт.docx")
    _make_report(p)
    new = "Исполнительная документация в наличии: 14.04.2026 г. проведена замена"
    assert write_svedenia(p, new)

    assert read_svedenia(p) == new
    # форматирование первого run сохранено (курсив), run теперь один
    d = docx.Document(p)
    for t in d.tables:
        for row in t.rows:
            if row.cells[0].text.strip() == LABEL:
                para = row.cells[1].paragraphs[0]
                assert len(para.runs) == 1
                assert para.runs[0].italic is True
                return
    pytest.fail("строка с меткой не найдена")


def test_write_no_label(tmp_path):
    p = str(tmp_path / "отчёт.docx")
    _make_report(p, with_label=False)
    assert write_svedenia(p, "текст") is False
    # файл не изменён по смыслу: метки нет
    assert read_svedenia(p) is None


def test_other_cells_untouched(tmp_path):
    p = str(tmp_path / "отчёт.docx")
    _make_report(p)
    write_svedenia(p, "новый текст")
    d = docx.Document(p)
    t = d.tables[0]
    assert t.rows[0].cells[1].text == "Технический паспорт №6"
    assert t.rows[1].cells[2].text == "42"     # кол-во листов не трогаем
    assert t.rows[2].cells[1].text == "№0340-ТУ-2017"


def test_multiparagraph_cell_collapsed(tmp_path):
    p = str(tmp_path / "отчёт.docx")
    _make_report(p)
    # добавить второй абзац в ячейку текста
    d = docx.Document(p)
    for t in d.tables:
        for row in t.rows:
            if row.cells[0].text.strip() == LABEL:
                row.cells[1].add_paragraph("хвост")
    d.save(p)

    write_svedenia(p, "чисто")
    assert read_svedenia(p) == "чисто"
