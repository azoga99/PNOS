# -*- coding: utf-8 -*-
"""
Поиск файлов пункта для модуля «Ремонт» (workers.repair_worker).

Проверка на временных папках: важна не начинка файлов, а выбор нужного из
нескольких похожих.
"""

from workers.repair_worker import find_point_xlsm, find_point_report


def _touch(folder, *names):
    for n in names:
        (folder / n).write_text("", encoding="utf-8")


# ─── Книга пункта ─────────────────────────────────────────────────────

def test_prefers_point_workbook_over_cleansource(tmp_path):
    """
    CleanSource.xlsx создаёт этап 5 из листа «Экспорт».

    Она шла первой по алфавиту (латиница раньше кириллицы), и модуль читал
    сечения из неё вместо книги пункта — сечений там нет вовсе.
    """
    _touch(tmp_path, "CleanSource.xlsx", "ПНОС_ТТП.xlsm")
    assert find_point_xlsm(str(tmp_path)).endswith("ПНОС_ТТП.xlsm")


def test_ignores_word_lock_files(tmp_path):
    _touch(tmp_path, "~$НОС_ТТП.xlsm", "ПНОС_ТТП.xlsm")
    assert find_point_xlsm(str(tmp_path)).endswith("ПНОС_ТТП.xlsm")


def test_xlsm_wins_over_xlsx(tmp_path):
    """Книга пункта — с макросами; прочее берём, только если её нет."""
    _touch(tmp_path, "AAA.xlsx", "ПНОС_ТТП.xlsm")
    assert find_point_xlsm(str(tmp_path)).endswith("ПНОС_ТТП.xlsm")


def test_falls_back_to_xlsx_when_no_xlsm(tmp_path):
    _touch(tmp_path, "книга.xlsx")
    assert find_point_xlsm(str(tmp_path)).endswith("книга.xlsx")


def test_none_when_no_workbook(tmp_path):
    _touch(tmp_path, "Отчет.docx")
    assert find_point_xlsm(str(tmp_path)) is None


# ─── Отчёт пункта ─────────────────────────────────────────────────────

def test_report_absent_before_stage5(tmp_path):
    """На ручном контроле отчёта ещё нет — только шаблон ПНОС_ТТП.docx."""
    _touch(tmp_path, "ПНОС_ТТП.docx", "ПНОС_ТТП.xlsm")
    assert find_point_report(str(tmp_path)) is None


def test_report_prefers_non_print_copy(tmp_path):
    _touch(tmp_path, "ЗЭПБ № ТТП-2692.26 - ттп № 108.docx",
           "ЗЭПБ № ТТП-2692.26 - ттп № 108 ПЕЧАТЬ.docx")
    assert "печать" not in find_point_report(str(tmp_path)).lower()


# ─── Шаблон отчёта для этапа 5 ────────────────────────────────────────

def test_template_found_when_present(tmp_path):
    from workers.stage4_worker import find_template_docx
    _touch(tmp_path, "ПНОС_ТТП.docx", "ПНОС_ТТП.xlsm")
    assert find_template_docx(str(tmp_path)).endswith("ПНОС_ТТП.docx")


def test_finished_report_is_not_a_template(tmp_path):
    """
    После первого прогона отчёт идёт по алфавиту раньше шаблона («ЗЭПБ…»
    перед «ПНОС_ТТП»). Прежний поиск «первый .docx» подставлял его: слияние
    шло поверх готового отчёта, макрос падал с «Не удалось установить связь
    с документом ПНОС_ТТП.docx», и появлялся мусорный второй отчёт.
    """
    from workers.stage4_worker import find_template_docx
    _touch(tmp_path, "ЗЭПБ № ТТП-2616.26 - ттп № 29.docx")
    assert find_template_docx(str(tmp_path)) is None


def test_template_wins_over_report(tmp_path):
    from workers.stage4_worker import find_template_docx
    _touch(tmp_path, "ЗЭПБ № ТТП-2616.26 - ттп № 29.docx", "ПНОС_ТТП.docx")
    assert find_template_docx(str(tmp_path)).endswith("ПНОС_ТТП.docx")
