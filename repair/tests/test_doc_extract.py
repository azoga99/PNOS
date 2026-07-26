# -*- coding: utf-8 -*-
"""Тесты извлечения полей из OCR-текста (Э4 слой 2, чистая логика)."""

from repair.doc_extract import extract_doc_info, normalize_numero


def test_act_number_and_word_date():
    lines = [
        "УТВЕРЖДАЮ:",
        "Заместитель главного инженера",
        "«24» февраля 2026 г.",
        "Акт ревизии и отбраковки трубопровода № 222-1-84",
        "по установке 36-30 «19» февраля 2026 г. произведена ревизия",
    ]
    info = extract_doc_info(lines)
    assert info.act_numbers == ["222-1-84"]
    assert info.act_date == "24.02.2026"
    assert info.unit == "36-30"


def test_handwritten_day_missing():
    """День вписан от руки и не распознался — остаётся месяц.год."""
    lines = ["УТВЕРЖДАЮ:", "«  » марта 2026 г.",
             "Акт ревизии и отбраковки трубопровода № 222-1-121"]
    info = extract_doc_info(lines)
    assert info.act_date == ".03.2026"     # день заполнит сотрудник


def test_numeric_date_near_act():
    lines = ["Акт ревизии и отбраковки трубопровода № 222-1-309 от 07.07.2023г."]
    info = extract_doc_info(lines)
    assert info.act_numbers == ["222-1-309"]
    assert info.act_date == "07.07.2023"


def test_line_and_unit():
    lines = ["Трубопровод №6 «Линия жидкого пропана от Е-1 1а, 1б к Н-2,3»",
             "по установке 36-30 произведена ревизия"]
    info = extract_doc_info(lines)
    assert info.line == "6"
    assert info.unit == "36-30"


def test_replacement_table_rows():
    lines = [
        "Наименование элементов трубопровода, подлежащих замене",
        "1 Труба 273х10 26 п. м. Бесшовная Ст. 20",
        "2 Труба 219х8 4 п. м. Бесшовная Ст. 20",
        "3 Отвод 90°-273х10 5 шт. Крутоизогнутый Ст. 20",
        "5 Переход 273х12-219х10 1 шт. Концентрический Ст. 20",
        "2 Бобышка Дн45 1 шт. Скошенная Ст. 20",
        "Сварка, термическая обработка, неразрушающий контроль",
    ]
    info = extract_doc_info(lines)
    assert len(info.replacements) == 5
    assert any("273х10" in r for r in info.replacements)
    assert any("Бобышка" in r for r in info.replacements)
    assert info.material.replace(" ", "").lower().startswith("ст")


def test_stainless_material():
    lines = ["1 Труба 89х5 6,0 п. м. Бесшовная 12Х18Н10Т"]
    info = extract_doc_info(lines)
    assert info.material == "12Х18Н10Т"


# ─── Артефакты OCR из обучающей выборки («для обучени») ────────────────

def test_numero_variants_normalized():
    """«№» распознаётся десятком способов: N2, Х!, У2, М•, даже «*2»."""
    assert normalize_numero("линия N2108") == "линия №108"
    assert normalize_numero("линия *2119") == "линия №119"
    assert normalize_numero("Трубопровод У2б") == "Трубопровод №б"


def test_numero_rule_keeps_sizes_and_grades():
    """Правило не должно портить типоразмеры и марки стали."""
    assert normalize_numero("Труба 273х10") == "Труба 273х10"
    assert normalize_numero("Бесшовная 12Х18Н10Т") == "Бесшовная 12Х18Н10Т"


def test_unit_written_with_dot():
    """В ИД установку пишут «ПКМ уст.39-40 тит. 39-40» — точка вместо пробела."""
    info = extract_doc_info(["ПКМ уст.39-40 тит. 39-40"])
    assert info.unit == "39-40"


def test_quality_certificate_material():
    """В удостоверении о качестве марку пишут словом: «Сталь 20»."""
    info = extract_doc_info([
        "1 Труба 108х6 Сертификат №6087 от 11.09.2022 Замена Сталь 20 ГОСТ 8732-78"])
    assert info.material.replace(" ", "").lower() == "сталь20"


def test_act_number_with_ocr_noise():
    """Копия акта внутри ИД сканирована хуже: «222-l.179» вместо «222-1-179»."""
    info = extract_doc_info(["Акт ревизии и отбраковки трубопр ода №222-l.179"])
    assert info.act_numbers == ["222-1-179"]


def test_act_number_glued_to_numero_sign():
    """«Акт ОТН Y2222-1-179»: знак номера слипся с номером, лишняя цифра."""
    info = extract_doc_info(["Акт ОТН Y2222-1-179"])
    assert info.act_numbers == ["222-1-179"]


def test_act_number_is_not_line_number():
    """«трубопровода № 222-1-149» — номер акта, а не линия «222» или «22»."""
    info = extract_doc_info([
        "Акт ревизии и отбраковки трубопровода № 222-1-149",
        "линия №119",
    ])
    assert info.line == "119"
    assert "222-1-149" in info.act_numbers


def test_line_with_letter_suffix():
    info = extract_doc_info(["Трубопровод №39-МЦК «Линия факельная»"])
    assert info.line == "39-МЦК"


def test_defect_description_is_not_a_replacement():
    """
    Ловушка из акта п.2763: в описании дефекта есть и название элемента, и
    типоразмер — «минимальная толщина для перехода Дн108х89 составляет 2,6мм».
    Такая строка попадала в колонку «Замена» вместо «Переход 108х6-89х6».
    """
    from repair.doc_extract import is_replacement_line

    assert not is_replacement_line(
        "(Коррозионный износ, минимальная толщина для перехода Дн108х89 "
        "составляет 2,6мм, отбраковочная толщина – 2,0мм)")
    assert not is_replacement_line("для перехода Дн108х89 составляет 2,6мм, от-")
    assert is_replacement_line("Переход 108х6-89х6")
    assert is_replacement_line("1 Труба 273х10 26 п. м. Бесшовная Ст. 20")


def test_replacement_line_survives_ocr_noise():
    """OCR глотает буквы: «Пе еход» — всё ещё позиция таблицы замены."""
    from repair.doc_extract import is_replacement_line

    assert is_replacement_line("1 Пе еход 108х6-89х6 шт. Концент ический Ст. 20")


def test_welding_journal_rows_are_skipped():
    """В комплекте ИД те же трубы перечислены в журнале сварочных работ."""
    from repair.doc_extract import is_replacement_line

    assert not is_replacement_line(
        "Труба 89х5 01.04.26 +15C серт.№ЕС27874О54 Годен №7-УК")


def test_replacements_deduplicated_by_meaning():
    """
    Зона читается в нескольких масштабах, и OCR всякий раз коверкает строку
    по-своему. Дубли отсеиваются по сути: элемент плюс типоразмеры.
    """
    info = extract_doc_info([
        "Отвод П90 219х12 4шт.",
        "отвед 7190 219х12",
        "отвод 690 219х12",
        "Труба 89х5",
    ])
    assert len(info.replacements) == 2
    assert info.replacements[0].startswith("Отвод")


def test_line_from_defect_table_cell():
    """В акте номер линии подписан в первой ячейке таблицы дефектов."""
    info = extract_doc_info([
        "Акт ревизии и отбраковки трубопровода № 222-1-179",
        "N2108 «N2108 ”фильтрат 2,3 ступени от Н-6, Н-6а в Кр-17 1 поток”»",
    ])
    assert info.line == "108"


def test_revision_date_not_taken():
    """Реальный OCR акта 222-1-143: гриф нечитаем («« ДО» ма а»), дальше дата
    РЕВИЗИИ — её брать нельзя. Дата остаётся пустой, номер акта извлекается
    несмотря на «N9» вместо «№»."""
    lines = [
        "УТВЕРЖДАЮ:",
        "« ДО» ма а 2026 г.",
        "Акт ревизии и отбраковки трубо овода N9 222-1-143",
        "по установке 39-40«25» марта 2026 г. произведена ЭПБ трубопровода.",
    ]
    info = extract_doc_info(lines)
    assert info.act_numbers == ["222-1-143"]
    assert info.act_date == ""             # не угадываем — впишет сотрудник
    assert info.unit == "39-40"


def test_multiple_acts_dedupe():
    lines = ["Акт ОТН №222-1-143,121",
             "Акт ревизии и отбраковки трубопровода № 222-1-143",
             "к Акту ревизии и отбраковки трубопровода № 222-1-143"]
    info = extract_doc_info(lines)
    assert info.act_numbers == ["222-1-143"]


def test_merge_fills_only_empty():
    from repair.doc_extract import DocInfo, merge_into_records
    from repair.model import ActRef, RepairRecord, ID_EXPECTED

    rec = RepairRecord(id_status=ID_EXPECTED,
                       acts=[ActRef("222-1-143", ""), ActRef("222-1-121", "24.03.2026")],
                       material="")
    info = DocInfo(act_numbers=["222-1-143"], act_date="30.03.2026",
                   material="Ст. 20", replacements=["Труба 108х6 2 п. м."])
    n = merge_into_records([rec], info)

    assert rec.acts[0].date == "30.03.2026"    # пустая — заполнена
    assert rec.acts[1].date == "24.03.2026"    # была — не тронута
    assert rec.material == "Ст. 20"
    assert "108х6" in rec.replacements
    assert n == 3


def test_merge_no_intersection_untouched():
    from repair.doc_extract import DocInfo, merge_into_records
    from repair.model import ActRef, RepairRecord

    rec = RepairRecord(acts=[ActRef("222-1-999")])
    n = merge_into_records([rec], DocInfo(act_numbers=["222-1-143"],
                                          material="Ст. 20"))
    assert n == 0
    assert rec.material == ""


def test_ocr_artifacts_in_act_number():
    """Реальный грязный скан ИД: дефис как «„», «№» прилип к цифрам."""
    lines = ["Акт ревизии и отбраковки трубу )Овода N2 222-1„84",
             "N9222-5-03-004 и ЛГТ2-5-ОЗ-ОО6 от"]
    info = extract_doc_info(lines)
    assert "222-1-84" in info.act_numbers
    assert "222-5-03-004" in info.act_numbers   # заключение РК — отфильтруется
                                                # ниже по конвейеру привязкой
