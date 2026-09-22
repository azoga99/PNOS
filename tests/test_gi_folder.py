# -*- coding: utf-8 -*-
"""
Папка актов ГИ на диске для номера участка из паспорт!B12 (этап 7).

Раньше папка искалась нечётко (rapidfuzz, порог 85) по номеру, у которого
тире уже заменены дефисом, а у имён на диске — нет. Итог в обе стороны:
«АВТ–5» (длинное тире на диске) не находилась вовсе, а «тит.521» находил
чужую папку «тит.512». УНСО тит.80–131/132 тоже попадали в «УНСО
тит.80-130» по похожести — там акты действительно общие, и теперь это
записано явно в GI_CONFIG.FOLDER_ALIASES.
"""

import pytest

from config import CONFIG
from workers.stage3_worker import normalize_gi_number
from workers.stage6_worker import gi_folder_key, match_gi_folder

# /ПНОС на диске шаблонов, 23.09.2026
DISK = ["302-01 (УПСК)", "36-30", "37-10", "37-40", "37-50", "39-10", "39-40",
        "АВТ-2", "АВТ–5", "УНСО тит.80-130", "УХВПтит.517", "тит.302-01М", "тит.512"]


def _gi(place):
    """Номер участка так, как его сохраняет этап 4 (gi_number)."""
    return normalize_gi_number(place)


def test_long_dash_on_disk_is_found():
    """п.3224: в паспорте «АВТ–5», сохранено «АВТ-5», на диске «АВТ–5»."""
    assert match_gi_folder(DISK, _gi("АВТ–5")) == ("АВТ–5", [], "ok")


@pytest.mark.parametrize("place, folder", [
    ("39-40", "39-40"),
    ("АВТ-2", "АВТ-2"),
    ("302-01 (УПСК)", "302-01 (УПСК)"),        # пробел в имени папки
    ("УХВП тит.517", "УХВПтит.517"),            # пробел только в графике
    ("УНСО тит.80-130", "УНСО тит.80-130"),
    ("тит.302-01М", "тит.302-01М"),
])
def test_existing_matches_are_kept(place, folder):
    assert match_gi_folder(DISK, _gi(place))[0] == folder


@pytest.mark.parametrize("place, wrong", [
    ("тит.521", "тит.512"),
    ("тит. 531", "тит.512"),
    ("тит. 520", "тит.512"),
])
def test_other_installation_is_never_taken(place, wrong):
    """Похожая папка — только подсказка в журнале, акты из неё не вставляются."""
    name, hint, reason = match_gi_folder(DISK, _gi(place), ALIASES)
    assert name is None and reason == "missing"
    assert wrong in hint


# ─── Общие акты (GI_CONFIG.FOLDER_ALIASES) ─────────────────────────────

ALIASES = CONFIG["GI_CONFIG"]["FOLDER_ALIASES"]


def test_80_131_shares_acts_with_80_130():
    """Акт в папке 80-130 назван «тит.80-130, 80-80, 80-131» — записано явно."""
    assert match_gi_folder(DISK, _gi("УНСО тит.80–131"), ALIASES) == \
        ("УНСО тит.80-130", [], "alias")


def test_80_132_is_not_in_that_act():
    """80-132 в названии акта нет — ждёт свою папку, а не берёт общую."""
    name, hint, reason = match_gi_folder(DISK, _gi("УНСО тит.80–132"), ALIASES)
    assert name is None and reason == "missing" and hint == ["УНСО тит.80-130"]


def test_own_folder_beats_alias():
    """Появится своя папка «УНСО тит.80-132» — брать её, а не общую."""
    disk = DISK + ["УНСО тит.80–132"]
    assert match_gi_folder(disk, _gi("УНСО тит.80-132"), ALIASES) == \
        ("УНСО тит.80–132", [], "ok")


def test_alias_to_missing_folder_is_missing():
    name, _, reason = match_gi_folder(["39-40"], _gi("УНСО тит.80-131"), ALIASES)
    assert name is None and reason == "missing"


def test_avt5_is_not_confused_with_avt2():
    """Порог снизить нельзя: «АВТ-2» так же похожа на «АВТ-5», как «АВТ–5»."""
    name, _, _ = match_gi_folder(["АВТ-2"], _gi("АВТ–5"))
    assert name is None


def test_two_spellings_on_disk_is_ambiguous():
    name, hint, reason = match_gi_folder(["АВТ-5", "АВТ–5"], "АВТ-5")
    assert name is None and reason == "ambiguous"
    assert sorted(hint) == ["АВТ-5", "АВТ–5"]


def test_nothing_similar_gives_no_hint():
    assert match_gi_folder(DISK, "Изомеризация") == (None, [], "missing")


@pytest.mark.parametrize("a, b", [
    ("АВТ–5", "АВТ-5"), ("АВТ—5", "авт-5"), (" 39 – 40 ", "39-40"),
    ("УНСО тит.80-130", "УНСОтит.80-130"),
])
def test_key_ignores_dashes_spaces_and_case(a, b):
    assert gi_folder_key(a) == gi_folder_key(b)
