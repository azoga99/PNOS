# -*- coding: utf-8 -*-
"""
Выбор кнопки в диалоге Office (workers.win_dialogs).

Функция чистая — тестируется без Windows API и без запущенного Office.
"""

from workers.win_dialogs import pick_dialog_button


def test_yes_wins_over_no():
    """«Нет» стоит первой, но нажата должна быть «Да»."""
    assert pick_dialog_button(["&Нет", "&Да"]) == 1


def test_ok_when_no_yes():
    assert pick_dialog_button(["Отмена", "OK"]) == 1


def test_russian_ok():
    assert pick_dialog_button(["ОК", "Отмена"]) == 0


def test_ampersand_and_spaces_ignored():
    assert pick_dialog_button(["  &Да  "]) == 0


def test_none_when_nothing_suitable():
    """Диалог без утвердительной кнопки закрывать нечем — решает человек."""
    assert pick_dialog_button(["Отмена", "Справка"]) is None
    assert pick_dialog_button([]) is None


def test_no_is_never_picked():
    """Страховка от «Нет»/«Cancel»: слепое нажатие первой кнопки недопустимо."""
    for подписи in (["Нет"], ["Cancel"], ["Отмена", "Нет"]):
        assert pick_dialog_button(подписи) is None
