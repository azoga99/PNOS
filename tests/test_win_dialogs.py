# -*- coding: utf-8 -*-
"""
Модуль осмотра диалогов Office (workers.win_dialogs).

Закрывают диалоги охотники в воркерах — здесь только диагностика для журнала,
поэтому проверяем ровно то, что можно проверить без запущенного Office.
"""

import workers.win_dialogs as wd


def test_module_does_not_press_anything():
    """
    Страховка от возврата «нажатия кнопкой по сообщению».

    Способ терял окна Office, у которых кнопки нарисованы собственными
    контролами: диалог Excel переставал закрываться вовсе. Закрытие живёт в
    охотниках воркеров и требует активного окна ПНОС — так решено осознанно.
    """
    for имя in ("click_dialog_button", "pick_dialog_button"):
        assert not hasattr(wd, имя), (
            f"{имя} вернулась в win_dialogs — закрытие диалогов должно "
            f"оставаться в охотниках воркеров")


def test_public_api_is_diagnostics_only():
    assert hasattr(wd, "dialog_buttons")
    assert hasattr(wd, "describe_dialogs")


def test_describe_dialogs_survives_unknown_pid():
    """Несуществующий процесс — пустой список, а не исключение."""
    assert wd.describe_dialogs(0x7FFFFFFF) == []
