# -*- coding: utf-8 -*-
"""
Осмотр всплывающих диалогов Office — только для журнала, ничего не нажимает.

Закрывают диалоги охотники в воркерах (`DialogKiller` для Excel,
`WordDialogKiller` для Word): Enter в окно плюс вывод его на передний план.
Способ требует, чтобы ПНОС был активным приложением, — это осознанное решение
после того, как закрытие «кнопкой по сообщению» стало терять окна Office,
у которых кнопки нарисованы собственными контролами.

Здесь остаётся только диагностика: когда длинный COM-вызов подозрительно
затянулся, в журнал пишется, какие окна висят на экране и что на них написано.
"""

from __future__ import annotations

import win32gui

DIALOG_CLASS = "#32770"


def dialog_buttons(hwnd) -> list[tuple[int, str]]:
    """Кнопки диалога: [(hwnd кнопки, подпись), …]. Отключённые пропускаются."""
    out: list[tuple[int, str]] = []

    def _cb(child, _):
        try:
            if win32gui.GetClassName(child) != "Button":
                return True
            if not win32gui.IsWindowEnabled(child):
                return True
            if not win32gui.IsWindowVisible(child):
                return True
            out.append((child, win32gui.GetWindowText(child) or ""))
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _cb, None)
    except Exception:
        pass
    return out


def describe_dialogs(pid: int) -> list[str]:
    """
    Описание видимых диалогов процесса (для журнала при подозрении на зависание).

    Нужно, когда всплывёт окно, которое охотник закрыть не смог: из журнала
    пользователя будет видно, что именно висело, а не догадки.
    """
    import win32process

    найдено: list[str] = []

    def _cb(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            класс = win32gui.GetClassName(hwnd)
            _, win_pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid and win_pid != pid:
                return True
            заголовок = win32gui.GetWindowText(hwnd) or ""
            if not заголовок:
                return True          # безымянные служебные окна — не диалоги
            подписи = ", ".join(f"«{c}»" for _, c in dialog_buttons(hwnd)) or "—"
            найдено.append(f"«{заголовок}» (класс {класс}, кнопки: {подписи})")
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return найдено
