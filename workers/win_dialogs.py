# -*- coding: utf-8 -*-
"""
Закрытие всплывающих диалогов Office без передачи им фокуса.

Зачем отдельный модуль. Охотники за диалогами (`DialogKiller` для Excel,
`WordDialogKiller` для Word) раньше жали Enter так: `PostMessage` с VK_RETURN,
а если не помогло — `SetForegroundWindow` плюс `WScript.Shell.SendKeys("~")`.
Оба приёма зависят от фокуса:

  * VK_RETURN разбирается в `IsDialogMessage` через `GetFocus()` — у диалога,
    который не активен, фокуса нет;
  * `SetForegroundWindow` Windows разрешает только процессу, который сам сейчас
    на переднем плане.

Отсюда живой симптом на этапе «Создание отчёта»: Word показывает подтверждение
SQL-команды при `MailMerge.OpenDataSource`, и слияние стоит до тех пор, пока
пользователь не кликнет по окну ПНОС. А когда фокус есть, `SendKeys` шлёт Enter
в активное окно — то есть может прилететь в чужое приложение.

Здесь диалог закрывается нажатием на конкретную кнопку сообщением BM_CLICK.
Это работает между процессами и не требует ни фокуса, ни переднего плана.
"""

from __future__ import annotations

import win32con
import win32gui

#: Подписи кнопок в порядке предпочтения. Диалоги Office двуязычны в
#: зависимости от установленного языкового пакета, поэтому и русские, и
#: английские. «Да» стоит раньше «ОК»: у подтверждения SQL-команды кнопки
#: именно «Да»/«Нет», и «ОК» там нет вовсе.
BUTTON_PRIORITY = ("да", "yes", "ок", "ok", "продолжить", "continue")

DIALOG_CLASS = "#32770"

#: Классы главных окон Office — их трогать не нужно. Всё остальное считаем
#: диалогом: у части диалогов Excel кнопки нарисованы собственными контролами,
#: класса «Button» у них нет, и признак «есть дочерние кнопки» отсекал их
#: вместе с мусором — окно переставало закрываться вовсе.
MAIN_WINDOW_CLASSES = ("OpusApp", "XLMAIN")


def is_dialog_class(class_name: str) -> bool:
    """
    Считать ли окно с таким классом диалогом, который надо закрыть.

    Диалогом считается всё, кроме самих окон приложений. Проверять «класс
    равен #32770» нельзя: часть окон Office этому классу не соответствует, и
    такие окна переставали закрываться вовсе.
    """
    return (class_name or "") not in MAIN_WINDOW_CLASSES


def _clean(caption: str) -> str:
    """Подпись кнопки без ускорителя и пробелов: «&Да» → «да»."""
    return (caption or "").replace("&", "").strip().lower()


def pick_dialog_button(captions) -> int | None:
    """
    Индекс кнопки, которую надо нажать, или None.

    Чистая функция над списком подписей — тестируется без Windows API.
    Порядок важнее позиции: «Нет» рядом с «Да» выбран не будет никогда.
    """
    подписи = [_clean(c) for c in captions]
    for нужная in BUTTON_PRIORITY:
        for i, c in enumerate(подписи):
            if c == нужная:
                return i
    return None


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


def click_dialog_button(hwnd, log=None, escalate: bool = False) -> str | None:
    """
    Нажать подходящую кнопку диалога. Возвращает её подпись или None.

    Сообщения отправляются через PostMessage — вызов не блокируется, даже если
    окно Office занято своим модальным циклом.

    `escalate` — окно уже видели на прошлом проходе, BM_CLICK не помог: тогда
    добавляем WM_COMMAND от имени диалога. Сразу оба сообщения слать нельзя:
    если BM_CLICK сработал, второе нажатие может проскочить в следующий шаг
    мастера.
    """
    кнопки = dialog_buttons(hwnd)
    if not кнопки:
        return None
    idx = pick_dialog_button([c for _, c in кнопки])
    if idx is None:
        return None

    btn, caption = кнопки[idx]
    try:
        win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)
        if escalate:
            ctrl_id = win32gui.GetDlgCtrlID(btn)
            if ctrl_id:
                win32gui.PostMessage(
                    hwnd, win32con.WM_COMMAND,
                    (win32con.BN_CLICKED << 16) | (ctrl_id & 0xFFFF), btn)
    except Exception:
        return None

    if log:
        title = win32gui.GetWindowText(hwnd) or "[без заголовка]"
        log(f"   [DialogKiller] диалог «{title}» закрыт кнопкой "
            f"«{caption.replace('&', '')}»")
    return caption


def describe_dialogs(pid: int) -> list[str]:
    """
    Описание видимых диалогов процесса (для журнала при подозрении на зависание).

    Нужно, когда всплывёт окно с непривычной раскладкой кнопок: из журнала
    пользователя будет видно, что именно висело, а не догадки.
    """
    import win32process

    найдено: list[str] = []

    def _cb(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            класс = win32gui.GetClassName(hwnd)
            if not is_dialog_class(класс):
                return True
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
