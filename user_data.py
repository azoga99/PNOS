# -*- coding: utf-8 -*-
"""
Пользовательские списки (специалисты B22, организации ЭПБ B52), которые можно
редактировать прямо в настройках без пересборки программы.

Хранятся в QSettings("PNOS","PNOS"). При первом обращении инициализируются
значениями по умолчанию из config.py. «Сбросить к стандартным» возвращает
списки к config-значениям.
"""

from PySide6.QtCore import QSettings

from config import CONFIG

_KEYS = {
    "specialists": "SPECIALISTS",
    "epb_orgs": "EPB_ORGS",
}


def _settings() -> QSettings:
    return QSettings("PNOS", "PNOS")


def get_list(name: str) -> list:
    """Возвращает актуальный список ('specialists' | 'epb_orgs')."""
    cfg_key = _KEYS[name]
    s = _settings()
    stored = s.value(f"lists/{name}", None)
    if stored is None:
        return list(CONFIG.get(cfg_key, []))
    if isinstance(stored, str):     # QSettings может вернуть строку для списка из 1 элемента
        stored = [stored]
    # Отфильтровываем пустые строки на всякий случай
    return [str(x) for x in stored if str(x).strip()]


def set_list(name: str, items: list):
    """Сохраняет список."""
    cleaned = [str(x).strip() for x in items if str(x).strip()]
    _settings().setValue(f"lists/{name}", cleaned)


def reset_list(name: str):
    """Возвращает список к значениям по умолчанию из config.py."""
    _settings().remove(f"lists/{name}")


def default_list(name: str) -> list:
    return list(CONFIG.get(_KEYS[name], []))
