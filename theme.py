# -*- coding: utf-8 -*-
"""
Система тем оформления ПНОС.

QSS не поддерживает переменные, поэтому style.qss хранится как шаблон с токенами
вида @token@. Здесь заданы две палитры (светлая и тёмная) и функция, которая
подставляет цвета в шаблон и возвращает готовый стиль.

Выбор темы хранится в QSettings, поэтому все тяжёлые импорты не нужны — модуль
лёгкий и может импортироваться при старте.
"""

import os

from utils import resource_path

# ─── Палитры ──────────────────────────────────────────────────────
# Ключи — семантические (назначение цвета), а не сами цвета. Это позволяет
# держать один QSS-шаблон и переключать темы простой заменой словаря.

LIGHT = {
    "bg_app": "#eef1f5",
    "bg_content": "#f1f5f9",
    "bg_panel": "#ffffff",
    "bg_panel_alt": "#f8fafc",
    "bg_input": "#f8fafc",
    "bg_input_focus": "#ffffff",
    "bg_rail": "#ffffff",
    "bg_rail_hover": "#f1f5f9",
    "border": "#e2e8f0",
    "border_strong": "#cbd5e1",
    "text_primary": "#1e293b",
    "text_secondary": "#475569",
    "text_muted": "#64748b",
    "text_faint": "#94a3b8",
    "accent": "#6366f1",
    "accent_hover": "#7c7ff7",
    "accent_press": "#4f46e5",
    "accent_soft": "#eef2ff",
    "accent2": "#8b5cf6",
    "accent2_hover": "#9d78fb",
    "accent2_press": "#7c3aed",
    "teal": "#2dd4bf",        # только для статусов, не для акцента
    "teal_dark": "#14b8a6",
    "success": "#10b981",
    "success_strong": "#16a34a",
    "success_hover": "#15803d",
    "error": "#ef4444",
    "error_strong": "#dc2626",
    "error_hover": "#f87171",
    "warn": "#f59e0b",
    "warn_bg": "#fef9c3",
    "warn_border": "#fde047",
    "info_blue": "#2563eb",
    "info_blue_hover": "#1d4ed8",
    "rail_text": "#64748b",
    "rail_text_hover": "#1e293b",
    "dark_panel": "#1e293b",
    "dark_panel_border": "#334155",
    "dark_panel_text": "#94a3b8",
    "row_alt": "#f8fafc",
    "selection": "#eff6ff",
    "scrollbar": "#94a3b8",
    "chip_new_bg": "#dcfce7",
    "chip_new_fg": "#15803d",
    "chip_improved_bg": "#dbeafe",
    "chip_improved_fg": "#1d4ed8",
    "chip_fixed_bg": "#fef3c7",
    "chip_fixed_fg": "#b45309",
    "chip_changed_bg": "#e2e8f0",
    "chip_changed_fg": "#475569",
}

DARK = {
    "bg_app": "#0b1120",
    "bg_content": "#0f172a",
    "bg_panel": "#1e293b",
    "bg_panel_alt": "#172033",
    "bg_input": "#0f172a",
    "bg_input_focus": "#1e293b",
    "bg_rail": "#0b1120",
    "bg_rail_hover": "#1e293b",
    "border": "#334155",
    "border_strong": "#475569",
    "text_primary": "#f1f5f9",
    "text_secondary": "#cbd5e1",
    "text_muted": "#94a3b8",
    "text_faint": "#64748b",
    "accent": "#818cf8",
    "accent_hover": "#a5b4fc",
    "accent_press": "#6366f1",
    "accent_soft": "#262a56",
    "accent2": "#a78bfa",
    "accent2_hover": "#c4b5fd",
    "accent2_press": "#8b5cf6",
    "teal": "#2dd4bf",        # только для статусов, не для акцента
    "teal_dark": "#14b8a6",
    "success": "#34d399",
    "success_strong": "#16a34a",
    "success_hover": "#15803d",
    "error": "#f87171",
    "error_strong": "#dc2626",
    "error_hover": "#fca5a5",
    "warn": "#fbbf24",
    "warn_bg": "#422006",
    "warn_border": "#a16207",
    "info_blue": "#3b82f6",
    "info_blue_hover": "#2563eb",
    "rail_text": "#94a3b8",
    "rail_text_hover": "#e2e8f0",
    "dark_panel": "#0f172a",
    "dark_panel_border": "#334155",
    "dark_panel_text": "#94a3b8",
    "row_alt": "#172033",
    "selection": "#1e3a5f",
    "scrollbar": "#475569",
    "chip_new_bg": "#14352a",
    "chip_new_fg": "#34d399",
    "chip_improved_bg": "#172a4a",
    "chip_improved_fg": "#60a5fa",
    "chip_fixed_bg": "#3a2a0a",
    "chip_fixed_fg": "#fbbf24",
    "chip_changed_bg": "#334155",
    "chip_changed_fg": "#cbd5e1",
}

PALETTES = {"light": LIGHT, "dark": DARK}


def get_palette(theme_name: str) -> dict:
    """Возвращает словарь цветов по имени темы ('light' / 'dark')."""
    return PALETTES.get(theme_name, LIGHT)


def build_stylesheet(theme_name: str) -> str:
    """
    Читает шаблон resources/style.qss и подставляет в него цвета выбранной темы.
    Токены в шаблоне имеют вид @имя_цвета@.
    """
    palette = get_palette(theme_name)
    template_path = resource_path(os.path.join("resources", "style.qss"))
    try:
        with open(template_path, "r", encoding="utf-8") as f:
            template = f.read()
    except Exception:
        return ""

    for key, color in palette.items():
        template = template.replace(f"@{key}@", color)
    return template


def color(theme_name: str, key: str) -> str:
    """Достаёт конкретный цвет палитры (для случаев, где QSS не применить)."""
    return get_palette(theme_name).get(key, "#000000")
