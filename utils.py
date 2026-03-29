# -*- coding: utf-8 -*-
import os
import sys

def resource_path(relative_path: str) -> str:
    """Получает абсолютный путь к ресурсу, работает в разработке и в EXE."""
    try:
        # PyInstaller создает временную папку и сохраняет путь в _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
