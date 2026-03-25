# -*- coding: utf-8 -*-
"""
Модуль автообновления ПНОС.
Проверяет GitHub Releases, скачивает новый EXE и подменяет текущий.
"""

import os
import sys
import tempfile
import requests

from PySide6.QtCore import QThread, Signal

from version import APP_VERSION

# ─── Константы ──────────────────────────────────────────────────
GITHUB_REPO = "azoga99/PNOS"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
EXE_ASSET_NAME = "PNOS.exe"  # Имя файла в GitHub Release Assets


def _parse_version(tag: str) -> tuple:
    """Превращает строку вида 'v1.2.3' или '1.2.3' в кортеж (1, 2, 3)."""
    tag = tag.lstrip("vV")
    parts = []
    for p in tag.split("."):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts)


class UpdateWorker(QThread):
    """Фоновый поток: проверка + скачивание обновлений с GitHub Releases."""

    # Сигналы
    status = Signal(str)            # Текстовый статус для UI
    download_progress = Signal(int) # 0-100 прогресс скачивания
    finished_ok = Signal(bool, str) # (успех?, сообщение)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.new_exe_path = ""

    def run(self):
        try:
            # ── ШАГ 1: Проверка последнего релиза ─────────────────
            self.status.emit("Проверяем наличие обновлений на GitHub...")
            self.download_progress.emit(0)

            resp = requests.get(GITHUB_API_URL, timeout=15)
            if resp.status_code == 404:
                self.finished_ok.emit(False, "На GitHub пока нет ни одного релиза.\nОпубликуйте первый Release с EXE-файлом.")
                return
            resp.raise_for_status()

            release = resp.json()
            remote_tag = release.get("tag_name", "0.0.0")
            remote_ver = _parse_version(remote_tag)
            local_ver = _parse_version(APP_VERSION)

            if remote_ver <= local_ver:
                self.finished_ok.emit(True, f"✅ У вас уже последняя версия ({APP_VERSION})!")
                return

            # ── ШАГ 2: Ищем EXE-файл в Assets ────────────────────
            self.status.emit(f"Найдена новая версия: {remote_tag} (текущая: {APP_VERSION})")

            download_url = None
            file_size = 0
            for asset in release.get("assets", []):
                name = asset.get("name", "").lower()
                if name.endswith(".exe"):
                    download_url = asset["browser_download_url"]
                    file_size = asset.get("size", 0)
                    break

            if not download_url:
                self.finished_ok.emit(False, "❌ В релизе не найден EXE-файл.\nПрикрепите скомпилированный .exe к Release на GitHub.")
                return

            # ── ШАГ 3: Скачиваем новый EXE ────────────────────────
            self.status.emit("Скачиваем обновление...")

            # Определяем путь для скачивания (рядом с текущим EXE)
            if getattr(sys, 'frozen', False):
                # Скомпилированный EXE
                current_dir = os.path.dirname(sys.executable)
            else:
                # Режим разработки (python main.py)
                current_dir = os.path.dirname(os.path.abspath(__file__))

            new_exe_path = os.path.join(current_dir, "PNOS_update.exe")

            dl_resp = requests.get(download_url, stream=True, timeout=120)
            dl_resp.raise_for_status()

            downloaded = 0
            with open(new_exe_path, "wb") as f:
                for chunk in dl_resp.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if file_size > 0:
                            pct = int(downloaded / file_size * 100)
                            self.download_progress.emit(min(pct, 100))

            self.download_progress.emit(100)
            self.new_exe_path = new_exe_path
            self.status.emit("Скачивание завершено!")

            self.finished_ok.emit(True,
                f"✅ Версия {remote_tag} скачана!\n"
                f"Нажмите «Установить и перезапустить», чтобы применить обновление."
            )

        except requests.exceptions.ConnectionError:
            self.finished_ok.emit(False, "❌ Нет подключения к интернету.")
        except requests.exceptions.Timeout:
            self.finished_ok.emit(False, "❌ Время ожидания истекло (сервер не ответил).")
        except requests.exceptions.HTTPError as e:
            self.finished_ok.emit(False, f"❌ Ошибка HTTP: {e}")
        except Exception as e:
            self.finished_ok.emit(False, f"❌ Неизвестная ошибка: {e}")


def apply_update_and_restart(new_exe_path: str):
    """
    Создаёт .bat скрипт, который:
    1. Ждёт 2 секунды (пока текущий EXE закроется).
    2. Удаляет старый EXE.
    3. Переименовывает PNOS_update.exe → PNOS.exe.
    4. Запускает новый EXE.
    5. Удаляет сам себя.
    """
    if getattr(sys, 'frozen', False):
        current_exe = sys.executable
    else:
        # В режиме разработки — просто удаляем скачанный файл
        try:
            os.remove(new_exe_path)
        except Exception:
            pass
        return

    current_dir = os.path.dirname(current_exe)
    current_name = os.path.basename(current_exe)
    new_name = os.path.basename(new_exe_path)

    bat_path = os.path.join(current_dir, "_updater.bat")

    bat_content = f"""@echo off
chcp 65001 >nul
echo Обновление ПНОС...
timeout /t 2 /nobreak >nul
del /f /q "{current_exe}"
rename "{new_exe_path}" "{current_name}"
start "" "{os.path.join(current_dir, current_name)}"
del /f /q "{bat_path}"
"""

    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat_content)

    # Запускаем батник и немедленно выходим
    os.startfile(bat_path)
    sys.exit(0)
