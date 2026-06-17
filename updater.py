# -*- coding: utf-8 -*-
"""
Модуль автообновления ПНОС.
Проверяет GitHub Releases, скачивает новый EXE и подменяет текущий.
"""

import os
import sys
import subprocess
import tempfile
import zipfile
import shutil
import requests

from PySide6.QtCore import QThread, Signal

from version import APP_VERSION

# ─── Константы ──────────────────────────────────────────────────
GITHUB_REPO = "azoga99/PNOS"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

# Точное имя ZIP-архива, который нужно прикрепить к GitHub Release.
# GitHub автоматически добавляет «Source code.zip» — мы его игнорируем,
# поэтому имя должно совпадать с тем, что загружает разработчик вручную.
ZIP_ASSET_NAME = "PNOS.zip"


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

    status = Signal(str)             # Текстовый статус для UI
    download_progress = Signal(int)  # 0-100 прогресс скачивания
    finished_ok = Signal(bool, str)  # (успех?, сообщение)

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
                self.finished_ok.emit(
                    False,
                    "На GitHub пока нет ни одного релиза.\n"
                    "Опубликуйте первый Release с EXE-файлом."
                )
                return
            resp.raise_for_status()

            release = resp.json()
            remote_tag = release.get("tag_name", "0.0.0")
            remote_ver = _parse_version(remote_tag)
            local_ver = _parse_version(APP_VERSION)

            if remote_ver <= local_ver:
                self.finished_ok.emit(True, f"✅ У вас уже последняя версия ({APP_VERSION})!")
                return

            # ── ШАГ 2: Ищем ZIP по точному имени ─────────────────
            # Сравниваем без учёта регистра, чтобы не зависеть от того,
            # как именно назван файл на GitHub (ПНОС.zip / пнос.zip и т.д.)
            self.status.emit(f"Найдена новая версия: {remote_tag} (текущая: {APP_VERSION})")

            download_url = None
            file_size = 0
            target_name = ZIP_ASSET_NAME.lower()

            for asset in release.get("assets", []):
                if asset.get("name", "").lower() == target_name:
                    download_url = asset["browser_download_url"]
                    file_size = asset.get("size", 0)
                    break

            if not download_url:
                self.finished_ok.emit(
                    False,
                    f"❌ В релизе {remote_tag} не найден файл «{ZIP_ASSET_NAME}».\n"
                    f"Прикрепите архив с таким именем к Release на GitHub."
                )
                return

            # ── ШАГ 3: Скачиваем ZIP во временную папку системы ──
            # Используем системный %TEMP% вместо папки с EXE:
            # это безопасно при любых правах доступа и любых путях.
            self.status.emit("Скачиваем обновление...")

            if getattr(sys, 'frozen', False):
                current_dir = os.path.dirname(sys.executable)
            else:
                current_dir = os.path.dirname(os.path.abspath(__file__))

            new_exe_path = os.path.join(current_dir, "PNOS_update.exe")
            temp_zip = os.path.join(tempfile.gettempdir(), "pnos_update_pkg.zip")

            dl_resp = requests.get(download_url, stream=True, timeout=120)
            dl_resp.raise_for_status()

            downloaded = 0
            with open(temp_zip, "wb") as f:
                for chunk in dl_resp.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if file_size > 0:
                            self.download_progress.emit(min(int(downloaded / file_size * 100), 99))

            # ── ШАГ 4: Проверка целостности скачанного файла ──────
            # Сравниваем реальный размер с тем, что объявил GitHub.
            # Расхождение = обрыв соединения или повреждение данных.
            if file_size > 0:
                actual_size = os.path.getsize(temp_zip)
                if actual_size != file_size:
                    os.remove(temp_zip)
                    self.finished_ok.emit(
                        False,
                        f"❌ Ошибка загрузки: файл повреждён\n"
                        f"(ожидалось {file_size} байт, получено {actual_size}).\n"
                        f"Попробуйте ещё раз."
                    )
                    return

            self.download_progress.emit(100)
            self.status.emit("Распаковка обновления...")

            # ── ШАГ 5: Распаковываем EXE из архива ───────────────
            # Распаковка идёт в системный TEMP — избегаем проблем с правами
            # доступа, если current_dir находится в сетевой папке или Program Files.
            try:
                with zipfile.ZipFile(temp_zip, "r") as zf:
                    exe_inside = next(
                        (info.filename for info in zf.infolist()
                         if info.filename.lower().endswith(".exe")),
                        None
                    )
                    if not exe_inside:
                        os.remove(temp_zip)
                        self.finished_ok.emit(False, "❌ Внутри ZIP-архива не найден файл .exe")
                        return

                    tmp_dir = tempfile.mkdtemp()
                    try:
                        zf.extract(exe_inside, tmp_dir)
                        extracted = os.path.join(tmp_dir, exe_inside)
                        # os.replace атомарно перезаписывает цель, даже если она уже существует
                        os.replace(extracted, new_exe_path)
                    finally:
                        shutil.rmtree(tmp_dir, ignore_errors=True)

                os.remove(temp_zip)

            except Exception as e:
                try:
                    os.remove(temp_zip)
                except Exception:
                    pass
                self.finished_ok.emit(False, f"❌ Ошибка при распаковке: {e}")
                return

            self.new_exe_path = new_exe_path
            self.status.emit("Обновление готово!")
            self.finished_ok.emit(
                True,
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
    Запускает PowerShell-скрипт, который:
      1. Завершает текущий процесс.
      2. Удаляет старый EXE (с повтором, пока файл не освободится).
      3. Перемещает PNOS_update.exe на место старого.
      4. Убеждается, что файл на месте, и запускает новую версию.
      5. Удаляет временный .ps1 скрипт.

    Почему PowerShell, а не .bat:
      - .bat пишется в ASCII/cp1251, из-за чего пути с кириллицей
        (имя EXE «ПНОС.exe», кириллица в пути к папке) ломаются.
      - PowerShell работает в Unicode и обрабатывает любые пути без
        дополнительных трюков (короткие имена, chcp и т.д.).
      - «-ExecutionPolicy Bypass» разрешает запуск .ps1 для одного
        конкретного процесса, не меняя системную политику.
    """
    if not getattr(sys, 'frozen', False):
        # В режиме разработки просто чистим скачанный файл
        try:
            os.remove(new_exe_path)
        except Exception:
            pass
        return

    current_exe = sys.executable
    current_name = os.path.basename(current_exe)

    def _ps(s: str) -> str:
        """Экранирует строку для одинарных кавычек PowerShell ('...')."""
        return s.replace("'", "''")

    # {{ и }} в f-строке Python → { и } в реальном тексте PS1
    ps1_content = f"""\
# ПНОС auto-updater — сгенерировано приложением, удалится после запуска
$old  = '{_ps(current_exe)}'
$new  = '{_ps(new_exe_path)}'
$name = '{_ps(current_name)}'

taskkill /f /im $name 2>$null | Out-Null
Start-Sleep -Seconds 3

# Удаляем старый EXE (повтор при блокировке антивирусом или ОС)
$tries = 10
while ((Test-Path $old) -and ($tries -gt 0)) {{
    Remove-Item -Path $old -Force -ErrorAction SilentlyContinue
    if (Test-Path $old) {{ Start-Sleep -Seconds 1 }}
    $tries--
}}

# Перемещаем новый EXE на место старого
$moves = 5
while ((-not (Test-Path $old)) -and ($moves -gt 0)) {{
    Move-Item -Path $new -Destination $old -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path $old)) {{ Start-Sleep -Seconds 1 }}
    $moves--
}}

# Запускаем только если файл реально появился
if (Test-Path $old) {{
    Start-Sleep -Seconds 1
    Start-Process -FilePath $old
}} else {{
    [System.Windows.Forms.MessageBox]::Show(
        'Не удалось применить обновление. Запустите программу вручную.',
        'ПНОС Обновление'
    )
}}

Remove-Item -Path $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
"""

    # Пишем в %TEMP% (ASCII-путь) в UTF-8 с BOM — PowerShell 5.1 читает его корректно
    ps1_path = os.path.join(tempfile.gettempdir(), "_pnos_updater.ps1")
    with open(ps1_path, "w", encoding="utf-8-sig") as f:
        f.write(ps1_content)

    subprocess.Popen(
        [
            "powershell.exe",
            "-ExecutionPolicy", "Bypass",
            "-WindowStyle", "Hidden",
            "-NonInteractive",
            "-File", ps1_path,
        ],
        creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
        close_fds=True,
    )
    sys.exit(0)
