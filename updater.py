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


def build_updater_ps1(current_exe: str, new_exe_path: str, log_path: str) -> str:
    """Собирает текст PowerShell-скрипта подмены EXE и перезапуска.
    Вынесено отдельно, чтобы логику можно было протестировать без заморозки.

    Отличия от прежней версии (устраняют молчаливый сбой обновления):
      - ждём РЕАЛЬНОГО выхода процесса (опрос по имени), а не фиксированный Sleep;
      - долгий повтор удаления (до 30с) — переживает блокировку файла антивирусом;
      - при неудаче показываем понятное окно с инструкцией (Add-Type загружен!),
        а не падаем молча; и всё равно запускаем то, что есть, — пользователь
        не остаётся без программы;
      - ведём лог в %TEMP%\\pnos_update.log для диагностики.
    """
    current_name = os.path.basename(current_exe)
    base_name = os.path.splitext(current_name)[0]

    def _ps(s: str) -> str:
        """Экранирует строку для одинарных кавычек PowerShell ('...')."""
        return s.replace("'", "''")

    # {{ и }} в f-строке Python → { и } в реальном тексте PS1
    return f"""\
# ПНОС auto-updater — сгенерировано приложением, удалится после запуска
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue | Out-Null
$old  = '{_ps(current_exe)}'
$new  = '{_ps(new_exe_path)}'
$name = '{_ps(current_name)}'
$base = '{_ps(base_name)}'
$log  = '{_ps(log_path)}'

function L($m) {{ Add-Content -Path $log -Value ((Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + '  ' + $m) -Encoding UTF8 -ErrorAction SilentlyContinue }}
L "=== apply update: old=$old new=$new ==="

# 1. Завершаем процессы программы и ЖДЁМ полного выхода (тогда файл разблокируется)
taskkill /f /im $name 2>$null | Out-Null
$waited = 0
while ((Get-Process -Name $base -ErrorAction SilentlyContinue) -and ($waited -lt 40)) {{
    Start-Sleep -Milliseconds 500
    $waited++
}}
L "process exited after $([math]::Round($waited * 0.5, 1)) s"

# 2. Удаляем старый EXE — долгий повтор на случай блокировки антивирусом/ОС
$deleted = $false
for ($i = 0; $i -lt 30; $i++) {{
    Remove-Item -Path $old -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path $old)) {{ $deleted = $true; break }}
    Start-Sleep -Seconds 1
}}
L "old deleted=$deleted"

# 3. Перемещаем новый EXE на место старого
if ($deleted) {{
    for ($i = 0; $i -lt 10; $i++) {{
        Move-Item -Path $new -Destination $old -Force -ErrorAction SilentlyContinue
        if (Test-Path $old) {{ break }}
        Start-Sleep -Seconds 1
    }}
}}

$installed = (Test-Path $old) -and (-not (Test-Path $new))
L "installed=$installed"

# 4. Запускаем новую версию, либо честно сообщаем об ошибке с инструкцией
if ($installed) {{
    Start-Sleep -Milliseconds 500
    Start-Process -FilePath $old
    L "relaunched new version"
}} else {{
    L "FAILED to install (old deleted=$deleted)"
    $msg = 'Не удалось применить обновление автоматически.' + [Environment]::NewLine + [Environment]::NewLine +
           'Новая версия уже скачана. Закройте программу, удалите старый файл ' +
           'и переименуйте "' + [IO.Path]::GetFileName($new) + '" в "' + $name + '".' + [Environment]::NewLine + [Environment]::NewLine +
           ('Лог: ' + $log)
    try {{ [System.Windows.Forms.MessageBox]::Show($msg, 'ПНОС — обновление', 'OK', 'Warning') | Out-Null }} catch {{ }}
    # Чтобы пользователь не остался без программы — запускаем то, что доступно
    if (Test-Path $old) {{ Start-Process -FilePath $old }}
    elseif (Test-Path $new) {{ Start-Process -FilePath $new }}
}}

Remove-Item -Path $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
"""


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
    log_path = os.path.join(tempfile.gettempdir(), "pnos_update.log")
    ps1_content = build_updater_ps1(current_exe, new_exe_path, log_path)

    # Пишем в %TEMP% (ASCII-путь) в UTF-8 с BOM — PowerShell 5.1 читает его корректно.
    # BOM обязателен: без него PS 5.1 читает файл как cp1251 и кириллическое имя
    # «ПНОС.exe» превращается в мусор → подмена срывается.
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
