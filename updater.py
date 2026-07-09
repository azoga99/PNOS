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
    """Собирает текст PowerShell-скрипта бесшовной подмены EXE и перезапуска.
    Вынесено отдельно, чтобы логику можно было протестировать без заморозки.

    Главное отличие от прежней схемы (устраняет «вторую иконку»):
      - НЕ удаляем работающий EXE (это ненадёжно: onefile-загрузчик PyInstaller
        ещё держит файл после выхода, и удаление срывается) — вместо этого
        ПЕРЕИМЕНОВЫВАЕМ старый EXE в сторону (на NTFS это работает даже для
        заблокированного файла), а на его место КЛАДЁМ новый. Swap без удаления
        практически не может провалиться;
      - в ЛЮБОМ исходе на диске остаётся ровно ОДИН файл ПНОС.exe: при успехе
        скачанный PNOS_update.exe переезжает на место старого, при неудаче —
        откатываемся и удаляем скачанный дубль. Второй иконки не остаётся;
      - ждём реального выхода процесса, ведём лог в %TEMP%\\pnos_update.log.
    """
    current_name = os.path.basename(current_exe)
    base_name = os.path.splitext(current_name)[0]
    bak_path = current_exe + ".old"  # старый EXE, сдвинутый в сторону

    def _ps(s: str) -> str:
        """Экранирует строку для одинарных кавычек PowerShell ('...')."""
        return s.replace("'", "''")

    # {{ и }} в f-строке Python → { и } в реальном тексте PS1
    return f"""\
# ПНОС auto-updater — сгенерировано приложением, удалится после запуска
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue | Out-Null
$old  = '{_ps(current_exe)}'
$new  = '{_ps(new_exe_path)}'
$bak  = '{_ps(bak_path)}'
$name = '{_ps(current_name)}'
$base = '{_ps(base_name)}'
$log  = '{_ps(log_path)}'

function L($m) {{ Add-Content -Path $log -Value ((Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') + '  ' + $m) -Encoding UTF8 -ErrorAction SilentlyContinue }}
L "=== apply update: old=$old new=$new ==="

# 1. Завершаем процессы программы и ЖДЁМ полного выхода
taskkill /f /im $name 2>$null | Out-Null
$waited = 0
while ((Get-Process -Name $base -ErrorAction SilentlyContinue) -and ($waited -lt 40)) {{
    Start-Sleep -Milliseconds 500
    $waited++
}}
L "process exited after $([math]::Round($waited * 0.5, 1)) s"

# 2. Убираем возможный бэкап от прошлой попытки
if (Test-Path $bak) {{ Remove-Item -Path $bak -Force -ErrorAction SilentlyContinue }}

# 3. Сдвигаем старый EXE в сторону ПЕРЕИМЕНОВАНИЕМ.
#    На NTFS переименование работает даже если файл ещё заблокирован процессом —
#    в отличие от удаления. Это ключ к надёжности.
$moved = $false
for ($i = 0; $i -lt 30; $i++) {{
    Move-Item -Path $old -Destination $bak -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path $old)) {{ $moved = $true; break }}
    Start-Sleep -Seconds 1
}}
L "old moved aside=$moved"

# 4. Кладём новый EXE на освободившееся место старого
$installed = $false
if ($moved) {{
    for ($i = 0; $i -lt 15; $i++) {{
        Move-Item -Path $new -Destination $old -Force -ErrorAction SilentlyContinue
        if ((Test-Path $old) -and (-not (Test-Path $new))) {{ $installed = $true; break }}
        Start-Sleep -Seconds 1
    }}
}}
L "installed=$installed"

# 5. Запускаем новую версию — бесшовно; либо откатываемся без «второй иконки»
if ($installed) {{
    Start-Sleep -Milliseconds 300
    Start-Process -FilePath $old
    L "relaunched new version"
    # Бэкап больше не нужен. Если занят антивирусом — удалится при след. старте.
    Remove-Item -Path $bak -Force -ErrorAction SilentlyContinue
}} else {{
    L "FAILED to install (moved=$moved)"
    # Откатываем старый EXE на место, чтобы пользователь не остался без программы
    if ((-not (Test-Path $old)) -and (Test-Path $bak)) {{
        Move-Item -Path $bak -Destination $old -Force -ErrorAction SilentlyContinue
    }}
    # Убираем скачанный дубль, чтобы НЕ появилась вторая иконка
    Remove-Item -Path $new -Force -ErrorAction SilentlyContinue
    $msg = 'Не удалось применить обновление автоматически.' + [Environment]::NewLine +
           'Программа запущена в текущей версии — попробуйте обновиться позже.' + [Environment]::NewLine + [Environment]::NewLine +
           ('Лог: ' + $log)
    try {{ [System.Windows.Forms.MessageBox]::Show($msg, 'ПНОС — обновление', 'OK', 'Warning') | Out-Null }} catch {{ }}
    if (Test-Path $old) {{ Start-Process -FilePath $old }}
}}

Remove-Item -Path $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
"""


def _log_line(log_path: str, msg: str):
    """Дописывает строку в лог обновления (не мешает работе при любой ошибке)."""
    try:
        import datetime
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S.%f}  [py] {msg}\n")
    except Exception:
        pass


def _spawn_outside_job(cmdline: str, log_path: str) -> bool:
    """Запускает командную строку так, чтобы процесс ПЕРЕЖИЛ выход приложения.

    КРИТИЧНО: PyInstaller-onefile держит приложение в Windows Job Object с
    флагом «убить все процессы при закрытии» (breakaway запрещён). Обычный
    subprocess.Popen наследует этот Job, поэтому PowerShell-скрипт подмены
    погибал вместе с программой в момент os._exit — обновление «не применялось,
    оставался PNOS_update.exe». Проверено журналом Windows PowerShell: движок
    в момент обновления даже не стартовал.

    Решение: создавать процесс через WMI Win32_Process.Create — его родителем
    становится служба WMI (WmiPrvSE), которая в наш Job не входит.
    Возвращает True при успешном запуске вне Job.
    """
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            wmi = win32com.client.GetObject("winmgmts:root\\cimv2")
            startup = wmi.Get("Win32_ProcessStartup").SpawnInstance_()
            startup.ShowWindow = 0  # SW_HIDE — без мелькания окна
            proc_cls = wmi.Get("Win32_Process")
            method = proc_cls.Methods_("Create")
            params = method.InParameters.SpawnInstance_()
            params.CommandLine = cmdline
            params.ProcessStartupInformation = startup
            res = proc_cls.ExecMethod_("Create", params)
            rc = int(res.ReturnValue)
            pid = getattr(res, "ProcessId", None)
            _log_line(log_path, f"launcher: wmi rc={rc} pid={pid}")
            if rc == 0:
                return True
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        _log_line(log_path, f"launcher: wmi failed: {e}")
    return False


def apply_update_and_restart(new_exe_path: str):
    """
    Запускает PowerShell-скрипт, который:
      1. Завершает текущий процесс.
      2. Сдвигает старый EXE переименованием и кладёт новый на его место.
      3. Убеждается, что файл на месте, и запускает новую версию.
      4. Удаляет временный .ps1 скрипт.

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

    # Полный путь к powershell.exe — надёжнее, чем расчёт на PATH у замороженного EXE.
    ps_exe = os.path.join(
        os.environ.get("SystemRoot", r"C:\Windows"),
        "System32", "WindowsPowerShell", "v1.0", "powershell.exe",
    )
    if not os.path.exists(ps_exe):
        ps_exe = "powershell.exe"

    cmdline = (
        f'"{ps_exe}" -ExecutionPolicy Bypass -WindowStyle Hidden '
        f'-NonInteractive -File "{ps1_path}"'
    )

    # Основной путь: WMI — скрипт переживает закрытие Job-объекта PyInstaller.
    launched = _spawn_outside_job(cmdline, log_path)

    if not launched:
        # Аварийный путь: старый Popen. Внутри Job он, скорее всего, будет убит
        # вместе с приложением, но это лучше, чем ничего (и след остаётся в логе).
        _log_line(log_path, "launcher: fallback to Popen (may be killed by job)")
        subprocess.Popen(
            [
                ps_exe,
                "-ExecutionPolicy", "Bypass",
                "-WindowStyle", "Hidden",
                "-NonInteractive",
                "-File", ps1_path,
            ],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
            close_fds=True,
        )

    # Форсированно завершаем процесс. sys.exit() внутри Qt-слота лишь бросает
    # SystemExit, который цикл событий может проглотить — тогда EXE останется
    # запущенным и заблокированным. os._exit гарантирует немедленный выход и
    # разблокировку файла для подмены.
    os._exit(0)


def cleanup_update_leftovers():
    """Удаляет остаточные файлы обновления рядом с EXE.

    Вызывается при старте программы. Подчищает:
      - <EXE>.old — старую версию, сдвинутую в сторону при прошлом обновлении
        (могла не удалиться сразу, если файл был занят антивирусом);
      - PNOS_update.exe — недокачанный/незавершённый апдейт.
    Так на рабочем столе не копятся лишние иконки.
    """
    if not getattr(sys, "frozen", False):
        return
    current_dir = os.path.dirname(sys.executable)
    for leftover in (sys.executable + ".old", os.path.join(current_dir, "PNOS_update.exe")):
        try:
            if os.path.exists(leftover):
                os.remove(leftover)
        except Exception:
            pass  # занят — уберём при следующем запуске
