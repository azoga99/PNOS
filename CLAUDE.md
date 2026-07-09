# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

ПНОС (Пакетный генератор отчётов) — a Windows desktop app that automates batch generation of industrial pipeline inspection reports (ЭПБ). It downloads folders from Yandex Disk, processes Word/Excel files via COM automation, runs macros, and generates final Word reports.

## Running and Building

```powershell
# Install dependencies (no requirements.txt exists — install manually)
pip install PySide6 aiohttp aiofiles python-docx openpyxl Pillow rapidfuzz pandas requests pywin32

# Run from source
python main.py

# Build a standalone EXE (output: dist/ПНОС.exe)
pyinstaller PNOS.spec
```

No test suite exists. Manual testing requires running the full GUI. Windows-only: COM automation depends on installed Microsoft Word and Excel.

## Architecture

### Entry flow
`main.py` (shows `SplashLoader` — program icon with a spinning arc — while the window builds) → `MainWindow` (`main_window.py`, PySide6 QMainWindow) → `QStackedWidget` with 5 pages: **Процесс** (pipeline of 7 `StepCard`s + launch params + single log), **Ручной контроль** (step-3 form), **Обзор** (status matrix), **Итоговый отчёт** (per-step result tables built from `_stage_results`), **Настройки**. Workers live in `workers/stage*_worker.py` (QThreads).

### Design rules
- Strict business look: **no emojis** in UI text (the only exception is 📂 on folder/file pickers), no gradients; colors only via QSS tokens.
- Statuses are shown with colored dots (`widgets.dot_icon`) and `StepCard` badges (`STATUS_COLORS`: idle/run/done/error/wait), not glyph emojis.
- Per-step result tables/metrics bars were deliberately removed from step UI — results go to the Обзор matrix and the Итоговый отчёт page only; each step card shows a one-line summary.

### Startup performance — lazy imports (do not break)
Heavy modules (pandas, aiohttp, win32com, python-docx) are deliberately **not** imported at module top level in `main.py` / `main_window.py`. Workers and `report_dialog` are imported lazily inside the stage-launch methods, and `main()` prewarms them in a background daemon thread after the window is shown (`_prewarm_heavy_imports`). Adding a top-level import of a worker or heavy library to these files silently reverts the fast-startup optimization.

### Key files
| File | Purpose |
|---|---|
| `main_window.py` | All pages: the Процесс pipeline (`_setup_process_page`, step cards wired to `_start_stage*`), manual-control form (`_setup_manual_page`), Обзор matrix (`_setup_overview_page`/`_refresh_overview`), Итоговый отчёт (`_setup_report_page`/`_refresh_report_page`), single activity log (`_add_activity`), auto-mode chaining in `_on_step_finished` (Settings checkbox `chk_auto`; stops at manual step 3 with a "wait" card), and `_force_stop_active_stage` (emergency stop) |
| `widgets.py` | `StepCard` (pipeline card: status badge, summary line, action/stop buttons, thin progress bar, hover shadow animation) and `dot_icon(color)` (colored-dot QIcon used instead of emojis in lists) |
| `theme.py` | Light/dark theme system. `style.qss` is a token template (`@name@`); `build_stylesheet(theme)` substitutes a palette dict (`LIGHT`/`DARK`) into it. Theme choice persists in `QSettings("PNOS","PNOS")` key `theme`. Never hard-code colors in QSS — add a token to both palettes. |
| `pages/settings_page.py` | Settings page — auto-mode checkbox, theme toggle (emits `theme_changed`), master-file path override, download speed toggle, update check & install buttons |
| `config.py` | Single source of truth — Yandex Disk OAuth tokens, remote paths, folder names, macro paths |
| `version.py` | `APP_VERSION` string — bump before every release |
| `status_manager.py` | Per-point incremental tracking via hidden `.pnos_status.json` in each point folder |
| `yandex_api.py` | Fully async Yandex Disk REST client (aiohttp + aiofiles, semaphore-throttled) |
| `excel_service.py` | Reads the master Excel schedule and filters rows by columns E/F/H |
| `updater.py` | Checks GitHub Releases (`azoga99/PNOS`), downloads `PNOS.zip`, extracts the `.exe`, then swaps it in via a generated PowerShell script and relaunches. See "Self-update mechanism" below. |
| `utils.py` | `resource_path()` — resolves paths for both dev and PyInstaller frozen EXE |

### Worker → UI signal contract
All workers (`QThread`) expose the same signal set:
- `log = Signal(str)` — verbose technical log
- `info = Signal(str, str)` — `(message, category)` for the activity feed; categories: `info / done / wait / warn / error`
- `progress = Signal(int)` — 0–100
- `finished_ok = Signal(bool)` — success/failure
- `report_ready = Signal(...)` — structured result for the inline table

### 7-stage pipeline

Note the worker-number vs UI-stage-number offset: UI stage 3 is manual (no worker), so `StageNWorker` maps to UI stage `N+1` for N≥3 (e.g. `Stage5Worker` = UI stage 6, `Stage6Worker` = UI stage 7).

| Stage (UI) | Worker | What it does |
|---|---|---|
| 1 – Создание структуры | `Stage1Worker` | Async parallel download from Yandex Disk. Finds point folders (`п.NNNN`), downloads `Паспорт`, `Первичка`, `Стар. ЭПБ` + optional folders, copies templates. Uses `asyncio.gather` for all points simultaneously. |
| 2 – Копирование таблиц | `Stage2Worker` | Reads table #3 from `Акт замеров толщины стенки.docx` via `python-docx` (no COM for Word), writes it to `ПНОС_ТТП.xlsm` sheets `Импорт` and `Импорт (ПР)` via Excel COM. |
| 3 – Ручной контроль | `Manual3Worker` (`workers/manual3_worker.py`) | On-screen **form** (no manual Excel needed). Left: per-point list with status (✅ filled / ⬜ empty / ⏭ skipped / 🚫 invalid). Right: fields written to the `паспорт` sheet — B14 purpose (prefilled `транспортировки `), B22 specialist (dropdown, `config.SPECIALISTS`), B33 date, B52 org (dropdown, `config.EPB_ORGS`), B51 auto да/нет from B52, B53 reg-num, B54 EPB date, B55 sheet count. Required: B14/B22/B33. Dates written as **text** (`NumberFormat="@"`) so Excel keeps the exact `ДД.ММ.ГГГГ`. Drafts persist in `.pnos_status.json` (`manual3_draft`); "Записать всё в Excel" validates required fields for every non-skipped/non-invalid point, then `Manual3Worker` writes each xlsm via COM and sets `manual3`. "Пропустить пункт" sets `skipped`; "Открыть в Excel" opens the xlsm. |
| 4 – Запуск макросов | `Stage3Worker` | Reads Word passport DocVariables via Word COM, writes values to Excel `паспорт` sheet, runs `Restore` then `Save` VBA macros. Also reads column-A↔B mapping from master ПНОС file. Caches `паспорт!B12` (ГИ section number) into `.pnos_status.json` as `gi_number` for stage 7. |
| 5 – Создание отчёта | `Stage4Worker` | Runs `TransTab` macro, creates `CleanSource.xlsx` from `Экспорт` sheet, does MailMerge in Word, fills table from `Спецификация трубопровода.docx`, recalculates residual-life tables, replaces ЛНК certificate number based on act date. Saves the report as `ЗЭПБ № ТТП-<F9>.<YY> - ттп № <B17>.docx` (F9/B17 read from the `паспорт` sheet, year suffix from `config.REPORT_NAME`; falls back to `Отчет_<num>.docx` if either cell is empty). Later stages locate the report via `find_report_file()` in `stage4_worker.py`, which matches both the new `ЗЭПБ…` name and the legacy `Отчет_*` (and prefers the non-`ПЕЧАТЬ` copy). |
| 6 – Вставка картинок | `Stage5Worker` | Color-analyzes photos (red vs blue via Pillow), inserts rotated photos into `[ФОТО1]`/`[ФОТО2]` markers in the report, deletes empty/valve rows from the UT table, renumbers column 1, flags Г.3 table violations. |
| 7 – Вставка актов ГИ | `Stage6Worker` | Reads `gi_number` (from stage 3 cache, or openpyxl fallback on the xlsm). Downloads hydro-test act JPGs from `/ПНОС/<gi_number>/` on the templates account (`TOKEN_BASE`) into `<local_root>/_gi_cache/<num>/` (once per unique number). Inserts them one-per-page at the `[ГИ]` marker (sorted by trailing `_page-NNNN` number, page-duplication like `[ФОТО]`); each act is **EXIF-orientation-corrected** via `ImageOps.exif_transpose` (scans are stored sideways with an orientation tag Word ignores at `AddPicture`, so without it acts come out landscape) and sized/centered exactly like `[ФОТО]` (fit into `520×750` pt, `wdShapeCenter`). Then **fully recomputes the СОДЕРЖАНИЕ page numbers** from actual Word page numbers — TOC scanning starts only **after the `СОДЕРЖАНИЕ` heading** (title-page registration numbers / dates like `59-ЗС-…-2026` would otherwise be mistaken for section headings and abort the recompute), matching each TOC line to its body heading by section-number/appendix key. Missing act folder is not an error (distinct log messages for "no folder" vs "insert failed" vs partial `N/M`). |

### COM automation rules
- Stages 2, 4, 5, 6, 7 use `win32com.client` + `pythoncom.CoInitialize/CoUninitialize` in the QThread.
- `DialogKiller` / `WordDialogKiller` — daemon threads that auto-dismiss Excel/Word popups by posting `WM_KEYDOWN VK_RETURN` to `#32770` class windows.
- `safe_close_com(obj, save=False)` — always use this for cleanup; never call `obj.Close()` directly.
- During Stage 4 (macro run), user **must not** open other Excel files or steal mouse focus — the worker calls `xl_app.Run(...)` synchronously.
- On forced stop, `_force_stop_active_stage` calls `taskkill /F /IM EXCEL.EXE` and `WINWORD.EXE`.

### Self-update mechanism (`updater.py`)
In-app update: user clicks "Проверить обновления" in Settings → `UpdateWorker` (QThread) hits `releases/latest`, compares tags, downloads the asset named exactly `PNOS.zip` (case-insensitive), verifies its size, and extracts the first `.exe` inside to `PNOS_update.exe` next to the running EXE. Clicking "Установить и перезапустить" calls `apply_update_and_restart()`.

The app is a **onefile** PyInstaller EXE, typically run directly from the Desktop (`ПНОС.exe`), so the running file is locked by the OS. The swap uses a **rename-based** strategy (not delete-based), which is the key to reliability:
1. `apply_update_and_restart()` writes a PowerShell script to `%TEMP%\_pnos_updater.ps1` (**must be `utf-8-sig` / with BOM** — PS 5.1 otherwise reads it as cp1251 and the Cyrillic name `ПНОС.exe` becomes garbage), launches it **via WMI `Win32_Process.Create`** (`_spawn_outside_job`), then calls `os._exit(0)` (**not** `sys.exit(0)`, which a Qt slot can swallow → process stays alive and locks the file). **WMI launch is mandatory**: the PyInstaller-onefile bootloader keeps the app in a Windows Job Object with kill-on-close and breakaway forbidden, so a plain `subprocess.Popen` child (even `DETACHED_PROCESS`) inherits the Job and is killed the instant the app exits — the update silently never applies (symptom: `pnos_update.log` empty, `_pnos_updater.ps1` not self-deleted, no engine-start event in the "Windows PowerShell" event log). WMI makes WmiPrvSE the parent, outside the Job. Launch attempts are logged to `pnos_update.log` as `[py] launcher: ...`.
2. The script (`build_updater_ps1`) kills the process and waits for real exit, then **renames** the old `ПНОС.exe` → `ПНОС.exe.old` (works on NTFS even while the file is locked, unlike delete) and moves `PNOS_update.exe` onto the exact old path. The new file always ends up with the **same name/path as the current EXE** (taken from `sys.executable`, ignoring the exe name inside the ZIP).
3. On success it relaunches and deletes the `.old` backup; on failure it rolls back (`.old` → original) and removes the downloaded `PNOS_update.exe`. **Either way exactly one `ПНОС.exe` remains** — no orphan second icon.
4. On next startup `main()` calls `cleanup_update_leftovers()` to remove any leftover `<EXE>.old` / `PNOS_update.exe` (e.g. if AV briefly locked the backup).

Log: `%TEMP%\pnos_update.log`. Caveat: the updater code that runs is the one **already installed** on the user's machine — a fix to the swap logic only takes effect from the version that contains it onward (the *first* update to a new-logic version is still applied by the old code). In dev mode (not frozen) `apply_update_and_restart` just deletes the downloaded file and returns.

### Incremental execution
Each point folder stores a hidden `.pnos_status.json`. `get_stage_status(folder, "stage1")` returns `True` if already done → worker skips the point. `set_stage_status(folder, "stageN", True)` writes and re-hides the file (values may be any JSON type, not just bool — used for `gi_number`, `missing_folders`, `manual3_draft`). Delete this file to reprocess a point. Worker key ↔ UI step offset lives in `MainWindow.STEP_STATUS_KEYS` (used by the Overview matrix): step 3 = `manual3`, steps 4–7 = `stage3`–`stage6`.

### Invalid / skipped points (stage 1 gate)
Stage 1 checks each downloaded point for the required subfolders `config.REQUIRED_FOLDERS` (`Паспорт`, `Первичка` — `Стар. ЭПБ`/`Исполнительная`/`Акт отбраковки` are exempt). If any is missing, the folder is renamed to `п.NNNN (не корректный)` and `.pnos_status.json` gets `invalid: true` + `missing_folders`. `status_manager.is_point_blocked(folder)` returns `(True, reason)` for such folders (or user-`skipped` ones), and **every** worker (2–7) calls it at the top of its per-point loop to skip them. The Overview page and final summary surface these as problems.

### Configuration to update
- `config.py → TOKEN_ITEMS` / `TOKEN_BASE` — Yandex Disk OAuth tokens
- `config.py → DISK_PATHS_ITEMS` — list of remote paths containing `ТТП` subfolders
- `config.py → MACRO_MASTER_WB_PATH` — absolute path to master ПНОС summary Excel file (or override in Settings UI)
- `config.py → GI_CONFIG` — stage 7: `PASSPORT_CELL` (B12), `MARKER` (`[ГИ]`), `DISK_BASE` (`/ПНОС`, where `<number>/` act folders live)
- `config.py → REPORT_NAME` — final report filename: `YEAR_SUFFIX` (e.g. `26` → `.26`, bump yearly like `DISK_PATHS_ITEMS`), `SHEET` (`паспорт`), `TTP_CELL` (F9), `POINT_CELL` (B17)
- `config.py → REQUIRED_FOLDERS` / `INVALID_SUFFIX` — stage-1 validity gate (see "Invalid / skipped points")
- `config.py → MANUAL_FORM` (cell map for step 3) / `SPECIALISTS` (B22 list) / `EPB_ORGS` (B52 list)
- `version.py → APP_VERSION` — must be bumped before pushing a GitHub Release
