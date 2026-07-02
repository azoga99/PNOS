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
`main.py` → `MainWindow` (PySide6 QMainWindow) → `QStackedWidget` with pages for each step → `workers/stage*_worker.py` (QThreads).

### Key files
| File | Purpose |
|---|---|
| `config.py` | Single source of truth — Yandex Disk OAuth tokens, remote paths, folder names, macro paths |
| `version.py` | `APP_VERSION` string — bump before every release |
| `status_manager.py` | Per-point incremental tracking via hidden `.pnos_status.json` in each point folder |
| `yandex_api.py` | Fully async Yandex Disk REST client (aiohttp + aiofiles, semaphore-throttled) |
| `excel_service.py` | Reads the master Excel schedule and filters rows by columns E/F/H |
| `updater.py` | Checks GitHub Releases (`azoga99/PNOS`), downloads ZIP, extracts EXE, restarts via `.bat` |
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
| 3 – Ручной контроль | *(no worker)* | Manual user step; just marks itself done when the button is clicked. |
| 4 – Запуск макросов | `Stage3Worker` | Reads Word passport DocVariables via Word COM, writes values to Excel `паспорт` sheet, runs `Restore` then `Save` VBA macros. Also reads column-A↔B mapping from master ПНОС file. Caches `паспорт!B12` (ГИ section number) into `.pnos_status.json` as `gi_number` for stage 7. |
| 5 – Создание отчёта | `Stage4Worker` | Runs `TransTab` macro, creates `CleanSource.xlsx` from `Экспорт` sheet, does MailMerge in Word, fills table from `Спецификация трубопровода.docx`, recalculates residual-life tables, replaces ЛНК certificate number based on act date. |
| 6 – Вставка картинок | `Stage5Worker` | Color-analyzes photos (red vs blue via Pillow), inserts rotated photos into `[ФОТО1]`/`[ФОТО2]` markers in the report, deletes empty/valve rows from the UT table, renumbers column 1, flags Г.3 table violations. |
| 7 – Вставка актов ГИ | `Stage6Worker` | Reads `gi_number` (from stage 3 cache, or openpyxl fallback on the xlsm). Downloads hydro-test act JPGs from `/ПНОС/<gi_number>/` on the templates account (`TOKEN_BASE`) into `<local_root>/_gi_cache/<num>/` (once per unique number). Inserts them one-per-page at the `[ГИ]` marker (sorted by trailing `_page-NNNN` number, no rotation, page-duplication like `[ФОТО]`). Then **fully recomputes the СОДЕРЖАНИЕ page numbers** from actual Word page numbers — matching each TOC line to its body heading by section-number/appendix key. Missing act folder is not an error. |

### COM automation rules
- Stages 2, 4, 5, 6, 7 use `win32com.client` + `pythoncom.CoInitialize/CoUninitialize` in the QThread.
- `DialogKiller` / `WordDialogKiller` — daemon threads that auto-dismiss Excel/Word popups by posting `WM_KEYDOWN VK_RETURN` to `#32770` class windows.
- `safe_close_com(obj, save=False)` — always use this for cleanup; never call `obj.Close()` directly.
- During Stage 4 (macro run), user **must not** open other Excel files or steal mouse focus — the worker calls `xl_app.Run(...)` synchronously.
- On forced stop, `_force_stop_active_stage` calls `taskkill /F /IM EXCEL.EXE` and `WINWORD.EXE`.

### Incremental execution
Each point folder stores a hidden `.pnos_status.json`. `get_stage_status(folder, "stage1")` returns `True` if already done → worker skips the point. `set_stage_status(folder, "stageN", True)` writes and re-hides the file. Delete this file to reprocess a point.

### Configuration to update
- `config.py → TOKEN_ITEMS` / `TOKEN_BASE` — Yandex Disk OAuth tokens
- `config.py → DISK_PATHS_ITEMS` — list of remote paths containing `ТТП` subfolders
- `config.py → MACRO_MASTER_WB_PATH` — absolute path to master ПНОС summary Excel file (or override in Settings UI)
- `config.py → GI_CONFIG` — stage 7: `PASSPORT_CELL` (B12), `MARKER` (`[ГИ]`), `DISK_BASE` (`/ПНОС`, where `<number>/` act folders live)
- `version.py → APP_VERSION` — must be bumped before pushing a GitHub Release
