# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

ПНОС (Пакетный генератор отчётов) — a Windows desktop app that automates batch generation of industrial pipeline inspection reports (ЭПБ). It downloads folders from Yandex Disk, processes Word/Excel files via COM automation, runs macros, and generates final Word reports.

## Running and Building

The working interpreter with all dependencies installed is **`C:\Users\volod\AppData\Local\Python\bin\python.exe`** (Python 3.14). Beware: bare `python`/`py` is the WindowsApps stub, and `...\Programs\Python\Python313\python.exe` lacks the app's dependencies and has a broken lxml. Use the AppData\Local\Python interpreter for all runs and builds.

```powershell
$py = "C:\Users\volod\AppData\Local\Python\bin\python.exe"

# Install dependencies (no requirements.txt exists — install manually)
& $py -m pip install PySide6 aiohttp aiofiles python-docx openpyxl Pillow rapidfuzz pandas requests pywin32

# Run from source
& $py main.py

# Build a standalone EXE (output: dist/ПНОС.exe, onefile, ~90 MB)
& $py -m PyInstaller PNOS.spec --clean --noconfirm
```

```powershell
# Tests (pytest.ini: testpaths = repair/tests tests) — ~220 tests, ~80 s
& $py -m pytest
```

Tests cover the pure core of the «Ремонт» module (`repair/tests`) and worker
helpers (`tests`). Several of them run against **real files on this machine**
(`C:\Users\volod\Desktop\Сданы на ПНОС — копия`) and skip when those are absent —
if a run reports skips, check the `_CANDIDATES` paths in `repair/tests/test_map_labels.py`
and `repair/tests/test_pdf_preview.py` before assuming the tests are fine.

Everything outside those two folders (pipeline stages, GUI) has no automated
tests: verify by running the full GUI. Windows-only: COM automation depends on
installed Microsoft Word and Excel, and OCR on the Windows Russian language pack.

## Commits and releases

- Commit messages and release notes are written in **Russian**, and must **not** contain a `Co-Authored-By` trailer (explicit user requirement).
- Repo is `azoga99/PNOS`; `gh` is authenticated as azoga99.
- **Releasing a new version is a strict procedure — always follow the `release-pnos` skill**, do not improvise. Summary: bump `version.py → APP_VERSION` → commit + `git push origin main` → build EXE via PyInstaller → `Compress-Archive dist\ПНОС.exe -DestinationPath dist\PNOS.zip` (exactly one `.exe` inside) → `gh release create vX.Y.Z dist/PNOS.zip --repo azoga99/PNOS --title vX.Y.Z --latest --notes "..."`. The in-app updater hits `releases/latest` and requires an asset named exactly `PNOS.zip` (case-insensitive) and a tag that parses higher than users' current version; draft/prerelease releases are invisible to it.

## Architecture

### Entry flow
`main.py` (shows `SplashLoader` — program icon with a spinning arc — while the window builds) → `MainWindow` (`main_window.py`, PySide6 QMainWindow) → `QStackedWidget` with 6 pages: **Процесс** (pipeline of 7 `StepCard`s + launch params + single log), **Ручной контроль** (step-3 form + embedded «Ремонт» module, see below), **Обзор и отчёт** (one page, two tabs: the Обзор status matrix and the Отчёт о прогоне per-step result tables built from `_stage_results`), **Правка отчётов** (standalone batch fixer for finished reports, see below), **Анализ пунктов** (standalone readiness statistics, see below), **Настройки**. Workers live in `workers/*.py` (QThreads).

Page indices live in `MainWindow.PAGE_*` (`PAGE_SETTINGS = 5`). There is deliberately **no separate «Ремонт» nav item**: the module is opened from the manual-control form for a point that has repair documents.

### Design rules
- Strict business look: **no emojis** in UI text (the only exception is 📂 on folder/file pickers), no gradients; colors only via QSS tokens.
- Statuses are shown with colored dots (`widgets.dot_icon`) and `StepCard` badges (`STATUS_COLORS`: idle/run/done/error/wait), not glyph emojis.
- Per-step result tables/metrics bars were deliberately removed from step UI — results go to the Обзор matrix and the Итоговый отчёт page only; each step card shows a one-line summary.

### Startup performance — lazy imports (do not break)
Heavy modules (pandas, aiohttp, win32com, python-docx) are deliberately **not** imported at module top level in `main.py` / `main_window.py`. Workers and `report_dialog` are imported lazily inside the stage-launch methods, and `main()` prewarms them in a background daemon thread after the window is shown (`_prewarm_heavy_imports`). Adding a top-level import of a worker or heavy library to these files silently reverts the fast-startup optimization.

### Key files
| File | Purpose |
|---|---|
| `main_window.py` | All pages: the Процесс pipeline (`_setup_process_page`, step cards wired to `_start_stage*`), manual-control form (`_setup_manual_page`), Обзор и отчёт tabs (`_setup_overview_page`/`_refresh_overview`/`_refresh_report_page`), Правка отчётов (`_setup_fixreports_page`/`_start_fix_reports`), single activity log (`_add_activity`), auto-mode chaining in `_on_step_finished` (Settings checkbox `chk_auto`; stops at manual step 3 with a "wait" card), and `_force_stop_active_stage` (emergency stop) |
| `widgets.py` | `StepCard` (pipeline card: status badge, summary line, action/stop buttons, thin progress bar, hover shadow animation), `dot_icon(color)` (colored-dot QIcon used instead of emojis in lists), `SegmentBar` (one line per phase — used where a percentage would lie because the total work is not known upfront), `DateCellDelegate` + `date_cell_value` (table cell with a hard ДД.ММ.ГГГГ input mask) |
| `theme.py` | Light/dark theme system. `style.qss` is a token template (`@name@`); `build_stylesheet(theme)` substitutes a palette dict (`LIGHT`/`DARK`) into it. Theme choice persists in `QSettings("PNOS","PNOS")` key `theme`. Never hard-code colors in QSS — add a token to both palettes. |
| `pages/settings_page.py` | Settings page — auto-mode checkbox, theme toggle (emits `theme_changed`), master-file path override (session-only; last non-empty pick saved to `QSettings` key `master_excel` as a dialog prefill hint), download speed toggle, update check & install buttons |
| `config.py` | Single source of truth — Yandex Disk OAuth tokens, remote paths, folder names, macro paths |
| `version.py` | `APP_VERSION` string — bump before every release |
| `status_manager.py` | Per-point incremental tracking via hidden `.pnos_status.json` in each point folder |
| `yandex_api.py` | Fully async Yandex Disk REST client (aiohttp + aiofiles, semaphore-throttled) |
| `excel_service.py` | Reads the master Excel schedule: `analyze_excel` filters rows by columns E/F/H for stage 1; `read_analytics_rows` reads the wide schedule (columns up to RE via `usecols`) for the Анализ пунктов page and computes the approval-status ladder |
| `updater.py` | Checks GitHub Releases (`azoga99/PNOS`), downloads `PNOS.zip`, extracts the `.exe`, then swaps it in via a generated PowerShell script and relaunches. See "Self-update mechanism" below. |
| `workers/fix_reports_worker.py` | `FixReportsWorker` — Правка отчётов page (see "Правка отчётов" below) |
| `workers/analytics_worker.py` | `ScheduleLoadWorker` (background Excel read), `AnalyticsWorker` (async read-only Yandex Disk scan), `export_analysis_word` (python-docx report) — Анализ пунктов page (see below) |
| `pages/repair_page.py` | `RepairPage` — the whole «Ремонт» UI (four tabs, embedded into Ручной контроль via `set_embedded(True)`), plus `MapView` (zoom/pan corrosion-map viewer) |
| `workers/repair_worker.py` | `RepairScanWorker` (read-only point scan), `RepairOcrWorker` (OCR of one document, phase signals), `MapLabelsWorker` (section labels on all map sheets), `RepairApplyWorker` (backup → COM write → journal), `find_point_xlsm` / `find_point_report` |
| `workers/win_dialogs.py` | Diagnostics only: `dialog_buttons`, `describe_dialogs` — used by the stage-5 merge watchdog to log which Office windows are hanging. **Does not press anything** (see "COM automation rules") |
| `repair/*` | Pure core of the «Ремонт» module — no Qt, no COM (see "Модуль «Ремонт»" below) |
| `utils.py` | `resource_path()` (resolves paths for both dev and PyInstaller frozen EXE), `title_date_text()` («06.07.2026» → «06 июля 2026 г.»), `find_order_scan()` (newest JPG/PNG in a folder) |

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
| 3 – Ручной контроль | `Manual3Worker` (`workers/manual3_worker.py`) | On-screen **form** (no manual Excel needed) + the «Ремонт» module for points that have repair documents. Left: per-point list with status (✅ filled / ⬜ empty / ⏭ skipped / 🚫 invalid) and a `• ремонт` / `• ремонт (вручную)` mark. Right: fields written to the `паспорт` sheet — B14 purpose (prefilled `транспортировки `), B22 specialist (dropdown, `config.SPECIALISTS`), B33 date, B52 org (dropdown, `config.EPB_ORGS`), B51 auto да/нет from B52, B53 reg-num, B54 EPB date, B55 sheet count. Required: B14/B22/B33. Dates written as **text** (`NumberFormat="@"`) so Excel keeps the exact `ДД.ММ.ГГГГ`. Drafts persist in `.pnos_status.json` (`manual3_draft`); "Записать всё в Excel" validates required fields for every non-skipped/non-invalid point, then `Manual3Worker` writes each xlsm via COM and sets `manual3`. "Пропустить пункт" sets `skipped`; "Открыть в Excel" opens the xlsm. |
| 4 – Запуск макросов | `Stage3Worker` | Reads Word passport DocVariables via Word COM, writes values to Excel `паспорт` sheet, runs `Restore` then `Save` VBA macros. Also reads column-A↔B mapping from master ПНОС file. Caches `паспорт!B12` (ГИ section number) into `.pnos_status.json` as `gi_number` for stage 7. |
| 5 – Создание отчёта | `Stage4Worker` | Runs `TransTab` macro, creates `CleanSource.xlsx` from `Экспорт` sheet, does MailMerge in Word, fills table from `Спецификация трубопровода.docx`, recalculates residual-life tables, replaces ЛНК certificate number based on act date. Saves the report as `ЗЭПБ № ТТП-<F9>.<YY> - ттп № <B17>.docx` (F9/B17 read from the `паспорт` sheet, year suffix from `config.REPORT_NAME`; falls back to `Отчет_<num>.docx` if either cell is empty). Later stages locate the report via `find_report_file()` in `stage4_worker.py`, which matches both the new `ЗЭПБ…` name and the legacy `Отчет_*` (and prefers the non-`ПЕЧАТЬ` copy). |
| 6 – Вставка картинок | `Stage5Worker` | Color-analyzes photos (red vs blue via Pillow), inserts rotated photos into `[ФОТО1]`/`[ФОТО2]` markers in the report, deletes empty/valve rows from the UT table, renumbers column 1, flags Г.3 table violations. |
| 7 – Вставка актов ГИ | `Stage6Worker` | Reads `gi_number` (from stage 3 cache, or openpyxl fallback on the xlsm). Downloads hydro-test act JPGs from `/ПНОС/<gi_number>/` on the templates account (`TOKEN_BASE`) into `<local_root>/_gi_cache/<num>/` (once per unique number). Inserts them one-per-page at the `[ГИ]` marker (sorted by trailing `_page-NNNN` number, page-duplication like `[ФОТО]`); each act is **EXIF-orientation-corrected** via `ImageOps.exif_transpose` (scans are stored sideways with an orientation tag Word ignores at `AddPicture`, so without it acts come out landscape) and sized/centered exactly like `[ФОТО]` (fit into `520×750` pt, `wdShapeCenter`). Then **fully recomputes the СОДЕРЖАНИЕ page numbers** from actual Word page numbers — TOC scanning starts only **after the `СОДЕРЖАНИЕ` heading** (title-page registration numbers / dates like `59-ЗС-…-2026` would otherwise be mistaken for section headings and abort the recompute), matching each TOC line to its body heading by section-number/appendix key. Missing act folder is not an error (distinct log messages for "no folder" vs "insert failed" vs partial `N/M`). |

### Модуль «Ремонт» (v2.0.0+)

Учёт ремонтов трубопровода: акты отбраковки и исполнительная документация →
лист «Ремонт» книги пункта, номиналы заменённых сечений на листах «Импорт»,
строка «Сведения о ремонте» в таблице 1 отчёта, значения на коррозионной карте.

**Где живёт.** `RepairPage` встроена в «Ручной контроль» (`set_embedded(True)`
прячет заголовок и панель выбора пункта). Кнопка «Внести данные по ремонту» в
баннере появляется, если `repair.detect.find_repair_docs(folder).has_repair`.
Рядом «Оформить по старому» — ставит `repair_manual` в `.pnos_status.json`,
после чего модуль пункт не трогает (флаг снимается той же кнопкой).

**Принципы (нарушать нельзя).**
- **П-1**: всё хранится в обычных файлах пункта — лист «Ремонт» видимый, его
  можно править руками в Excel, модуль подхватит.
- **П-2**: ничего не пишется без предпросмотра и явной кнопки.
- **П-4**: перед каждой записью — бэкап (`repair/safety.py`), после — запись в
  `_журнал_ремонта.txt`; ошибка записи → автооткат (`repair/orchestrator.py`).

**Слои.** `repair/*` — чистое ядро без Qt и COM, покрытое тестами:

| Модуль | Отвечает за |
|---|---|
| `detect.py` | поиск документов ремонта и разбор их имён (установка, линия, номера актов, дата) |
| `ocr.py` | Windows.Media.Ocr через pywinrt: `ocr_page` с автоподбором поворота, `ocr_zone_multi` (несколько масштабов — мелкие ячейки читаются нестабильно), координаты слов |
| `id_map.py` | адресная навигация по комплекту ИД: разбор реестра, правило Н1 (смещение), поиск копии акта/удостоверения/таблицы замены по заголовкам |
| `dates.py` | правило Д1 (самая поздняя дата в блоке подписей) и Д2 (гриф → имя файла → шапка реестра) |
| `doc_extract.py` | извлечение полей из распознанного текста |
| `registry.py` | лист «Ремонт» (колонки A…I), сериализация записей |
| `imports.py` | чтение сечений листа «Импорт» и диф записи. **Значения пишутся текстом** с `NumberFormat="@"`: блок отчёта (AH…AS) переносит значение как есть, и число 6.0 печаталось бы как «6» |
| `nominals.py` | `SectionEdit` (Тисп / диаметр / замер по отдельности) и `apply_edits` |
| `svedenia.py` | сборка строки «Сведения о ремонте» по шаблонам `repair/config.py` |
| `map_labels.py` | поиск подписей «NN (V.V)» на карте по координатам слов OCR |
| `map_sheets.py` | какие PNG «Первички» править: рабочая карта — без точки в имени (`2704_1.png`), исходник ОТН — с точкой (`2704._1.png`); `likely_sheet` подсказывает лист по соседним номерам |
| `map_editor.py` | стирание старого значения и отрисовка нового. Стирание идёт **по цветности**, а не по яркости — бледная кайма сглаживания иначе остаётся розовыми точками |
| `suggest.py` | подбор сечений по таблице замены (обычно неоднозначный — отдаёт список кандидатов) |
| `verifier.py` | независимая проверка пункта R1…R5; принципиально не спрашивает, кто делал пункт |
| `com_writer.py` | запись в книгу через Excel COM |
| `word_writer.py` | чтение и запись строки «Сведения о ремонте» через python-docx |

**Поток работы.** Открытие пункта → `RepairScanWorker` (только чтение) →
**очередь распознавания по всем документам сразу**, без нажатий
(`_start_ocr_queue`): у пункта обычно два документа, и кроп даты даёт только
комплект ИД, который в списке не первый. Прогресс — `SegmentBar` по этапам всей
очереди. В конце очереди показывается кроп блока подписей и курсор встаёт в
ячейку «Дата выполнения» (маска ДД.ММ.ГГГГ). Нечитаемый PDF пропускается с
записью в журнал, очередь не рвётся.

**Вкладки идут в порядке работы**: Документы и реестр → Сечения → Карта →
Предпросмотр. Карта сохраняется **своей** кнопкой и от «Применить выбранное» не
зависит; состояние карты показывается в плашке предпросмотра.

**Карта из нескольких листов.** У 9 из 34 сданных пунктов карта разбита на 2–3
листа. Номера сечений между листами практически не пересекаются, поэтому лист
определяется по факту находки подписи. Страховка: лист, где распознано меньше
`MIN_SECTION_LABELS` подписей, к правке не допускается (у исходников ОТН их 0–1,
у рабочих листов — от 66). **Синие подписи внутри рабочей карты отсеивать
нельзя** — часть настоящих сечений подписана синим (п.2763 — 26 из 107).

**Отчёта на шаге 3 ещё нет** — его создаёт этап 5. Поэтому строка «Word ·
таблица 1» в предпросмотре снята с галочки и помечена «(отчёта ещё нет)», а
плашка после записи честно говорит, что строка вносится позже. Сотрудник
возвращается в модуль после этапа 5 и жмёт «Применить выбранное» второй раз.

### Правка отчётов (standalone page, v1.7.0+)
`FixReportsWorker` (`workers/fix_reports_worker.py`) batch-edits **finished** Word reports via Word COM — it is autonomous: not tied to the pipeline, `.pnos_status.json`, or the Обзор pages. It processes every report `.docx` in the chosen folder and in the root of each `п.*` subfolder (including `ПЕЧАТЬ` copies, skipping the `ПНОС_ТТП.docx` template) and applies four fixes driven by `config.FIX_REPORTS`: title date (MERGEFIELD `ДатаОтчета`, rendered as words via `title_date_text`), expert qualification-certificate lines (`CERT_FIXES`, keyed by OLD_KEY/DONE_KEY so re-runs are idempotent), the production-control regulation cell in Таблица 1 (`TABLE_FIX`), and the order-scan image in appendix Б/В (`ORDER_HEADING`; replacement image keeps the original's dimensions, sourced via `find_order_scan`).

### Анализ пунктов (standalone page, v1.8.0+)
Readiness statistics for a specialist's assigned points — read-only, autonomous (no pipeline, no `.pnos_status.json`, nothing downloaded or written). The page has two tabs («Выбор пунктов» / «Результаты», auto-switches to results after a scan) and **no file picker**: the schedule is the master file already configured for stage 4 (`MainWindow._master_excel_path()`: session-only Settings override `edit_macro_path`, else server `MACRO_MASTER_WB_PATH` — the Settings field is deliberately NOT restored across launches, so a fresh launch always prefers the server copy for data freshness) and is auto-read on startup (`_an_auto_load` in `_switch_page` is a fallback; «Перечитать график» re-reads). `MasterCheckWorker` probes the path in a background thread (a dead network drive can block for tens of seconds); if the server is unreachable, `MasterScheduleDialog` opens on **every** launch — the user explicitly confirms which copy to use — prefilled with the last manual pick (`QSettings` key `master_excel`, a hint only, never auto-used) and showing the file's mtime. On successful read a 4-second toast (`_show_toast`, `QLabel#toast` styled via style.qss tokens) reports «подтянут с сервера» or «из файла (изменён <date>)». Flow: `ScheduleLoadWorker` reads the schedule in the background (`read_analytics_rows`: rows from `ANALYTICS.DATA_START_ROW`, always excluding rows where column G is non-empty = "deleted" points); the UI then filters client-side by specialist (column MF, `ANALYTICS.SPECIALISTS`), equipment type (column E), approval status, and a point-number search box. Approval status ladder from date columns (a cell counts only if it is a real date or a `ДД.ММ.ГГГГ` string): QZ = оформлен → +RB = отправлен Шарову → +RC = согласован Шаровым → +RE = отправлен Жиганьшину. The user checks up to `MAX_POINTS` (30) points (checkboxes, rubber-band select + «Отметить выделенные», or «Все»), then `AnalyticsWorker` scans them on Yandex Disk in parallel (`find_point_folder`, then fuzzy `find_folder_by_name` + `get_folder_contents` for Паспорт/Первичка/Стар. ЭПБ/Исполнительная/Акт отбраковки). Verdicts: ready (both `REQUIRED_FOLDERS` non-empty) / missing_critical / empty / not_found. Results go to a colored table + chips and can be exported to a Word summary via `export_analysis_word` (python-docx, no COM). Status labels are duplicated in `main_window.AN_STATUS_LABELS` deliberately — importing `excel_service` at module level would pull pandas and break lazy startup.

### COM automation rules
- Stages 2, 4, 5, 6, 7 use `win32com.client` + `pythoncom.CoInitialize/CoUninitialize` in the QThread.
- Paths coming from `QFileDialog` must go through `os.path.normpath` before Word COM `Documents.Open` — forward slashes break it (there is also an 8.3 short-name fallback for stubborn paths).
- Wildcard search in Word via pywin32: arguments passed to `Find.Execute(FindText=..., MatchWildcards=...)` are silently ignored — set everything through `Find` **properties** first, then call `Execute()` (see `range_find` in `fix_reports_worker.py`).
- `DialogKiller` (`stage3_worker.py`) / `WordDialogKiller` (`stage4_worker.py`) — daemon threads that auto-dismiss Excel/Word popups: `WM_KEYDOWN VK_RETURN` plus `SetForegroundWindow` + `WScript.Shell.SendKeys("~")`.
  - **This requires the ПНОС window to be the foreground app** — Windows forbids `SetForegroundWindow` to a background process. Consequence: during stage 5 the Word SQL-command confirmation at `MailMerge.OpenDataSource` hangs until the user clicks the ПНОС window. This is a **known, accepted trade-off**: closing dialogs by posting `BM_CLICK` to the button was tried (works without focus) and **reverted** — it silently lost Office windows whose buttons are custom-drawn controls with no `Button` class, and the Excel dialog stopped being dismissed at all. `tests/test_win_dialogs.py` guards against reintroducing it.
  - `Stage4Worker._merge_watchdog` logs which Office windows are on screen if `OpenDataSource` takes longer than 20 s (`workers/win_dialogs.describe_dialogs`) — diagnostics only.
- `safe_close_com(obj, save=False)` — always use this for cleanup; never call `obj.Close()` directly.
- During Stage 4 (macro run), user **must not** open other Excel files or steal mouse focus — the worker calls `xl_app.Run(...)` synchronously.
- `find_point_xlsm` (`workers/repair_worker.py`) must skip `CleanSource.xlsx`: stage 5 creates it in the point folder, and it sorted first alphabetically (Latin before Cyrillic), so the repair module read sections from the temporary export instead of `ПНОС_ТТП.xlsm`.
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
Each point folder stores a hidden `.pnos_status.json`. `get_stage_status(folder, "stage1")` returns `True` if already done → worker skips the point. `set_stage_status(folder, "stageN", True)` writes and re-hides the file (values may be any JSON type, not just bool — used for `gi_number`, `missing_folders`, `manual3_draft`, `repair_manual`, `repair_findings`). Delete this file to reprocess a point. Worker key ↔ UI step offset lives in `MainWindow.STEP_STATUS_KEYS` (used by the Overview matrix): step 3 = `manual3`, steps 4–7 = `stage3`–`stage6`.

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
- `config.py → FIX_REPORTS` — Правка отчётов page: `DATE_FIELD`, `CERT_FIXES` (certificate numbers change when experts recertify), `TABLE_FIX` (regulation number/date, bump when a new положение is issued), `ORDER_HEADING`
- `config.py → ANALYTICS` — Анализ пунктов page: `SPECIALISTS` (surname list), `DATA_START_ROW`, schedule column letters (`COL_PERSON` MF, `COL_DELETED` G, status columns QZ/RB/RC/RE), `MAX_POINTS`
- `repair/config.py → REPAIR_CONFIG` — модуль «Ремонт»: шаблоны строки «Сведения о ремонте» (`SV`), `SMALL_ELEMENTS` / `EXCLUDE_FROM_TABLE3` (задвижки и бобышки не обновляются), пороги
- `version.py → APP_VERSION` — must be bumped before pushing a GitHub Release (see "Commits and releases")

### Documentation for users
`ИНСТРУКЦИЯ_Ремонт.md` — одностраничная инструкция для сотрудников (от ввода
номера пункта до готового отчёта). Держать её короткой: это не справочник, а
шпаргалка по кнопкам. Подписи кнопок в ней должны совпадать с интерфейсом —
сверяйте по коду перед правкой.
