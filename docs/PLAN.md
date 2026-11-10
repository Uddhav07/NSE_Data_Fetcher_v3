# Plan — robustness, cleanup and structure (drafted 10 Nov 2026, against v3.8.3)

Goal: same behaviour and the same workbook, but sturdier, with no dead code, and organised so each file has one job.
Nothing here changes the Excel contract (columns A–W, fill-only-empty, cell-only colours + Tuesday row, ledger).

## Method (why this order)
1. **Safety net first.** Add characterization ("golden") tests that lock in today's exact output *before* touching code.
   Every later step must leave them green. That is what makes a restructure safe.
2. **Delete before you move.** Removing dead code first means less to restructure.
3. **Fix behaviour in small, separately tested commits.** Never mix a behaviour change with a file move.
4. **Restructure as pure moves.** No logic edits, golden tests unchanged.
5. **Docs last,** once the code has settled, and one release per phase.

---

## Phase 0 — Safety net (no production code changes) · v3.8.4 ✅ done
| # | Task | Why |
|---|---|---|
| 0.1 | **Golden workbook test:** a fixed OHLC/RSI/futures input → `update_workbook` → compare every cell's value, formula, number format, fill and CF rule against a stored JSON snapshot | Catches any accidental change during refactor |
| 0.2 | **`run_fetch` test** with `fetch_ohlc_data` / `fetch_futures` / `fetch_rsi` / `session_in_progress` mocked (live + after-close cases, `--full-refresh`, futures failure must not block OHLC) | `CLAUDE.md` demands this but no such test exists. The orchestrator is untested |
| 0.3 | Unit tests: `wilder_rsi` against a known series, `_to_45m` buckets, nakshatra for 3 reference dates, `session_in_progress` at 15:34/15:35/Saturday | Pure functions with no tests today |
| 0.4 | Add `pyproject.toml`: pytest + **ruff** config (lint + import sort), replacing ad-hoc pyflakes | One command for lint |
| 0.5 | **CI (GitHub Actions)**: pytest on ubuntu, macos **and windows**, Python 3.9 + 3.12 | Only cheap way to test the Windows path we've never run |

## Phase 1 — Remove what exists for no reason · v3.8.5
| # | Remove | Evidence |
|---|---|---|
| 1.1 | `main._check_environment()` | Package check duplicates `bootstrap`. Python-version check runs after 3.9-only code has already imported. Network `HEAD` to Yahoo adds up to 5 s every run, and fetch already retries |
| 1.2 | GUI remnants: `setup_logging(extra_handler=…)`, "used by CLI and GUI" docstrings, `paths.get_config_path` "exe" docstring | The GUI was removed in v3.2 |
| 1.3 | `config_manager.KNOWN_TICKERS` | Never used |
| 1.4 | `excel_writer.get_last_date` + the `min(start, last+1)` dance in `run_fetch` | Result is always `start_date`, because we re-check from start every run. Its only effect is an extra workbook load and a log line |
| 1.5 | Second `save_default_config` + `load_config` in `run()` (config is read twice) | Pass the loaded config into `run_fetch` |
| 1.6 | `_Guard` class hack in `_add_conditional_formatting` | A 3-line helper function does the same |
| 1.7 | Defaults written 3 times (dataclass, dict in `save_default_config`, `config.json`) | Generate the file from `AppConfig` (`dataclasses.asdict`) |
| 1.8 | Stale comments: "RSI (R:U)", "Date Total (Q)", `A1:V` in `_apply_table`, `--no-rsi` help "R-U", dead width check in `_ensure_headers` (`letter not in …` after the lookup created it) | Misleading |
| 1.9 | `CLAUDE.md` "Scope: No Windows… `.bat`" | Contradicts reality: the user runs Windows and `.bat` launchers ship |

**Keep (they look removable but aren't):** `bootstrap.py`, `.command` wrappers (a `.sh` double-click opens an editor), all `.bat` files, legacy CF formulas in `_ours` (needed to clean up old sheets), A→B→C migrations, ledger, `scripts/` package name. The name stays because the `.bat` launchers already on the Windows machine run `-m scripts.main`; a rename would break them mid-update.

## Phase 2 — Robustness fixes · v3.9.0
| # | Fix | Problem today |
|---|---|---|
| 2.1 | **User config out of git:** track `config/config.example.json`, git-ignore `config/config.json`, create it on first run (migrate an existing one untouched) | ⚠️ **Real bug:** `config.json` is tracked, and `update_and_run` skips the update whenever a tracked file has changed. Editing any setting silently stops all future updates |
| 2.2 | **Atomic save:** write to `*.tmp` in the same folder, then `os.replace` | A crash or power loss mid-`wb.save` corrupts the workbook (only the backup saves you) |
| 2.3 | **Excel-open detection:** if `~$<name>.xlsx` exists, stop with "close Excel first" | macOS Excel doesn't lock the file. We write under it, and Excel later overwrites our changes or shows a conflict |
| 2.4 | **Backup retention:** keep the newest N (config `backup_keep`, default 30); add a counter to the timestamped name | `archive/` grows by one copy every run, forever (18 files already). Two backups in the same second overwrite each other (found while writing the Phase 0 tests) |
| 2.5 | **One shared `retry()` helper** (exponential backoff, logs) used by OHLC, RSI, futures | Three copy-pasted loops with slightly different behaviour |
| 2.6 | **Validate the futures scrape:** price must be a number within ±5 % of the latest close; else drop it with a warning | Groww's structure can change and silently write garbage into U/M |
| 2.7 | Warn if `ticker` ≠ `^NSEI` while futures are on | Futures are always Nifty, so U/M would be wrong for Bank Nifty |
| 2.8 | Config: reject `bool` for int fields (`True` is accepted as `max_retries=1` today) | Small type hole |
| 2.9 | `requirements.txt`: tested upper bounds (e.g. `yfinance>=0.2.40,<2`, `pandas>=2,<3`, `openpyxl>=3.1,<3.2`) | Unbounded `>=` means a future major release can break a fresh install |
| 2.10 | Single workbook load per run, plus a log summary of what was filled/refreshed per column | Easier support from a log alone |

## Phase 3 — Restructure (pure moves, golden tests unchanged) · v3.10.0
`excel_writer.py` (700 lines, magic column numbers everywhere) is the main target. Proposed layout:
```
scripts/
  main.py            stdlib-only entry: bootstrap → argparse → pipeline.run_fetch   (keep: launchers call it)
  bootstrap.py       unchanged
  pipeline.py        run_fetch() + FetchResult (dataclass); orchestration only
  core/              config.py · paths.py · logging_setup.py · market.py (IST session) · retry.py
  sources/           yahoo.py (daily OHLC + intraday closes) · groww.py (futures scrape + validation)
  indicators/        rsi.py (pure math: wilder, 45m, last-per-day) · astro.py (moon_calc + nakshatra merged)
  workbook/
    layout.py        COLUMNS + a named column registry (Col.CLOSE, Col.RSI[5] … instead of 3, 6, 15)
    styles.py        fills, fonts, borders
    migrations.py    _migrate_layout + legacy CF formula list
    formulas.py      canonical formulas, stale-formula repair
    formatting.py    conditional formats, table, freeze/scroll
    ledger.py        refresh ledger (state.json)
    io.py            load / atomic save / backup + retention / lock detection
    writer.py        update_workbook() — the ordered steps, nothing else
tests/               split by area: test_workbook.py · test_migrations.py · test_indicators.py · test_pipeline.py · golden/
```
Rules: one commit per move, the golden test passes after each, and import paths are updated in the same commit.
Also: move `ensure_dependencies()` out of import time (call it from `main()` before importing `pipeline`), so tests and tools can import modules without side effects.

## Phase 4 — Docs and tidy · ship with v3.10.0
- Merge this plan's leftovers and HANDOVER §12 into one "Open items" list. Drop `docs/ARCHITECTURE.md` "Issues found in the original" (history belongs in `CHANGELOG.md`).
- Rewrite `ARCHITECTURE.md` against the new layout. Update the HANDOVER map, `CLAUDE.md` rules and the README config section (`config.example.json`).
- Optional, outside the repo: delete the duplicate `~/Music/NSE_Data_Fetcher_v3_repo` clone.

## Verification for every phase
`pytest` + `ruff` green locally and in CI (3 OSes). Then run old and new code on a **copy** of the real workbook with the same mocked data, and diff all cells. The only differences allowed are the ones the phase intends.

## Later ideas (not part of this plan)
Backfill older RSI from another intraday source · sturdier futures source than the Groww scrape · `launchd`/Task Scheduler 15:40 IST run · tithi / panchang columns · other tickers.
