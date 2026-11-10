# Project Handover — NSE Data Fetcher

Written for release **v3.8.3** (10 Nov 2026). Purpose: let a new person, or a new Claude session, take over without the original chat.
Read order for a newcomer: **this file → `PROJECT.md` (why) → `README.md` (use) → `docs/ARCHITECTURE.md` (internals) → `CLAUDE.md` (editing rules)**.

---

## 1. What this project is, in one paragraph
A Python script that keeps a daily **Nifty 50 journal in one Excel file** (`data/stock_market_NSE50.xlsx`, sheet `NSE_Data`, an Excel Table `NSE_Table`, columns A–W). Each run adds missing trading days (Yahoo Finance OHLC), formulas (daily/weekly change, High/Low breaks, HH/HL, LL/LH), the current Nifty futures price (Groww scrape), RSI(14) at the 15:30 IST close for 5/15/30/45/60-minute candles, a "Date Total" (e.g. `9, 3`) and the Moon's nakshatra at 09:15 IST. The sheet is the user's working document: the script only **adds rows and fills empty cells**; the user's edits and notes survive.

## 2. Who, where, and what environment
- Owner: GitHub user **Uddhav07**, repo `github.com/Uddhav07/NSE_Data_Fetcher_v3` (default branch `main`, tags `v3.1.0` … `v3.8.3`).
- Dev machine: **macOS, system Python 3.9.6** (no brew Python) — hence the code supports **Python 3.9+** (`from __future__ import annotations`).
- User also runs it on **Windows** (they reported the Nakshatra column blank there → fixed in v3.4.0). Windows scripts and Excel rendering have **never been tested on real Windows/Excel by the developer**.
- Local copies on the Mac (all equal to GitHub `main`):
  - `~/Music/NSE_Data_Fetcher_v3-main` — **the working folder** (a git clone since v3.3.1; `origin` is SSH). Use this one: it holds `venv/`, the workbook and its ledger (all git-ignored), and runs/pushes happen from here.
  - `.claude/worktrees/<name>/` — git worktrees the Claude desktop app may create per session (ignored via `.git/info/exclude`). They have **no `venv/` and no workbook**; edit and commit there on the `claude/…` branch, then merge into `main` from the working folder (`git merge --ff-only <branch>`) and push from there. Run tests with the working folder's venv (`../../../venv/bin/python -m pytest`).
  - `~/Music/NSE_Data_Fetcher_v3_repo` — an older duplicate clone created for pushing; safe to delete.
  - `~/Music/NSE_Data_Fetcher_v3-main_backup.zip` — the untouched original v3.1.0 import, taken before any change.
- GitHub access: SSH key `~/.ssh/id_ed25519` is registered on the account (user added the public key). `gh` CLI is **not** installed. No global git identity is set; commits use
  `git -c user.name=Uddhav07 -c user.email=Uddhav07@users.noreply.github.com` (user chose the noreply address). Commit/PR trailers: `Co-Authored-By: Claude …`.

## 3. How the conversation went (summary)
| Step | What the user asked | Outcome |
|---|---|---|
| 1 | Understand the imported repo (`NSE_Data_Fetcher_v3`), replicate, plan, instructions; **drop the Windows app**, scripts + steps only | Read all source; wrote `docs/ARCHITECTURE.md`, `docs/PLAN.md`, `CLAUDE.md` |
| 2 | "Yes, in place" | Backed up to zip; removed GUI/PyInstaller/Inno/.bat; added `setup.sh`/`run.sh`; fixed Windows-only `os.startfile`; Python 3.9 compat; tests; live run |
| 3 | Add **Date Total** (`9, 3`) and **RSI 5/15/30/45**; user edits must survive updates | Clarified "9, 3" = day, digit-root of date; RSI = last candle of day; fill-only-empty logic |
| 4 | Moon nakshatra 09:15; RSI at 15:30; RSI ≥70 / ≥80 colours; Date Total `3, 6`/`6, 3` highlight; **always cell-only** | Added column; CF rules; saved the cell-only preference to memory |
| 5 | Colour fixes (distinct colours; 70+ must not include 80+) | Made RSI rules mutually exclusive (Excel priority bug) |
| 6 | Start from Jan 2026 once; sheet opens at latest date | Full refresh; scroll-to-latest |
| 7 | Left-align nakshatra, no leading spaces | Done |
| 8 | Upload to GitHub, update README, consistency | Cleaned code/docs; first push (v3.2.0) after SSH key setup |
| 9 | Windows install guide; `.command`/`.bat` launchers | Added; README Windows section |
| 10 | User chooses install folder; backfill gaps; "no module named pandas" | `bootstrap.py`; gap insertion with formula repair (v3.3.0) |
| 11 | Windows: nakshatra missing; move Future/Expiry/Comments to the end | Pure-Python Moon (validated vs Swiss Ephemeris); layout migration (v3.4.0) |
| 12 | Format as Excel Table for filtering | `NSE_Table` (v3.5.0) |
| 13 | Add RSI 60m | New column + 2nd migration; 60m history ~2 years (v3.6.0) |
| 14 | "Today's 5m RSI is wrong in Excel?" | Root cause: Yahoo revises final candles/close; added refresh ledger + `--refresh-recent` (v3.7.0) |
| 15 | Two scripts: run+open, and update-clone+run+open | `update_and_run.*`, `--open/--no-open` (v3.8.0) |
| 16 | One file with requirements/why/theme/strategy | `PROJECT.md` (v3.8.1) |
| 17 | This handover | `HANDOVER.md` |

## 4. Repository map (what each file is for)
```
scripts/
  main.py          CLI + run_fetch() orchestration; calls bootstrap first; flags below
  bootstrap.py     stdlib-only: if pandas/yfinance/openpyxl/requests missing → build ./venv, pip install, re-run inside it
  config_manager.py  AppConfig dataclass, JSON load/validate, writes default config
  paths.py         base/data/logs/archive/config dirs; resolve_relative()
  fetcher.py       yfinance daily OHLC with retry/backoff; drops NaN/zero rows
  futures.py       Groww scrape (__NEXT_DATA__) → (price, expiry label); soft failure
  rsi.py           5/15/30m (60d) + 60m (700d) candles → Wilder RSI(14) → last candle per day; 45m built from 15m
  market.py        IST helpers: session_in_progress() = weekday & before 15:35 IST
  nakshatra.py / moon_calc.py   pure-Python Meeus Moon longitude + Lahiri ayanamsa → 27 nakshatras
  excel_writer.py  THE core: layout constants, migrations, row insertion, formula repair, fill-empty, refresh ledger, CF rules, table, scroll
  __init__.py      __version__ (bump every release)
tests/   offline pytest suite (see §8) · tests/golden/ snapshot · .github/workflows/tests.yml CI · pyproject.toml tool config
config/config.json   user settings (ticker, excel_file, start_date, retries, open_excel_after_run, backup…)
setup.sh/.command/.bat   one-time venv + deps          run.sh/.command/.bat   update + open (per config)
update_and_run.sh/.command/.bat   git pull --ff-only (only if tracked files unmodified) + deps + run with --open
requirements.txt (runtime) · requirements-dev.txt (+pytest, ruff) · .gitattributes (.bat CRLF; .sh/.command LF) · .gitignore
docs/ARCHITECTURE.md  internals     docs/PLAN.md  status & ideas     PROJECT.md  requirements/why/theme/strategy
CHANGELOG.md  every release         CLAUDE.md  rules for editors        README.md  user guide (Mac/Linux/Windows)
```
Runtime-generated, git-ignored: `data/*.xlsx`, `data/*.state.json`, `archive/` (timestamped backups before each write), `logs/*.log`, `venv/`.

## 5. How a run works (`main.run_fetch`)
1. `bootstrap.ensure_dependencies()` (at import of `main`) — self-heals missing packages.
2. Load config; optionally `--full-refresh` (backup, delete, recreate workbook).
3. Fetch window = `min(start_date, last_date_in_sheet + 1)` → always re-checks from `start_date` for gaps.
4. Backup workbook to `archive/`.
5. Fetch OHLC. If the NSE session is still open (before 15:35 IST on a weekday): drop today's row, skip futures, skip today's RSI.
6. Fetch futures (unless `--no-futures`/live) and RSI map (unless `--no-rsi`).
7. `update_workbook()`: migrate old layouts → normalise headers → fill empty OHLC on existing dates → **insert missing dates in order** (`insert_rows`) → repair/fill formulas → futures (newest row, empty cells only) → Nakshatra/RSI (empty cells only) → refresh-ledger pass for the latest 3 sessions → CF rules → Excel Table → freeze panes/scroll → save → save ledger.
8. Open Excel (`open`/`os.startfile`/`xdg-open`) per `--open/--no-open` or config.

CLI flags: `--config`, `--start`, `--full-refresh`, `--no-futures`, `--no-rsi`, `--refresh-recent [N]`, `--open`, `--no-open`, `--version`.

## 6. The current Excel layout (a contract — don't reorder casually)
`A Day · B Date · C–F OHLC · G Daily Change · H Weekly Change · I Close Above prev High · J Close Below prev Low · K HH/HL · L LL/LH · M Diff in Future · N Date Total · O–S RSI 5/15/30/45/60m · T Moon Nakshatra (9:15am) · U Future as on Date · V Future Expiry · W Comments`
- Headers are single-line unique strings (Excel Table requirement). Table ref `A1:W<last>` is rebuilt every run.
- Layout history is encoded in `_migrate_layout`: layout **A** (≤3.3: Future/Expiry/Comments at N–P) → **B** (3.4–3.5) → **C** (current, adds RSI 60m). Any change to the layout needs a new migration step + a test.
- Colours (cell-only): RSI <30 red, 70–79.99 light green, ≥80 dark green; Date Total `3, 6`/`6, 3` purple; change columns green/red; **Exception:** Tuesday rows get a full-row light-blue fill (A–W, `DDEBF7`) **only on rows the script creates** — the only whole-row style allowed; the green/red change rules skip Tuesdays, and the cell-only RSI/Date Total colours still show on top.
- M (Diff in Future) **stays at M** — decided by the user (Nov 2026); only Future/Expiry/Comments live at the end.

## 7. Rules that must keep holding (the design contract)
1. **Fill only empty cells.** Never overwrite a value, fill, font, header, note or user CF rule.
2. **One exception:** the refresh ledger (`data/<workbook>.state.json`) lets the latest 3 sessions' OHLC/RSI cells update to Yahoo's final values **only if the cell still equals what the script wrote**. `--refresh-recent` is the explicit override. Legacy sheets without a ledger are *not* auto-refreshed.
3. Comments (W) is user-owned and stays last.
4. Highlights are **cell-only**, never rows (also saved as a persistent feedback memory). Sole exception: the Tuesday full-row fill on script-created rows.
5. Backup before every write; friendly errors (e.g. "close Excel").
6. No hard-coded user folders in code or docs; `excel_file` can be any path.
7. Futures failing must not block OHLC; each optional source degrades one column only.
8. Code is truth where docs disagree — then fix the doc in the same commit.
9. Never sort the table (formulas reference the previous row; the script assumes date order). Filters only.

## 8. Testing and verification
- Run: `venv/bin/python -m pytest` → **54 pass**. Lint: `venv/bin/ruff check scripts tests` → clean.
  - `test_core.py`: gap insertion + formula repair, edit preservation, A→C and B→C migrations, table, refresh ledger, bootstrap, futures/RSI/Date Total.
  - `test_golden.py`: **the exact workbook output** of a two-run scenario vs `tests/golden/workbook_scenario.json`. If an output change is intended: `UPDATE_GOLDEN=1 venv/bin/python -m pytest tests/test_golden.py`, then review the JSON diff.
  - `test_pipeline.py`: `run_fetch` with all sources mocked (live session, futures failure, flags, full refresh, locked file).
  - `test_indicators.py`: RSI, 45m candles, 15:30 cut-off, Moon vs Swiss Ephemeris/Meeus, nakshatra, session hours.
- CI (`.github/workflows/tests.yml`): same suite on Ubuntu/macOS/Windows × Python 3.9/3.12 on every push.
- Manual scenarios already verified on Mac: fresh folder + bare system Python without pandas (bootstrap), deleted-row restore with edits intact, update script in a scratch clone (clean / locally-modified / diverged).
- **Not verified:** real Windows run of `.bat` files, opening in real Excel (table validity only inspected structurally), LibreOffice.

## 9. Gotchas learned (read before editing)
- openpyxl `insert_rows` shifts cells but **not** formulas, CF, row heights or hidden rows → `_ensure_formulas` repairs formulas whose first reference ≠ own row; filters are cleared if a mid-sheet insert happens.
- Excel CF with two overlapping RSI rules showed the wrong colour → rules are mutually exclusive formulas.
- Yahoo: 5/15/30m ≈ 60 days, 60m ≈ 2 years (`700d`); last intraday candles and the day's close are revised after the session (see ledger); 60m last bar 15:15–15:30.
- A `.sh` double-click opens an editor on macOS → `.command` wrappers. `.bat` must be CRLF.
- zsh quirks seen in this project: unquoted globs in `rsync --exclude` and `rm`; `$VAR` holding a multi-word `git -c …` command fails — write the command out.
- **Process lesson (happened twice):** a multi-step Python edit script aborted on a failed `assert` but the commit/push ran anyway, shipping a release with stale version/docs (fixed by follow-up tags 3.5.1 and 3.7.1). Always confirm the edit output says success **before** committing.
- Don't rewrite published tags; make a follow-up release instead.

## 10. Release checklist (what the previous sessions did each time)
1. Make the change in `~/Music/NSE_Data_Fetcher_v3-main`; keep `docs/` + `README.md` + `CLAUDE.md` + `PROJECT.md` consistent.
2. Bump `scripts/__init__.py` `__version__`; add a `CHANGELOG.md` section.
3. `pytest` + `ruff check scripts tests`; run the script on the real sheet or a scratch copy.
4. `git add -A && git -c user.name=Uddhav07 -c user.email=Uddhav07@users.noreply.github.com commit …` (with the Co-Authored-By trailer), annotated tag `vX.Y.Z`.
5. `GIT_SSH_COMMAND="ssh -o BatchMode=yes" git push origin main --follow-tags`.
6. Optionally `git pull --ff-only` in `~/Music/NSE_Data_Fetcher_v3_repo` (or delete it).
7. GitHub *Releases* pages are not created (no `gh`); changelog sections can be pasted manually.

## 11. State right now
- `main` = `v3.8.3`; working folder clean and equal to `origin/main`.
- The workbook in the working folder (`data/stock_market_NSE50.xlsx`, 210 rows 2026-01-01 → 2026-11-09) is **test/demo data built by the script**; it is git-ignored. It may be open in Excel (a `~$` lock file was seen).
- The user's real Windows copy is on another machine; it upgrades itself (layout migration + bootstrap) after `git pull` or `update_and_run.bat`. It may still hold values saved before the ledger existed → tell them to run `--refresh-recent` once if a recent day looks off.

## 12. Open items / ideas (none are in progress)
- Verify on real Windows + real Excel (table opens without repair prompt, colours render, `.bat` flow).
- Sturdier futures source than the Groww scrape.
- Create GitHub Release notes from `CHANGELOG.md` (needs `gh` or manual).
- Optional: delete the duplicate `NSE_Data_Fetcher_v3_repo` clone.
- Possible extras the user hasn't asked for: Task Scheduler/launchd automation (documented for Windows only), more panchang columns (tithi), other tickers (`^NSEBANK`).

## 13. Quick commands
```bash
cd ~/Music/NSE_Data_Fetcher_v3-main
./run.sh                    # update + open per config      ./update_and_run.sh   # pull, update, open
./run.sh --refresh-recent   # one-time fix of latest 3 rows  venv/bin/python -m pytest
git status -sb && git log --oneline -5
```
