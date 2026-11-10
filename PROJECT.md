# NSE Data Fetcher — Project Brief

One file that answers: **what must this project do, why was each decision made, what is the theme, and which strategic choices shaped it.**
Other docs: `HANDOVER.md` (state of the project, file map, gotchas), `README.md` (how to use), `docs/ARCHITECTURE.md` (how it works inside), `CHANGELOG.md` (what changed when), `CLAUDE.md` (rules for editing).

Legend: **[User]** = stated by the project owner · **[Eng]** = engineering reasoning added while building.

---

## 1. Purpose and use case
A **daily Nifty 50 trading journal in one Excel file.** Each run appends the new trading days' OHLC, price-action formulas, the current Nifty futures price, intraday RSI at the close, and the Moon's nakshatra — so the user can scan history in Excel, filter it, annotate it, and keep their own notes next to the data.

The sheet is a **working document the user edits**, not a report the script owns. That single fact drives most decisions below.

## 2. Requirements

### 2.1 Functional
| # | Requirement | Source |
|---|---|---|
| F1 | Fetch Nifty 50 (`^NSEI`) daily OHLC from Yahoo Finance; incremental (only what's missing) | original v2.1 / v3 |
| F2 | Futures price + expiry on the newest row only (Groww scrape) | original |
| F3 | Formula columns: daily change, weekly change (5 trading rows back, holiday-aware), close above prev high / below prev low, HH-HL, LL-LH, futures difference | original |
| F4 | Tuesday rows highlighted across the whole row A–W (created rows only) — the one exception to cell-only | original |
| F5 | **Date Total** column, e.g. `9, 3` = day number, then the date's digits summed to one digit (2026-11-09 → 21 → 3) | [User] |
| F6 | **RSI(14)** for 5 / 15 / 30 / 45 / 60-minute candles, always as of the **15:30 IST** close | [User] |
| F7 | **Moon Nakshatra** at 09:15 IST for each day | [User] |
| F8 | Cell-only colour rules: RSI <30 red, 70–79.99 light green, ≥80 dark green; Date Total `3, 6` / `6, 3` purple. **Never whole rows** (except F4 Tuesday fill) | [User] |
| F9 | Data formatted as an Excel Table so every column can be filtered | [User] |
| F10 | Sheet opens scrolled to the latest date | [User] |
| F11 | Column order ends with Future, Expiry, Comments (Comments always last) | [User] |
| F12 | Running the script fills **all** missing days and empty cells; already-filled cells and the user's row edits stay untouched | [User] |
| F13 | Two launchers: (1) run + open the workbook, (2) update the clone to latest, then run + open | [User] |

### 2.2 Non-functional
| # | Requirement | Source |
|---|---|---|
| N1 | Script/CLI only — **no Windows app, GUI or installer** | [User] |
| N2 | Works on macOS **and Windows** with copy-paste steps; Linux as a bonus | [User] |
| N3 | The user chooses where the project lives; nothing assumes a folder | [User] |
| N4 | "No module named pandas" must never happen again | [User] |
| N5 | The project is versioned on GitHub with a changelog; local clones stay current | [User] |
| N6 | Docs and code must never contradict each other | [User] / [Eng] |
| N7 | Never crash silently; friendly messages (e.g. "close Excel"); every write preceded by a backup | original / [Eng] |

## 3. Theme
1. **The user owns the sheet.** The script is a careful guest in a document the user edits.
2. **Boring reliability over cleverness.** Same result on every OS, few dependencies, self-repairing setup.
3. **Honest about data limits.** Free data sources have gaps and revisions; the project says so instead of hiding it.
4. **Zero-knowledge usability.** A non-programmer can install, run, update and recover by copy-pasting.
5. **Docs are part of the product.** README, changelog and architecture are updated with every behaviour change.

## 4. Decisions and why

### 4.1 Data sources
| Decision | Why | Trade-off |
|---|---|---|
| **Yahoo Finance via `yfinance`** for OHLC and intraday | Free, no key, one library for daily and intraday, works for `^NSEI` | Intraday history is capped (~60 days for 5/15/30m, ~2 years for 60m); final candles get revised after close |
| **Groww page scrape** for futures | Only free source of the current-month Nifty future found in the original project | Fragile (depends on private `__NEXT_DATA__`); isolated in `futures.py`, failure never blocks OHLC |
| **RSI = Wilder, 14 periods, last candle of the day** | Standard definition; I checked an SMA-seeded variant and it gave the identical value, so seeding is not a source of error | Differs slightly from other platforms: different data feeds |
| **45-min candles built from 15-min** | Yahoo has no native 45m | The last 45m candle covers only 15:15–15:30 |
| **RSI at 15:30 consistently** | Requested; every timeframe's last candle ends at 15:30, which I enforce by dropping candles starting at/after 15:30 | — |
| **Today's row, RSI and futures skipped until 15:35 IST** | A mid-session candle is not final; saving it would be wrong data in a "never overwrite" sheet | Run after the close |

### 4.2 Excel design
| Decision | Why |
|---|---|
| **Derived columns stay formulas** (A, G–L, M, N), not values | The user can edit/extend them in Excel and see how they work |
| **Fill-only-empty-cells, never overwrite filled ones** | The user's explicit requirement; the sheet doubles as their notebook |
| **Missing days are inserted in date order; shifted formulas are re-pointed** | "All previous un-updated rows should be updated" without moving or rewriting rows that are already right |
| **RSI and Nakshatra are written as static values; Date Total is a formula** | RSI and nakshatra cannot be Excel formulas, so Python computes them and writes them once; Date Total can be a formula, so it stays editable |
| **Comments column is user-owned and always last** | New script columns can be added before it without ever disturbing the user's notes |
| **Excel Table (`NSE_Table`) instead of a bare auto-filter** | Filter buttons, auto-growth, and it is what "format as table" means to the user |
| **Headers are single-line text** | Excel tables require column names to match cells exactly; line breaks break that |
| **Filter, don't sort** | Formulas reference the previous row, and the script assumes date order; sorting would scramble them |
| **Cell-only highlighting** | User preference: highlight the cell that meets the condition, never the row |
| **Tuesday fill applied only to rows the script creates** | Restyling existing rows would overwrite the user's own formatting |
| **Column layout is a contract; changes ship with an automatic migration** | Existing workbooks (including the user's Windows copy) must upgrade without losing notes, values, styles or cell comments |

### 4.3 The refresh ledger (the one exception to "never overwrite")
**Problem found:** Yahoo revises its last intraday candles and the daily close after the session. A value saved a few minutes too early stayed frozen (e.g. a 5m RSI using a provisional candle; a mid-session Close).
**Decision:** record every OHLC/RSI value the script writes (`data/<workbook>.state.json`). For the **latest 3 sessions**, refresh a cell to the final number **only if it still holds exactly what the script wrote**. A user's edit changes the value, so it is never touched.
**Escape hatch:** `--refresh-recent [N]` for sheets created before the ledger existed (explicit user command = permission).
**Why not simply allow overwrites:** that would destroy user edits; the ledger separates "script's value" from "user's value".

### 4.4 Calculations chosen by the user
- **Date Total** `day, digit-root`: implemented with an Excel formula (`1+MOD(yyyymmdd−1,9)`) so it recalculates and works with no extra tools. Digit-root of 20261109 = 3, matching the user's example (`9, 3`).
- **Moon Nakshatra** at 09:15 IST: sidereal (Lahiri ayanamsa) Moon longitude ÷ 13°20′.
  - First built on the Swiss Ephemeris package (`pyswisseph`), then **replaced by a pure-Python Meeus calculation** because the compiled package can fail to install on Windows, leaving the column blank.
  - Validated against Swiss Ephemeris over 3,650 days: max error 0.008°, **0 nakshatra differences**. Same answer on every OS, one dependency fewer.

### 4.5 Platform and setup
| Decision | Why |
|---|---|
| **Drop the Windows GUI / PyInstaller / Inno Setup** | User: scripts and steps are enough; far less surface to maintain |
| **Python 3.9+ (not 3.10+)** | The Mac in use only had 3.9; `from __future__ import annotations` makes the code compatible |
| **Project-local `venv/`, called directly (`venv/.../python`)** | No activation needed; avoids global-Python surprises |
| **`bootstrap.py`: auto-install and relaunch in the venv if a package is missing** | Direct answer to "no module named pandas, make sure this does not happen again"; works from system Python, IDEs, any interpreter |
| **`setup/run` as `.sh`, `.command` (Finder double-click) and `.bat`** | A `.sh` opens in an editor when double-clicked on macOS; `.command` runs in Terminal. `.bat` for Windows |
| **`.gitattributes` (`.bat` CRLF, `.sh`/`.command` LF)** | Scripts break if line endings are wrong |
| **Console set to replace unencodable characters** | Legacy Windows code pages can't print "—" and would raise errors |
| **User picks the install folder; `excel_file` can be any path** | User requirement; docs show how to open a terminal in any folder |

### 4.6 Update workflow
- **`update_and_run`**: `git pull --ff-only` only if tracked files are unmodified; otherwise it **skips the update with a message** and still runs. Never auto-stash, reset or force. Offline, diverged or non-git copies also fall back to running the current version.
- The script body is wrapped so a pull that rewrites the script itself cannot break the run.
- `--open` / `--no-open` make opening Excel independent of the config.

### 4.7 Engineering and repository practice
- **Modular package** (`fetcher`, `futures`, `rsi`, `nakshatra`, `moon_calc`, `excel_writer`, `config_manager`, `bootstrap`, `market`, `paths`, `main`); `run_fetch()` returns a result object and never exits.
- **Tests without network** (13 pytest tests): gap insertion, formula repair, edit preservation, layout migrations, table, refresh rules, bootstrap.
- **Real-world checks beyond unit tests:** clean-room fresh-folder runs, a bare system Python without pandas, a deleted-row restore, a scratch-remote test of the update script.
- **Semantic version tags and a changelog** for every release; the local clones are kept equal to GitHub `main`.
- **Commit identity:** GitHub noreply address, chosen by the user, rather than any work email.
- **Honesty about verification:** Windows scripts and Excel rendering have not been exercised on real Windows/Excel; this is stated in the docs.

## 5. Strategic choices
1. **Evolve, don't rewrite.** Keep v3's proven engine (incremental, formula-driven, backups) and add features on top; remove only what the user said they don't need (the Windows app).
2. **Fix causes, not symptoms.** Blank Nakshatra on Windows → removed the fragile dependency. Wrong "final" values → ledger, not "overwrite everything".
3. **Make failure soft.** Futures failing, RSI unavailable for old days, offline updates — each degrades one column or step, never the whole run.
4. **Ship migrations with every layout change.** Column order changed three times; each change auto-migrates older sheets so no user data is stranded.
5. **Protect user data over developer convenience.** Backup before every write, fill-only-empty, ledger-guarded refresh, update skips when files are modified.
6. **Keep the user in control of location and timing.** No hard-coded paths; the update step is a separate, explicit launcher.

## 6. Known limitations / honest caveats
- Yahoo intraday depth: 5/15/30/45m RSI exist for ~60 days only; 60m for ~2 years. Older rows stay blank.
- Differences from other charting platforms are expected (different feeds; Yahoo's final 5m candle includes the official close).
- Groww scraping can break without notice.
- openpyxl drops charts/images/pivots added to the sheet when it saves.
- If a missing day is inserted mid-sheet, active filters are cleared (row visibility cannot follow inserted rows).
- Not tested on real Windows or inside real Excel (file validity checked structurally, not visually).
- User-added conditional-format rules on moved columns do not follow a layout migration.

## 7. Layout contract (current)
`A Day · B Date · C–F Open/High/Low/Close · G Daily Change · H Weekly Change · I Close Above prev High · J Close Below prev Low · K HH/HL · L LL/LH · M Diff in Future · N Date Total · O–S RSI 5/15/30/45/60m · T Moon Nakshatra (9:15am) · U Future as on Date · V Future Expiry · W Comments`

## 8. How the project evolved
| Version | Milestone |
|---|---|
| 3.1 (original) | Windows desktop app, 16-column sheet |
| 3.2 | Script edition for macOS; Date Total, RSI 5/15/30/45, Nakshatra; cell-only colours; edit-safe updates |
| 3.3 | Self-healing dependencies; gap backfill; user-chosen folder |
| 3.3.1 | Full Windows guide; `.gitattributes`; Windows console safety |
| 3.4 | Pure-Python nakshatra (Windows fix); Future/Expiry/Comments moved to the end |
| 3.5 | Excel Table with filters |
| 3.6 | RSI 60m |
| 3.7 | Refresh ledger so recent days self-correct |
| 3.8 | `update_and_run` launchers, `--open` / `--no-open` |
