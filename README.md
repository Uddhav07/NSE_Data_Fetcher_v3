# NSE Data Fetcher v3.2 — script edition

A daily Nifty 50 journal in one Excel file. Each run appends the new trading days'
OHLC (Yahoo `^NSEI`), price-action formulas, the current Nifty futures price/expiry
(Groww), intraday RSI at the 15:30 close, and the Moon's nakshatra at 09:15 IST.

Runs from the terminal on macOS/Linux (Python 3.9+). No GUI, no installer.
Why things are the way they are (requirements, decisions, strategy): see [PROJECT.md](PROJECT.md).
Rewritten from [v2.1](https://github.com/pmmittal-byte/NSE_Data_Fetcher_v2.1); v3.1 shipped a Windows app, which this edition drops.

## Quick start (macOS / Linux)
Open a terminal **in the folder where you want the project** (you choose the location), then:
```bash
git clone https://github.com/Uddhav07/NSE_Data_Fetcher_v3.git
cd NSE_Data_Fetcher_v3
./setup.sh                                    # one-time: venv + dependencies
./run.sh --full-refresh --start 2026-01-01    # one-time: build the sheet from Jan 2026
./run.sh                                      # every day after 15:35 IST
```
Close the workbook in Excel before running (a locked file gives a friendly error).
The sheet opens scrolled to the latest date.

**macOS double-click:** use `setup.command` / `run.command` in Finder (a `.sh` file opens in an editor instead of running). If macOS blocks it the first time: right-click → Open.

**What a run does:** adds every missing trading day (new days at the end, older gaps inserted in date order),
fills any empty cells (OHLC, formulas, Date Total, RSI, Nakshatra, futures) and **never changes a cell that
already has a value** — your edits stay. Before 15:35 IST on a trading day, today's row and futures price are
skipped (not final yet); just run again after close.

**"No module named pandas" can't happen:** if a required package is missing from the Python you used
(system Python, IDE interpreter...), the script creates `venv/` in the project, installs the packages, and
re-runs itself inside it automatically.

## The two scripts you need
| | macOS / Linux | Windows | What it does |
|---|---|---|---|
| **1. Run** | `./run.sh` or double-click `run.command` | `run.bat` | Adds missing days, fills empty cells, refreshes the latest sessions, opens the workbook (if `open_excel_after_run` is true in the config — the default) |
| **2. Update + run** | `./update_and_run.sh` or double-click `update_and_run.command` | `update_and_run.bat` | `git pull` to the latest version, refreshes packages, then does everything in (1) and **always** opens the workbook |

The update step never overwrites your work: if you changed tracked project files, or you're offline / not a git clone, it says so, skips the update and still runs.
Extra flags pass through, e.g. `./update_and_run.sh --no-futures` or `update_and_run.bat --no-open`. Use `--open` / `--no-open` on any run to force opening or not.

## Windows (10 / 11) — complete guide
Everything below is copy & paste. **You choose where the project lives** — nothing is hard-coded.
The Windows launchers are untested on a real Windows PC; the Python code itself is cross-platform
(the console is set to never crash on special characters, and Excel opens via `os.startfile`).

### 1. Install Python 3.9 or newer (3.9–3.13)
Download from <https://www.python.org/downloads/windows/> and **tick "Add python.exe to PATH"** on the first screen.
Or, in PowerShell:
```powershell
winget install Python.Python.3.12
```
Close and reopen PowerShell, then check:
```powershell
python --version
```
(If `python` opens the Microsoft Store, install from python.org or use `py -3 --version`.)

### 2. Install Git — or skip it
```powershell
winget install Git.Git
```
No Git? On the GitHub page click **Code → Download ZIP**, extract it to the folder you want, and skip `git clone` below
(updating later = download the ZIP again and copy over your `data\` and `config\` folders).

### 3. Pick a folder and open PowerShell in it
In File Explorer open the folder where you want the project (e.g. `D:\Trading`), click the address bar, type `powershell`, press Enter.

### 4. Get the project
```powershell
git clone https://github.com/Uddhav07/NSE_Data_Fetcher_v3.git
cd NSE_Data_Fetcher_v3
```

### 5. One-time setup + first build from Jan 2026
Easiest: double-click **`setup.bat`**, then in PowerShell:
```powershell
venv\Scripts\python.exe -m scripts.main --full-refresh --start 2026-01-01
```
Or do the setup by hand (same result):
```powershell
python -m venv venv
venv\Scripts\python.exe -m pip install --upgrade pip
venv\Scripts\python.exe -m pip install -r requirements.txt
venv\Scripts\python.exe -m scripts.main --full-refresh --start 2026-01-01
```
The workbook is created at `data\stock_market_NSE50.xlsx` and opens in Excel on the latest date.
You never need to "activate" the venv — and if a package is ever missing, the script installs it by itself.

### 6. Every day (after 15:35 IST, Excel file closed)
Either double-click **`run.bat`**, or open PowerShell/cmd in the project folder and run:
```powershell
venv\Scripts\python.exe -m scripts.main
```
Options go on the end: `--no-futures`, `--no-rsi`, `--config PATH`. With the batch file: `run.bat --no-rsi`.
A run adds missing days, fills empty cells, and never changes cells that already have values.

### 7. Update to a newer version
```powershell
git pull
venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Optional: activate the venv (to type plain `python`)
PowerShell:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
venv\Scripts\Activate.ps1
python -m scripts.main
deactivate
```
Command Prompt (cmd): `venv\Scripts\activate.bat`

### Optional: run automatically every weekday
Replace `C:\Trading\NSE_Data_Fetcher_v3` with **your** project path; the time is your PC's local time (15:40 = after the IST close — convert if your PC is not on IST).
```powershell
schtasks /Create /TN "NSE Data Fetcher" /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 15:40 /TR "cmd /c cd /d C:\Trading\NSE_Data_Fetcher_v3 && venv\Scripts\python.exe -m scripts.main"
```
To stop it: `schtasks /Delete /TN "NSE Data Fetcher" /F`. Set `"open_excel_after_run": false` in `config\config.json` if you don't want Excel popping up.

### Windows troubleshooting
| Problem | Fix |
|---|---|
| `python` is not recognized / opens Microsoft Store | Reinstall from python.org with "Add python.exe to PATH", reopen PowerShell; or use `py -3` in place of `python` |
| `Activate.ps1 cannot be loaded` (execution policy) | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, or skip activation and use `venv\Scripts\python.exe` |
| "Cannot save ... file is locked" | Close the workbook in Excel (check Task Manager for a hidden Excel), run again |
| `No module named pandas` (or another package) | Should self-heal. Manually: `venv\Scripts\python.exe -m pip install -r requirements.txt` |
| Moon Nakshatra column blank | Update to the latest version (`git pull`) — it is now calculated in pure Python, no extra package needed. Run `venv\Scripts\python.exe -m scripts.main` again; empty cells are filled |
| SmartScreen / antivirus blocks `setup.bat` | It's a plain text script — right-click → Properties → Unblock, or run the PowerShell commands above |
| `git` is not recognized | Reopen PowerShell after installing Git, or use the ZIP download |
| Corporate proxy / SSL errors | Set `HTTPS_PROXY` in the terminal first: `$env:HTTPS_PROXY="http://proxy:port"` |
| Excel shows `####` | Widen the column; values are fine |

## Options
```
./run.sh                               add missing days + fill empty cells (checks from start_date)
./run.sh --full-refresh --start DATE   rebuild (old file is backed up to archive/ first)
./run.sh --no-futures                  skip Groww futures
./run.sh --no-rsi                      skip intraday RSI
./run.sh --open / --no-open            force opening (or not) the workbook after the run
./run.sh --refresh-recent              one-time fix: overwrite the latest 3 rows' prices/RSI with fresh data
                                        (use if a value was saved before the session was final; `--refresh-recent 5` for 5 rows)
./run.sh --config PATH                 use another config.json
```

## Excel layout (sheet `NSE_Data`)
| Col | Content |
|---|---|
| A–B | Day, Date |
| C–F | Open, High, Low, Close |
| G–H | Daily change; Weekly change (vs 5 trading days back, holiday-aware) |
| I–L | Close above prev High / below prev Low, HH-HL, LL-LH (formulas) |
| M | Diff in Future (`=T − Close`) |
| N | **Date Total** — `day, digit-root of full date`; 2026-11-09 → `9, 3` |
| O–S | **RSI(14)** on 5/15/30/45/60-min candles, as of the 15:30 IST close |
| T | **Moon Nakshatra** at 09:15 IST (sidereal, Lahiri), left-aligned |
| U–V | **Future as on Date**, **Future Expiry** (newest row only) |
| W | **Comments** — yours, never written by the script |

**Filtering:** the data is formatted as an Excel Table (`NSE_Table`), so every column has filter buttons — e.g. filter RSI ≥ 70, Date Total `3, 6`, or one Nakshatra. The table grows by itself as rows are added.
Use **filters only — don't sort** the table: Daily/Weekly change and the High/Low formulas compare with the row above, so sorting would scramble them (and the script expects dates in order). To analyse in another order, copy the data to a new sheet first. If a missing day is ever inserted in the middle of the sheet, active filters are cleared and you re-apply them.

**Updating from an older version?** Your existing workbook is migrated automatically on the next run (a backup goes to `archive/` first): columns shift into the new order (Future, Expiry and Comments stay at the end), and your notes, values and formatting travel with them.

Cell-only colours: RSI below 30 red, RSI 70–79.99 light green, RSI 80+ dark green, Date Total `3, 6` / `6, 3` purple.
Daily/weekly/futures change columns go green/red. Tuesday rows get a light-blue row fill when created.

Notes on the intraday data: Yahoo keeps only ~60 days of 5/15/30-min history (so older rows have no
5/15/30/45m RSI) but about 2 years of 60-min history, so **RSI 60m fills back much further**. 45-min candles are built from 15-min ones. Today's RSI is filled only on a run after 15:35 IST.

## Your edits are safe
The script only **adds missing rows** and **fills empty cells**. It never rewrites a cell that has a value,
nor your fills, fonts, headers, notes or your own conditional-format rules.
The one exception: Yahoo revises its final intraday candles after the close, so for the **latest 3 sessions** the script re-checks the
Open/High/Low/Close and RSI cells *it wrote itself* and updates them to the final numbers — but only if the cell still holds exactly what
the script wrote (it keeps a small ledger, `data/<workbook>.state.json`). Anything you typed over is left alone. Futures are written once per
row. Limitation: charts/images/pivots added to the sheet are dropped when openpyxl saves.

## Configuration — `config/config.json`
`ticker` (`^NSEI`, `^NSEBANK`, `^BSESN`) · `excel_file` (relative to the project, or an absolute path anywhere you like) · `start_date` · `log_level` ·
`max_retries` · `request_timeout` · `open_excel_after_run` · `backup_before_update`

## Project layout
```
setup.sh/run.sh (+ .command wrappers) · setup.bat/run.bat (Windows)
scripts/  main.py (CLI+orchestration) · fetcher.py (Yahoo OHLC) · futures.py (Groww)
          rsi.py (intraday RSI) · nakshatra.py + moon_calc.py (Moon, pure Python) · excel_writer.py · config_manager.py · paths.py
          bootstrap.py (auto-install deps) · market.py (session hours)
config/   config.json        tests/  test_core.py (pytest, no network)
data/ archive/ logs/         generated at runtime (git-ignored)
docs/     ARCHITECTURE.md · PLAN.md        CHANGELOG.md
```
Run tests: `venv/bin/pip install -r requirements-dev.txt && venv/bin/python -m pytest tests`

## Troubleshooting
| Problem | Fix |
|---|---|
| "file is locked" | Close the workbook in Excel |
| `No module named ...` | Should self-heal; if not: `venv/bin/python -m pip install -r requirements.txt` (Windows: `venv\\Scripts\\python.exe ...`) |
| "No new data" | Market closed or already up to date (futures/RSI still refresh) |
| Futures blank | Groww page changed or down; OHLC is still saved |
| RSI / Close of a recent day looks wrong vs your chart | Run `./run.sh --refresh-recent` once (Windows: `venv\\Scripts\\python.exe -m scripts.main --refresh-recent`). Small differences vs TradingView can remain: it uses a different data feed, and Yahoo's final 5-min candle includes the official closing value |
| RSI blank on old rows | Yahoo intraday history limit (~60 days; ~2 years for 60m) |

Logs: `logs/nse_fetcher.log`. Provided as-is for personal and educational use; not financial advice.
