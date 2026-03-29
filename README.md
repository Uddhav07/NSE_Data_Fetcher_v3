# NSE Data Fetcher v3

Automated **Nifty 50 Index** data fetcher with technical analysis, futures tracking, and professional Excel output. Available as a **standalone Windows desktop application** — no Python installation needed.

> Rewritten from [v2.1](https://github.com/pmmittal-byte/NSE_Data_Fetcher_v2.1) with a modular architecture, proper logging, retry logic, and an improved Excel layout.

---

## Download & Install

### Option A: Windows Installer (recommended)

1. Download **`NSE_Data_Fetcher_Setup.exe`** from [Releases](https://github.com/Uddhav07/NSE_Data_Fetcher_v3/releases).
2. Run the installer — no admin privileges required.
3. Launch **NSE Data Fetcher** from the Start Menu or Desktop shortcut.

### Option B: Standalone .exe (portable)

1. Download the **`NSE_Data_Fetcher_Portable.zip`** from [Releases](https://github.com/Uddhav07/NSE_Data_Fetcher_v3/releases).
2. Extract anywhere and double-click **`NSE Data Fetcher.exe`**.

---

## What's New in v3

| Area | v2.1 | v3 |
|---|---|---|
| **Architecture** | Single 290-line file | 5 focused modules |
| **GUI** | None | Desktop app with tkinter |
| **Distribution** | Requires Python | Standalone .exe + Windows installer |
| **Logging** | `print()` statements | Rotating file + console via `logging` |
| **Retries** | None | Exponential back-off (configurable) |
| **Futures bug** | Same price stamped on *every* row | Futures only written for today's row |
| **Excel formatting** | Basic colors | Conditional green/red, Tuesday highlighting, freeze panes |
| **CLI** | None | `--full-refresh`, `--start`, `--no-futures`, `--version` |
| **Config validation** | None | Dataclass-based with typed defaults |
| **Backups** | None | Timestamped archive before each update |
| **Data validation** | None | NaN/zero row filtering |
| **Weekly Change** | Fixed -5 rows | Holiday-aware date lookup |
| **Start Date** | Hard-coded | Defaults to Jan 1 of current year |
| **Execution** | Scheduled task | One-click `stock_market.bat` |

---

## Features

- **Desktop GUI** — Update Data, Open Excel, Settings, progress bar, live log output
- **OHLC Data** from Yahoo Finance (`^NSEI`)
- **Futures Price & Expiry** from Groww.in (current month)
- **Technical Analysis** formulas — HH/HL, LL/LH, daily & weekly changes
- **Tuesday Highlighting** — all Tuesday rows highlighted in Excel
- **Holiday-Aware Weekly Change** — compares with correct trading day ~7 calendar days back
- **Duplicate Prevention** — existing dates are never overwritten
- **Professional Excel** — color-coded headers, conditional red/green, frozen header row
- **Environment Health Checks** — validates Python, packages, and network on startup
- **Auto-Open Excel** — opens the file automatically after each update
- **Rotating Logs** — 5 MB × 3 backups in `logs/`

---

## Project Structure

```
NSE_Data_Fetcher_v3/
├── scripts/
│   ├── __init__.py          # Package marker + version
│   ├── __main__.py          # python -m scripts entry
│   ├── main.py              # CLI, orchestration, logging, health checks
│   ├── gui.py               # Desktop GUI (tkinter + ttk)
│   ├── paths.py             # Path resolution (dev vs frozen exe)
│   ├── config_manager.py    # Config loading and validation
│   ├── fetcher.py           # Yahoo Finance OHLC fetcher
│   ├── futures.py           # Groww futures scraper
│   └── excel_writer.py      # Workbook creation & update
├── config/
│   └── config.json          # User settings
├── assets/
│   └── icon.ico             # Application icon
├── installer/
│   └── installer.iss        # Inno Setup script
├── data/                    # Excel output (auto-created)
├── logs/                    # Rotating log files (auto-created)
├── archive/                 # Timestamped backups (auto-created)
├── app.py                   # GUI entry point (PyInstaller target)
├── nse_fetcher.spec         # PyInstaller build spec
├── build.bat                # One-click build: exe + installer
├── setup.bat                # Dev setup: venv + deps
├── stock_market.bat         # CLI entry: double-click to update
├── requirements.txt         # Runtime dependencies
├── requirements-dev.txt     # Build dependencies (pyinstaller)
├── .gitignore
└── README.md
```

---

## Quick Start

### 1. Clone & Setup

```bat
git clone https://github.com/Uddhav07/NSE_Data_Fetcher_v3.git
cd NSE_Data_Fetcher_v3
setup.bat
```

This creates a Python **virtual environment** (`venv/`) and installs all dependencies.

### 2. Fetch Data

Double-click **`stock_market.bat`**, or from a terminal:

```bat
venv\Scripts\python.exe -m scripts.main
```

The batch file will auto-run setup if the venv is missing.

---

## CLI Options

```
usage: main.py [-h] [--config PATH] [--start YYYY-MM-DD]
               [--full-refresh] [--no-futures] [--version]

  --config PATH         Path to config JSON (default: config/config.json)
  --start YYYY-MM-DD    Override start date for data fetch
  --full-refresh        Ignore existing data; re-fetch from start_date
  --no-futures          Skip Groww futures scraping
  --version             Show version and exit
```

---

## Excel Layout

### Headers

| Columns | Color | Fields |
|---------|-------|--------|
| A | Plain | Day (formula) |
| B – H | Blue | Date, Open, High, Low, Close, Daily Change, Weekly Change |
| I – L | Yellow | Close vs prev High, Close vs prev Low, HH/HL, LL/LH |
| M – O | Plain | Futures difference, price, expiry |

### Formatting

- **Tuesday rows**: light highlight on columns A–H
- **Daily / Weekly Change**: conditional green (positive) / red (negative)
- **Futures Difference**: conditional green / red
- **Weekly Change**: holiday-aware — finds the closest trading day to 7 calendar days back
- **Freeze Panes**: header row always visible

---

## Configuration

Edit `config/config.json`:

```json
{
    "ticker": "^NSEI",
    "excel_file": "data/stock_market_NSE50.xlsx",
    "start_date": "2026-01-01",
    "log_level": "INFO",
    "max_retries": 3,
    "request_timeout": 15,
    "open_excel_after_run": true,
    "backup_before_update": true
}
```

| Key | Description |
|-----|-------------|
| `ticker` | Yahoo Finance symbol (`^NSEI`, `^NSEBANK`, `^BSESN`) |
| `excel_file` | Output Excel path (relative to project root) |
| `start_date` | Earliest date to fetch (YYYY-MM-DD). Defaults to Jan 1 of current year. |
| `log_level` | DEBUG / INFO / WARNING / ERROR / CRITICAL |
| `max_retries` | HTTP retry attempts before giving up |
| `request_timeout` | HTTP timeout in seconds |
| `open_excel_after_run` | Auto-open Excel after update |
| `backup_before_update` | Archive Excel before each write |

---

## Requirements

For the **installed/portable .exe**: just Windows 10/11 and an internet connection. No Python needed.

For **developers** building from source:

- **Windows 10/11**
- **Python 3.10+** — [python.org](https://www.python.org/downloads/)
- **Internet** connection
- **Microsoft Excel** (for viewing output)

Python packages (installed automatically by `setup.bat`):

```
yfinance >= 0.2.31
pandas >= 2.0.0
openpyxl >= 3.1.0
requests >= 2.31.0
```

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `python` not found | Install Python 3.10+ and tick "Add to PATH" |
| venv creation fails | Run terminal as Administrator |
| No data fetched | Market closed or Yahoo Finance unavailable |
| Futures missing | Groww.in may be down — data still saved without futures |
| Excel locked | Close Excel before running |
| Duplicate warning | Normal — dates already in the sheet are skipped |

Check `logs/nse_fetcher.log` for detailed diagnostics.

---

## License

Provided as-is for personal and educational use.

---

## Building from Source

### Prerequisites
- Python 3.10+ with `pip`
- (Optional) [Inno Setup 6](https://jrsoftware.org/isinfo.php) for building the installer

### Steps

```bat
git clone https://github.com/Uddhav07/NSE_Data_Fetcher_v3.git
cd NSE_Data_Fetcher_v3
setup.bat                   # Creates venv, installs runtime deps
build.bat                   # Installs PyInstaller, builds exe + installer
```

**Output:**
- `dist\NSE Data Fetcher\NSE Data Fetcher.exe` — standalone exe (folder mode)
- `dist\installer\NSE_Data_Fetcher_Setup.exe` — Windows installer (if Inno Setup is available)

### Running in dev mode (no build)

```bat
setup.bat
venv\Scripts\python.exe app.py          # GUI mode
venv\Scripts\python.exe -m scripts.main  # CLI mode
```

## Version

**3.1.0** — March 2026
