# NSE Data Fetcher v3

Automated **Nifty 50 Index** data fetcher with technical analysis, futures tracking, and professional Excel output.

> Rewritten from [v2.1](https://github.com/pmmittal-byte/NSE_Data_Fetcher_v2.1) with a modular architecture, proper logging, retry logic, and an improved Excel layout.

---

## What's New in v3

| Area | v2.1 | v3 |
|---|---|---|
| **Architecture** | Single 290-line file | 5 focused modules |
| **Logging** | `print()` statements | Rotating file + console via `logging` |
| **Retries** | None | Exponential back-off (configurable) |
| **Futures bug** | Same price stamped on *every* row | Futures only written for today's row |
| **Excel formatting** | Basic colors | Conditional green/red for changes, freeze panes |
| **CLI** | None | `--full-refresh`, `--start`, `--no-futures`, `--version` |
| **Config validation** | None | Dataclass-based with typed defaults |
| **Backups** | None | Timestamped archive before each update |
| **Data validation** | None | NaN/zero row filtering |
| **Type hints** | None | Throughout |

---

## Features

- **OHLC Data** from Yahoo Finance (`^NSEI`)
- **Futures Price & Expiry** from Groww.in (current month)
- **Technical Analysis** formulas — HH/HL, LL/LH, daily & weekly changes
- **Duplicate Prevention** — existing dates are never overwritten
- **Professional Excel** — color-coded headers, Monday highlighting, conditional red/green, frozen header row
- **Daily Automation** via Windows Task Scheduler
- **Rotating Logs** — 5 MB × 3 backups in `logs/`

---

## Project Structure

```
NSE_Data_Fetcher_v3/
├── scripts/
│   ├── __init__.py          # Package marker + version
│   ├── __main__.py          # python -m scripts entry
│   ├── main.py              # CLI, orchestration, logging setup
│   ├── config_manager.py    # Config loading and validation
│   ├── fetcher.py           # Yahoo Finance OHLC fetcher
│   ├── futures.py           # Groww futures scraper
│   └── excel_writer.py      # Workbook creation & update
├── config/
│   └── config.json          # User settings
├── data/                    # Excel output (auto-created)
├── logs/                    # Rotating log files (auto-created)
├── archive/                 # Timestamped backups (auto-created)
├── setup.bat                # One-time venv + deps setup
├── run_manual.bat           # Manual fetch (double-click)
├── run_scheduled.bat        # Silent scheduled run
├── schedule_task.bat        # Register Windows daily task
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Quick Start

### 1. Clone & Setup

```bat
git clone https://github.com/pmmittal-byte/NSE_Data_Fetcher_v3.git
cd NSE_Data_Fetcher_v3
setup.bat
```

This creates a Python **virtual environment** (`venv/`) and installs all dependencies.

### 2. Fetch Data

Double-click **`run_manual.bat`**, or from a terminal:

```bat
venv\Scripts\python.exe -m scripts.main
```

### 3. Automate (optional)

Double-click **`schedule_task.bat`** to register a daily 5:30 PM task.

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

- **Monday rows**: light-blue highlight on columns A–H
- **Daily / Weekly Change**: conditional green (positive) / red (negative)
- **Futures Difference**: conditional green / red
- **Freeze Panes**: header row always visible

---

## Configuration

Edit `config/config.json`:

```json
{
    "ticker": "^NSEI",
    "excel_file": "data/NSE50_Data.xlsx",
    "start_date": "2025-10-01",
    "schedule_time": "17:30",
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
| `start_date` | Earliest date to fetch (YYYY-MM-DD) |
| `schedule_time` | Daily auto-run time (24 h) |
| `log_level` | DEBUG / INFO / WARNING / ERROR / CRITICAL |
| `max_retries` | HTTP retry attempts before giving up |
| `request_timeout` | HTTP timeout in seconds |
| `backup_before_update` | Archive Excel before each write |

---

## Requirements

- **Windows 10/11**
- **Python 3.10+** — [python.org](https://www.python.org/downloads/)
- **Internet** connection
- **Microsoft Excel** (for viewing output)

Python packages (installed automatically by `setup.bat`):

```
yfinance >= 0.2.31
pandas >= 2.0.0
openpyxl >= 3.1.0
beautifulsoup4 >= 4.12.0
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

## Version

**3.0.0** — March 2026
