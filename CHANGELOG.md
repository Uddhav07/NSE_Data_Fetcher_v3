# Changelog

## 3.3.1
### Added
- README: complete Windows guide (install Python/Git incl. `winget`, choose-your-folder, setup, daily use, update, venv activation, Task Scheduler, troubleshooting).
- `.gitattributes` so `.bat` stays CRLF and `.sh`/`.command` stay LF on any OS.
### Changed
- Console output never crashes on legacy Windows code pages (unencodable characters are replaced).
- `setup.bat` falls back to the `py -3` launcher and reports venv-creation failures.

## 3.3.0
### Added
- **Self-healing dependencies** (`scripts/bootstrap.py`): a missing package (e.g. `No module named pandas`) is installed into the project `venv/` automatically and the run restarts there — works from system Python, IDEs, any interpreter.
- **Gap backfill:** every run checks from `start_date`, inserts any missing trading days in date order (formulas re-pointed), and fills empty cells. Cells that already have values are never changed.
- `requirements-optional.txt` (`pyswisseph`): if it can't install, only the Nakshatra column is skipped instead of the whole run failing.
### Changed
- Before 15:35 IST on a trading day, today's row and futures are skipped (not final); run again after close.
- README no longer assumes a folder — you choose where to put the project (and `excel_file` can be any path).

## 3.2.3
### Changed
- README: full copy-paste Windows (PowerShell/cmd) instructions with venv support.

## 3.2.2
### Added
- `setup.bat` / `run.bat` for Windows (double-click; untested on real Windows).
- `setup.command` and `run.command`: double-click in Finder to run setup / update in Terminal (a `.sh` opens in an editor).

## 3.2.1
### Changed
- RSI colours (cell-only): below 30 red, 70–79.99 light green, 80+ dark green (previously orange / red).

## 3.2.0 — Script edition (macOS/Linux)
### Added
- Column Q **Date Total** (`9, 3` style), R–U **RSI(14)** for 5/15/30/45-min candles at the 15:30 IST close, V **Moon Nakshatra** at 09:15 IST.
- Cell-only colour rules: RSI 70–79.99 orange, 80+ red; Date Total `3, 6` / `6, 3` purple.
- `--no-rsi` flag; `setup.sh` / `run.sh`; pytest suite (`tests/`); `requirements-dev.txt`.
- Sheet opens scrolled to the latest date.
### Changed
- **Edit-safe updates:** only new rows are added and empty cells filled; user values, fills, headers and conditional formats are never overwritten. Tuesday fill applies to new rows only; futures written once per row.
- Futures and RSI are refreshed even when there are no new OHLC rows.
- Last-date detection scans column B (a stray lower row no longer breaks it); fixed Tuesday RGB instead of a theme colour.
- Runs on Python 3.9+; opens Excel with `open`/`xdg-open` instead of Windows-only `os.startfile`.
### Removed
- Windows GUI (`gui.py`, `app.py`), PyInstaller spec, Inno Setup installer, `.bat` scripts, icon assets.

## 3.1.0 — Windows desktop app (tkinter, PyInstaller, installer)
## 3.0.0 — Modular rewrite of v2.1 (retries, logging, backups, futures fix)
