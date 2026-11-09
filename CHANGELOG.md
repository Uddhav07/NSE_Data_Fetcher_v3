# Changelog

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
