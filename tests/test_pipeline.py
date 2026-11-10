"""run_fetch() orchestration with every network source mocked.

Covers the contract from CLAUDE.md: run_fetch returns a FetchResult and never exits;
futures failure never blocks OHLC; the live session skips today's row and futures;
--full-refresh backs up first; flags switch sources off.
"""
from __future__ import annotations

import json

import pytest
from openpyxl import load_workbook

import scripts.excel_writer as excel_writer
import scripts.main as main

from conftest import make_ohlc

DATES = ["2026-01-05", "2026-01-06", "2026-01-07", "2026-01-08", "2026-01-09"]


class Calls:
    def __init__(self):
        self.ohlc = self.futures = self.rsi = self.opened = 0


@pytest.fixture
def env(tmp_path, monkeypatch):
    """Isolated config/workbook/archive and mocked sources. Returns (config_path, xlsx, calls, archive)."""
    xlsx = tmp_path / "data" / "nse.xlsx"
    archive = tmp_path / "archive"
    cfg = tmp_path / "config.json"
    cfg.write_text(json.dumps({"ticker": "^NSEI", "excel_file": str(xlsx), "start_date": "2026-01-01",
                               "max_retries": 1, "open_excel_after_run": False}))
    calls = Calls()
    state = {"live": False, "ohlc": make_ohlc(DATES), "futures": (25500.5, "NIFTY Jan 2026 Fut"),
             "rsi": {d: {5: 50.0, 60: 60.0} for d in DATES}}

    def fake_ohlc(config, start):
        calls.ohlc += 1
        return state["ohlc"]

    def fake_futures(config):
        calls.futures += 1
        return state["futures"]

    def fake_rsi(config):
        calls.rsi += 1
        return state["rsi"]

    def fake_open(path):
        calls.opened += 1

    monkeypatch.setattr(main, "fetch_ohlc_data", fake_ohlc)
    monkeypatch.setattr(main, "fetch_futures", fake_futures)
    monkeypatch.setattr(main, "fetch_rsi", fake_rsi)
    monkeypatch.setattr(main, "_open_excel", fake_open)
    monkeypatch.setattr(main, "_check_environment", lambda: [])         # no network ping
    monkeypatch.setattr(main, "ensure_dirs", lambda: None)              # don't touch the project
    monkeypatch.setattr(main, "session_in_progress", lambda: state["live"])
    monkeypatch.setattr(main, "today_str", lambda: "2026-01-09")
    monkeypatch.setattr(excel_writer, "get_archive_dir", lambda: archive)
    return str(cfg), xlsx, calls, archive, state


def _dates(xlsx):
    ws = load_workbook(xlsx).active
    return [ws.cell(r, 2).value.strftime("%Y-%m-%d") for r in range(2, ws.max_row + 1)]


def test_after_close_writes_rows_futures_rsi(env):
    cfg, xlsx, calls, archive, _ = env
    res = main.run_fetch(config_path=cfg)
    assert isinstance(res, main.FetchResult) and res.ok and res.added == 5
    assert res.excel_path == str(xlsx)
    assert _dates(xlsx) == DATES
    ws = load_workbook(xlsx).active
    assert ws["U6"].value == 25500.5 and ws["V6"].value == "NIFTY Jan 2026 Fut"
    assert ws["O2"].value == 50.0 and ws["S6"].value == 60.0
    assert (calls.ohlc, calls.futures, calls.rsi, calls.opened) == (1, 1, 1, 0)
    assert not archive.exists()            # nothing to back up on the very first run


def test_second_run_backs_up_and_is_idempotent(env):
    cfg, xlsx, calls, archive, _ = env
    main.run_fetch(config_path=cfg)
    res = main.run_fetch(config_path=cfg)
    assert res.ok and res.added == 0 and res.skipped == 5
    assert len(list(archive.iterdir())) == 1


def test_live_session_skips_today_and_futures(env):
    cfg, xlsx, calls, _, state = env
    state["live"] = True
    res = main.run_fetch(config_path=cfg)
    assert res.ok and res.added == 4
    assert _dates(xlsx) == DATES[:-1]
    assert calls.futures == 0
    assert load_workbook(xlsx).active["U5"].value is None


def test_futures_failure_does_not_block_ohlc(env):
    cfg, xlsx, _, _, state = env
    state["futures"] = (None, None)
    res = main.run_fetch(config_path=cfg)
    assert res.ok and res.added == 5
    ws = load_workbook(xlsx).active
    assert ws["U6"].value is None and ws["F6"].value is not None


def test_no_ohlc_data_is_not_an_error(env):
    cfg, xlsx, _, _, state = env
    state["ohlc"] = None
    res = main.run_fetch(config_path=cfg)
    assert res.ok and res.added == 0
    assert xlsx.exists()                   # empty workbook with headers is still created


def test_flags_switch_sources_off(env):
    cfg, xlsx, calls, _, _ = env
    res = main.run_fetch(config_path=cfg, no_futures=True, no_rsi=True)
    assert res.ok and calls.futures == 0 and calls.rsi == 0
    ws = load_workbook(xlsx).active
    assert ws["U6"].value is None and ws["O2"].value is None


def test_full_refresh_backs_up_then_recreates(env):
    cfg, xlsx, _, archive, _ = env
    main.run_fetch(config_path=cfg)
    wb = load_workbook(xlsx); wb.active["W2"] = "note"; wb.save(xlsx)
    res = main.run_fetch(config_path=cfg, full_refresh=True)
    assert res.ok and res.added == 5
    assert load_workbook(xlsx).active["W2"].value is None          # recreated
    backups = list(archive.iterdir())
    assert backups and any(load_workbook(b).active["W2"].value == "note" for b in backups)


def test_start_override_is_passed_to_fetch(env, monkeypatch):
    cfg, _, _, _, state = env
    seen = []
    monkeypatch.setattr(main, "fetch_ohlc_data", lambda config, start: seen.append(start) or state["ohlc"])
    main.run_fetch(config_path=cfg, start_override="2025-12-01")
    assert seen == ["2025-12-01"]


def test_locked_workbook_returns_error_not_exit(env, monkeypatch):
    cfg, _, _, _, _ = env

    def locked(*a, **k):
        raise PermissionError("Cannot save — the file is locked.")

    monkeypatch.setattr(main, "update_workbook", locked)
    res = main.run_fetch(config_path=cfg)           # must return, not sys.exit
    assert not res.ok and "locked" in res.error


def test_open_flag_overrides_config(env):
    cfg, _, calls, _, _ = env
    main.run_fetch(config_path=cfg, open_excel=True)
    assert calls.opened == 1
    main.run_fetch(config_path=cfg, open_excel=False)
    assert calls.opened == 1
