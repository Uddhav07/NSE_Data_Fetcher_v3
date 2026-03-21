"""Desktop GUI for NSE Data Fetcher v3 — tkinter + ttk."""

import logging
import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
from typing import Optional

from . import __version__
from .config_manager import AppConfig, load_config, save_default_config
from .paths import (
    ensure_dirs,
    get_base_dir,
    get_config_path,
    resolve_relative,
)
from .main import FetchResult, run_fetch, setup_logging
from .excel_writer import get_last_date

logger = logging.getLogger(__name__)


# ── Logging handler that writes to a tkinter Text widget ──────────────────────

class _TextWidgetHandler(logging.Handler):
    """Thread-safe logging handler that appends messages to a ScrolledText."""

    def __init__(self, text_widget: scrolledtext.ScrolledText) -> None:
        super().__init__()
        self._widget = text_widget
        self._closed = False

    def emit(self, record: logging.LogRecord) -> None:
        if self._closed:
            return
        msg = self.format(record) + "\n"
        try:
            self._widget.after(0, self._append, msg)
        except (tk.TclError, RuntimeError):
            # Widget was destroyed (app closing) — silently ignore
            self._closed = True

    def _append(self, msg: str) -> None:
        try:
            self._widget.configure(state="normal")
            self._widget.insert(tk.END, msg)
            self._widget.see(tk.END)
            self._widget.configure(state="disabled")
        except tk.TclError:
            self._closed = True


# ── Settings dialog ───────────────────────────────────────────────────────────

class _SettingsDialog(tk.Toplevel):
    """Modal dialog for editing config.json fields."""

    def __init__(self, parent: tk.Tk, config: AppConfig) -> None:
        super().__init__(parent)
        self.title("Settings")
        self.resizable(False, False)
        self.grab_set()
        self.transient(parent)

        self._config = config
        self._vars: dict[str, tk.Variable] = {}

        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)

        fields = [
            ("Ticker", "ticker", "str"),
            ("Excel file", "excel_file", "str"),
            ("Start date (YYYY-MM-DD)", "start_date", "str"),
            ("Log level", "log_level", "str"),
            ("Max retries", "max_retries", "int"),
            ("Request timeout (s)", "request_timeout", "int"),
            ("Open Excel after run", "open_excel_after_run", "bool"),
            ("Backup before update", "backup_before_update", "bool"),
        ]

        for row, (label, attr, kind) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=3)
            val = getattr(config, attr)
            if kind == "bool":
                var = tk.BooleanVar(value=val)
                ttk.Checkbutton(frame, variable=var).grid(row=row, column=1, sticky="w", padx=(8, 0))
            elif kind == "int":
                var = tk.IntVar(value=val)
                ttk.Entry(frame, textvariable=var, width=8).grid(row=row, column=1, sticky="w", padx=(8, 0))
            else:
                var = tk.StringVar(value=val)
                ttk.Entry(frame, textvariable=var, width=36).grid(row=row, column=1, sticky="w", padx=(8, 0))
            self._vars[attr] = var

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=(12, 0))
        ttk.Button(btn_frame, text="Save", command=self._save).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side="left", padx=4)

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_reqwidth()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_reqheight()) // 2
        self.geometry(f"+{x}+{y}")

    def _save(self) -> None:
        import json

        # ── Validate before saving ────────────────────────────────────
        errors: list[str] = []

        ticker = self._vars["ticker"].get().strip()
        if not ticker:
            errors.append("Ticker cannot be empty.")

        start_date = self._vars["start_date"].get().strip()
        try:
            datetime.strptime(start_date, "%Y-%m-%d")
        except ValueError:
            errors.append(f"Invalid start date: '{start_date}'. Use YYYY-MM-DD format.")

        log_level = self._vars["log_level"].get().strip().upper()
        if log_level not in {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}:
            errors.append(f"Invalid log level: '{log_level}'.")

        try:
            retries = self._vars["max_retries"].get()
            if retries < 0:
                errors.append("Max retries must be >= 0.")
        except (tk.TclError, ValueError):
            errors.append("Max retries must be a number.")

        try:
            timeout = self._vars["request_timeout"].get()
            if timeout < 1:
                errors.append("Request timeout must be >= 1 second.")
        except (tk.TclError, ValueError):
            errors.append("Request timeout must be a number.")

        excel_file = self._vars["excel_file"].get().strip()
        if not excel_file:
            errors.append("Excel file path cannot be empty.")

        if errors:
            messagebox.showwarning(
                "Validation Error",
                "Please fix the following:\n\n• " + "\n• ".join(errors),
                parent=self,
            )
            return

        # ── Write to file ─────────────────────────────────────────────
        cfg_path = get_config_path()
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            data = {}

        for attr, var in self._vars.items():
            try:
                data[attr] = var.get()
            except (tk.TclError, ValueError):
                pass  # keep existing value if widget has bad data

        cfg_path.parent.mkdir(parents=True, exist_ok=True)
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

        messagebox.showinfo("Settings", "Configuration saved.", parent=self)
        self.destroy()


# ── Main application window ──────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title(f"NSE Data Fetcher v{__version__}")
        self.geometry("720x520")
        self.minsize(600, 420)

        # Try to load icon
        self._set_icon()

        self._running = False
        self._config: Optional[AppConfig] = None
        self._action_buttons: list[ttk.Button] = []  # buttons to disable during fetch

        self._build_ui()
        self._init_app()

        # Intercept window close to warn if fetch is running
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── Icon ──────────────────────────────────────────────────────────────

    def _set_icon(self) -> None:
        ico = get_base_dir() / "assets" / "icon.ico"
        if ico.exists():
            try:
                self.iconbitmap(str(ico))
            except tk.TclError:
                pass

    # ── UI layout ─────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        style.theme_use("vista" if "vista" in style.theme_names() else "clam")

        # Top frame — status + buttons
        top = ttk.Frame(self, padding=(12, 8))
        top.pack(fill="x")

        self._status_var = tk.StringVar(value="Initializing…")
        ttk.Label(top, textvariable=self._status_var,
                  font=("Segoe UI", 10)).pack(side="left")

        # Right-side buttons
        btn_frame = ttk.Frame(top)
        btn_frame.pack(side="right")

        self._btn_update = ttk.Button(btn_frame, text="⟳  Update Data",
                                       command=self._on_update)
        self._btn_update.pack(side="left", padx=4)

        self._btn_open = ttk.Button(btn_frame, text="📂  Open Excel",
                                     command=self._on_open_excel)
        self._btn_open.pack(side="left", padx=4)

        # Progress bar
        self._progress = ttk.Progressbar(self, mode="indeterminate")
        self._progress.pack(fill="x", padx=12, pady=(0, 4))

        # Log output
        self._log_text = scrolledtext.ScrolledText(
            self, height=18, state="disabled",
            font=("Consolas", 9), wrap="word",
            background="#1e1e1e", foreground="#d4d4d4",
            insertbackground="#d4d4d4",
        )
        self._log_text.pack(fill="both", expand=True, padx=12, pady=(0, 4))

        # Bottom bar
        bot = ttk.Frame(self, padding=(12, 6))
        bot.pack(fill="x")

        self._btn_settings = ttk.Button(bot, text="⚙  Settings",
                                         command=self._on_settings)
        self._btn_settings.pack(side="left", padx=4)
        self._btn_folder = ttk.Button(bot, text="📁  Data Folder",
                                       command=self._on_open_folder)
        self._btn_folder.pack(side="left", padx=4)
        ttk.Button(bot, text="ℹ  About", command=self._on_about).pack(side="left", padx=4)
        self._btn_refresh = ttk.Button(bot, text="Full Refresh",
                                        command=self._on_full_refresh)
        self._btn_refresh.pack(side="right", padx=4)

        # Collect all buttons that should be disabled during a fetch
        self._action_buttons = [
            self._btn_update, self._btn_open, self._btn_settings,
            self._btn_folder, self._btn_refresh,
        ]

    # ── Initialization ────────────────────────────────────────────────────

    def _init_app(self) -> None:
        ensure_dirs()
        save_default_config()

        try:
            self._config = load_config()
        except (ValueError, Exception) as exc:
            # Config is corrupt — reset to defaults and warn user
            logger.warning("Config error: %s — resetting to defaults.", exc)
            cfg_path = get_config_path()
            try:
                cfg_path.unlink(missing_ok=True)
            except OSError:
                pass
            save_default_config()
            self._config = load_config()
            messagebox.showwarning(
                "Configuration Reset",
                f"Your settings file had errors and was reset to defaults.\n\n"
                f"Original error: {exc}",
            )

        # Set up logging to file + GUI text widget
        handler = _TextWidgetHandler(self._log_text)
        setup_logging(self._config, extra_handler=handler)

        # Show last-updated date if Excel file exists
        status = "Ready — click Update Data to fetch latest prices."
        try:
            excel_path = str(resolve_relative(self._config.excel_file))
            if os.path.exists(excel_path):
                last = get_last_date(excel_path)
                if last is not None:
                    status = f"Last updated: {last.strftime('%Y-%m-%d')} — click Update Data for latest."
        except Exception:
            pass  # don't crash on startup for a status message

        self._status_var.set(status)

    # ── Button handlers ───────────────────────────────────────────────────

    def _on_update(self) -> None:
        self._start_fetch(full_refresh=False)

    def _on_full_refresh(self) -> None:
        if not messagebox.askyesno(
            "Full Refresh",
            "This will re-download ALL data from the start date, "
            "replacing the existing Excel file.\n\nContinue?",
        ):
            return
        self._start_fetch(full_refresh=True)

    def _start_fetch(self, full_refresh: bool) -> None:
        if self._running:
            return
        self._running = True
        for btn in self._action_buttons:
            btn.configure(state="disabled")
        self._progress.start(12)
        self._status_var.set("Fetching data — please wait…")

        thread = threading.Thread(
            target=self._fetch_worker,
            args=(full_refresh,),
            daemon=True,
        )
        thread.start()

    def _fetch_worker(self, full_refresh: bool) -> None:
        try:
            result = run_fetch(
                full_refresh=full_refresh,
                open_excel=False,  # don't auto-open from worker thread
            )
        except Exception as exc:
            result = FetchResult(error=str(exc))
        # Schedule UI update on main thread
        self.after(0, self._on_fetch_done, result)

    def _on_fetch_done(self, result: FetchResult) -> None:
        self._progress.stop()
        self._running = False
        for btn in self._action_buttons:
            btn.configure(state="normal")

        if result.ok:
            if result.added:
                self._status_var.set(
                    f"Done — {result.added} new row(s) added, "
                    f"{result.skipped} skipped."
                )
            else:
                self._status_var.set("Up to date — no new data.")
        else:
            self._status_var.set(f"Error: {result.error}")
            messagebox.showerror("Fetch Error", result.error)

    def _on_open_excel(self) -> None:
        if self._config is None:
            return
        path = str(resolve_relative(self._config.excel_file))
        if os.path.exists(path):
            os.startfile(os.path.abspath(path))
        else:
            messagebox.showinfo(
                "No Data",
                "Excel file not found yet. Click Update Data first.",
            )

    def _on_settings(self) -> None:
        if self._config is None or self._running:
            return
        _SettingsDialog(self, self._config)
        # Reload config after dialog closes (safely — handle corrupt file)
        try:
            self._config = load_config()
        except (ValueError, Exception) as exc:
            messagebox.showwarning(
                "Config Error",
                f"Settings may not have saved correctly.\n\n{exc}",
            )

    def _on_open_folder(self) -> None:
        folder = str(resolve_relative("data"))
        if not os.path.isdir(folder):
            os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

    def _on_about(self) -> None:
        messagebox.showinfo(
            "About NSE Data Fetcher",
            f"NSE Data Fetcher v{__version__}\n\n"
            "Nifty 50 OHLC + Futures tracker with\n"
            "professional Excel output.\n\n"
            "github.com/Uddhav07/NSE_Data_Fetcher_v3",
        )

    def _on_close(self) -> None:
        """Handle window close — warn if fetch is in progress."""
        if self._running:
            if not messagebox.askyesno(
                "Fetch in Progress",
                "Data is still being downloaded.\n\n"
                "If you close now, the current fetch will be interrupted "
                "and partial data may not be saved.\n\nClose anyway?",
            ):
                return
        self.destroy()


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
