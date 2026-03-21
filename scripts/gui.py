"""Desktop GUI for NSE Data Fetcher v3 — tkinter + ttk."""

import logging
import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
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

logger = logging.getLogger(__name__)


# ── Logging handler that writes to a tkinter Text widget ──────────────────────

class _TextWidgetHandler(logging.Handler):
    """Thread-safe logging handler that appends messages to a ScrolledText."""

    def __init__(self, text_widget: scrolledtext.ScrolledText) -> None:
        super().__init__()
        self._widget = text_widget

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record) + "\n"
        # Schedule on the main thread via after()
        self._widget.after(0, self._append, msg)

    def _append(self, msg: str) -> None:
        self._widget.configure(state="normal")
        self._widget.insert(tk.END, msg)
        self._widget.see(tk.END)
        self._widget.configure(state="disabled")


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
        cfg_path = get_config_path()
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            data = {}

        for attr, var in self._vars.items():
            data[attr] = var.get()

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

        self._build_ui()
        self._init_app()

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

        ttk.Button(bot, text="⚙  Settings", command=self._on_settings).pack(side="left", padx=4)
        ttk.Button(bot, text="📁  Data Folder", command=self._on_open_folder).pack(side="left", padx=4)
        ttk.Button(bot, text="ℹ  About", command=self._on_about).pack(side="left", padx=4)
        ttk.Button(bot, text="Full Refresh", command=self._on_full_refresh).pack(side="right", padx=4)

    # ── Initialization ────────────────────────────────────────────────────

    def _init_app(self) -> None:
        ensure_dirs()
        save_default_config()
        self._config = load_config()

        # Set up logging to file + GUI text widget
        handler = _TextWidgetHandler(self._log_text)
        setup_logging(self._config, extra_handler=handler)

        self._status_var.set("Ready — click Update Data to fetch latest prices.")

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
        self._btn_update.configure(state="disabled")
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
        self._btn_update.configure(state="normal")

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
        if self._config is None:
            return
        _SettingsDialog(self, self._config)
        # Reload config after dialog closes
        self._config = load_config()

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


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
