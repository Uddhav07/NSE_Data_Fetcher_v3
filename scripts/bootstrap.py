"""Self-healing dependency check — makes "No module named pandas" impossible.

If a required package is missing from the interpreter that started the script
(system Python, an IDE's interpreter, a different venv...), this module:
  1. creates the project venv (<project>/venv) if it doesn't exist,
  2. installs requirements.txt into it (and tries requirements-optional.txt),
  3. re-runs the same command inside that venv.
Uses the standard library only, so it can run before anything else is imported.
"""
from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VENV = ROOT / "venv"
REQUIRED = ("pandas", "yfinance", "openpyxl", "requests")
_ENV_FLAG = "NSE_FETCHER_BOOTSTRAPPED"


def _missing(names=REQUIRED) -> list:
    return [n for n in names if importlib.util.find_spec(n) is None]


def _venv_python() -> Path:
    if os.name == "nt":
        return VENV / "Scripts" / "python.exe"
    return VENV / "bin" / "python"


def _run(cmd) -> int:
    return subprocess.call([str(c) for c in cmd], cwd=str(ROOT))


def _venv_has_packages(py: Path) -> bool:
    code = "import " + ",".join(REQUIRED)
    return subprocess.call([str(py), "-c", code], cwd=str(ROOT),
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL) == 0


def ensure_dependencies() -> None:
    """Return normally if all packages import; otherwise fix the environment and re-run."""
    missing = _missing()
    if not missing:
        return

    if os.environ.get(_ENV_FLAG):
        print(f"[setup] Missing Python package(s): {', '.join(missing)}.")
        print("[setup] Automatic setup already ran but packages are still missing.\n"
              f"        Run manually:  {_venv_python()} -m pip install -r requirements.txt")
        sys.exit(1)

    py = _venv_python()
    if not py.exists() or not _venv_has_packages(py):
        print(f"[setup] Missing Python package(s): {', '.join(missing)}. Fixing automatically...",
              flush=True)
    if not py.exists():
        print(f"[setup] Creating virtual environment in {VENV} ...", flush=True)
        if _run([sys.executable, "-m", "venv", VENV]) != 0 or not py.exists():
            print("[setup] Could not create a venv. Install Python 3.9+ from python.org "
                  "(on Windows tick 'Add python.exe to PATH') and try again.")
            sys.exit(1)

    if not _venv_has_packages(py):
        print("[setup] Installing packages (first time only, about a minute)...", flush=True)
        base = [py, "-m", "pip", "install", "--disable-pip-version-check", "-q"]
        if _run(base + ["-r", ROOT / "requirements.txt"]) != 0:
            print("[setup] Package install failed. Check your internet connection and run again.")
            sys.exit(1)
        opt = ROOT / "requirements-optional.txt"
        if opt.exists() and _run(base + ["-r", opt]) != 0:
            print("[setup] Optional package(s) not installed (Moon Nakshatra column will stay "
                  "blank). Everything else works.")

    env = dict(os.environ, **{_ENV_FLAG: "1"})
    rc = subprocess.call([str(py), "-m", "scripts.main", *sys.argv[1:]], cwd=str(ROOT), env=env)
    sys.exit(rc)
