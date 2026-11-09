#!/usr/bin/env bash
# One-time setup: create venv and install dependencies.
set -euo pipefail
cd "$(dirname "$0")"
command -v python3 >/dev/null || { echo "python3 not found. Install Python 3.9+ (brew install python)."; exit 1; }
python3 -c 'import sys; sys.exit(sys.version_info < (3,9))' || { echo "Python 3.9+ required."; exit 1; }
python3 -m venv venv
venv/bin/pip install --quiet --upgrade pip
venv/bin/pip install --quiet -r requirements.txt
venv/bin/pip install --quiet -r requirements-optional.txt || echo "Note: optional package not installed - Moon Nakshatra column will stay blank."
echo "Setup complete. Run ./run.sh to update data."
