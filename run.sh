#!/usr/bin/env bash
# Update the Excel file. Extra flags are passed through (see README).
set -euo pipefail
cd "$(dirname "$0")"
[ -x venv/bin/python ] || ./setup.sh
exec venv/bin/python -m scripts.main "$@"
