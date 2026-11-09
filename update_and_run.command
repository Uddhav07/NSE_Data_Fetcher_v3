#!/usr/bin/env bash
# Double-click in Finder: update to the latest version, refresh the data, open Excel.
cd "$(dirname "$0")" || exit 1
./update_and_run.sh
echo
read -n 1 -s -r -p "Done. Press any key to close..."
