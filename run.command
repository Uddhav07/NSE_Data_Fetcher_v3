#!/usr/bin/env bash
# Double-click in Finder to update the Excel file (run after 15:35 IST, Excel closed).
cd "$(dirname "$0")" || exit 1
./run.sh
echo
read -n 1 -s -r -p "Done. Press any key to close..."
