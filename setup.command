#!/usr/bin/env bash
# Double-click in Finder to run the one-time setup in Terminal.
cd "$(dirname "$0")" || exit 1
./setup.sh
echo
read -n 1 -s -r -p "Press any key to close..."
