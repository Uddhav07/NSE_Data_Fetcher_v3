#!/usr/bin/env bash
# 1) pull the latest version from GitHub, 2) update the Excel data, 3) open the workbook.
# Extra flags are passed through, e.g. ./update_and_run.sh --no-futures
# Your local edits are never overwritten: if tracked files were modified, the update step is skipped.
main() {
  cd "$(dirname "$0")"

  if [ -d .git ] && command -v git >/dev/null 2>&1; then
    if git diff --quiet HEAD 2>/dev/null; then
      echo "Checking GitHub for a newer version..."
      if git pull --ff-only; then
        :
      else
        echo "Could not update (offline, or your copy has diverged) - running the current version."
      fi
    else
      echo "You have local changes to project files - skipping the update so nothing is overwritten."
    fi
  else
    echo "Not a git clone (or git not installed) - skipping the update step."
  fi

  [ -x venv/bin/python ] || ./setup.sh
  # keep packages in sync with the (possibly new) requirements; quick no-op when already satisfied
  venv/bin/python -m pip install --disable-pip-version-check -q -r requirements.txt || \
    echo "Note: could not refresh packages - continuing."

  ./run.sh --open "$@"
}
main "$@"
exit $?
