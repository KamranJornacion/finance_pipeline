#!/usr/bin/env bash
# Double-click launcher (Linux/Mac). Sets up Python the first time, then builds the report.
cd "$(dirname "$0")" || exit 1

pause() { echo; read -r -p "Press Enter to close this window..." _; }

PY=""
for c in python3 python; do
  if command -v "$c" >/dev/null 2>&1 && "$c" -c 'import sys; sys.exit(sys.version_info < (3, 10))'; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "Python 3.10 or newer is needed. Install it from https://www.python.org/downloads/ and try again."
  pause; exit 1
fi

if [ ! -x .venv/bin/python ]; then
  echo "First-time setup: installing what's needed (this takes a minute)..."
  "$PY" -m venv .venv || { echo "Could not create the Python environment."; pause; exit 1; }
fi
# reinstall only when requirements.txt changes
if ! cmp -s requirements.txt .venv/.installed-requirements 2>/dev/null; then
  .venv/bin/python -m pip install --quiet --disable-pip-version-check -r requirements.txt \
    || { echo "Could not install the required packages (check your internet connection)."; pause; exit 1; }
  cp requirements.txt .venv/.installed-requirements
fi

.venv/bin/python -m finpipe "$@"
status=$?
if [ $status -ne 0 ]; then
  echo; echo "Finished with problems - please read the messages above."
  pause
else
  sleep 3
fi
exit $status
