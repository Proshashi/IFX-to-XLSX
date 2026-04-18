#!/bin/bash
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
    echo "python3 not found. Install from https://www.python.org/downloads/macos/" >&2
    exit 1
fi

if ! python3 -c "import openpyxl" 2>/dev/null; then
    echo "Installing openpyxl..."
    python3 -m pip install --user openpyxl
fi

exec python3 ifx_to_xlsx_gui.py
