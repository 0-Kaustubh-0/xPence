#!/bin/bash
# xPence Report Generator — launcher for macOS / Linux
set -e
cd "$(dirname "$0")"

# ── Find Python ───────────────────────────────────────────────────────────────
PYTHON=$(command -v python3 2>/dev/null || command -v python 2>/dev/null)
if [ -z "$PYTHON" ]; then
    echo "Python not found. Install from https://www.python.org"
    exit 1
fi

# ── Install / verify dependencies ────────────────────────────────────────────
$PYTHON -c "import pandas"   2>/dev/null || $PYTHON -m pip install pandas
$PYTHON -c "import openpyxl" 2>/dev/null || $PYTHON -m pip install openpyxl

# ── Check tkinter (not pip-installable on all platforms) ─────────────────────
$PYTHON -c "import tkinter" 2>/dev/null || {
    echo ""
    echo "tkinter is missing. Install it with:"
    echo "  macOS : brew install python-tk"
    echo "  Ubuntu: sudo apt-get install python3-tk"
    echo "  Fedora: sudo dnf install python3-tkinter"
    exit 1
}

# ── Launch GUI ────────────────────────────────────────────────────────────────
exec $PYTHON xpence_gui.py
