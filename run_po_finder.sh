#!/bin/bash

# ── Find Python 3 ─────────────────────────────────────────────────────────────
PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" &>/dev/null; then
        VER=$("$cmd" -c "import sys; print(sys.version_info.major)")
        if [ "$VER" = "3" ]; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    osascript -e 'display dialog "Python 3 is not installed.\nPlease download it from https://www.python.org/downloads/" buttons {"OK"} with icon stop' 2>/dev/null \
    || echo "[ERROR] Python 3 not found. Install from https://www.python.org/downloads/"
    exit 1
fi

# ── Install dependencies if missing ──────────────────────────────────────────
echo "Checking dependencies..."
$PYTHON -c "import pandas" 2>/dev/null || $PYTHON -m pip install pandas --quiet
$PYTHON -c "import openpyxl" 2>/dev/null || $PYTHON -m pip install openpyxl --quiet

# ── Resolve script location ───────────────────────────────────────────────────
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# ── Launch (detached, no terminal window on Mac) ──────────────────────────────
if [[ "$OSTYPE" == "darwin"* ]]; then
    # macOS: open detached so no terminal stays open
    nohup $PYTHON "$SCRIPT_DIR/po_xml_finder.py" >/dev/null 2>&1 &
else
    # Linux: try to hide terminal, fall back to running directly
    if command -v nohup &>/dev/null; then
        nohup $PYTHON "$SCRIPT_DIR/po_xml_finder.py" >/dev/null 2>&1 &
    else
        $PYTHON "$SCRIPT_DIR/po_xml_finder.py" &
    fi
fi

echo "App launched!"