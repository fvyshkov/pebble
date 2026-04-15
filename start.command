#!/bin/bash
# Pebble — one-click launcher for macOS / Linux
# Double-click this file on Mac, or run: bash start.command

set -e
cd "$(dirname "$0")"
PORT=${PORT:-8000}

# ── Check / install Python ──
if ! command -v python3 &>/dev/null; then
    echo ""
    echo "  Python 3 not found. Installing..."
    echo ""
    if [[ "$OSTYPE" == "darwin"* ]]; then
        # macOS: install via Homebrew
        if ! command -v brew &>/dev/null; then
            echo "  Installing Homebrew first..."
            /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
            eval "$(/opt/homebrew/bin/brew shellenv 2>/dev/null || /usr/local/bin/brew shellenv 2>/dev/null)"
        fi
        brew install python@3.12
    else
        # Linux
        if command -v apt-get &>/dev/null; then
            sudo apt-get update -qq && sudo apt-get install -y -qq python3 python3-venv python3-pip
        elif command -v dnf &>/dev/null; then
            sudo dnf install -y python3 python3-pip
        elif command -v yum &>/dev/null; then
            sudo yum install -y python3 python3-pip
        else
            echo "  ERROR: Cannot install Python automatically."
            echo "  Please install Python 3.10+ manually."
            exit 1
        fi
    fi
fi

echo ""
echo "  Python: $(python3 --version)"

# ── Create venv ──
if [ ! -d ".venv" ]; then
    echo "  Creating virtual environment..."
    python3 -m venv .venv
fi
source .venv/bin/activate

# ── Install deps ──
echo "  Checking dependencies..."
pip install -q -r requirements.txt 2>/dev/null

# ── Open browser ──
(sleep 2 && open "http://localhost:$PORT" 2>/dev/null || xdg-open "http://localhost:$PORT" 2>/dev/null) &

# ── Start ──
echo ""
echo "  ========================================"
echo "   Pebble: http://localhost:$PORT"
echo "  ========================================"
echo ""
python -m uvicorn backend.main:app --host 0.0.0.0 --port "$PORT"
