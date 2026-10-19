#!/bin/bash
# Social Engagement Processor — Mac launcher
# Double-click this file in Finder to start the app.
# If macOS blocks it, right-click → Open the first time.

cd "$(dirname "$0")"
echo "Starting Social Engagement Processor..."
echo

# Find Python 3
if command -v python3 &>/dev/null; then
    PYTHON=python3
elif command -v python &>/dev/null; then
    PYTHON=python
else
    echo "ERROR: Python 3 not found."
    echo "Install from https://www.python.org/downloads/"
    echo "Or via Homebrew: brew install python"
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Installing / verifying dependencies..."
$PYTHON -m pip install -r requirements.txt -q
echo
$PYTHON -m streamlit run app.py
