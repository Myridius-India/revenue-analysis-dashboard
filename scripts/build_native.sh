#!/usr/bin/env bash
set -euo pipefail

NAME="${1:-RevenueAnalysisDesktop}"
PYTHON_EXE="${PYTHON_EXE:-$(pwd)/.venv/bin/python}"

if [[ ! -x "$PYTHON_EXE" ]]; then
	echo "Python executable not found: $PYTHON_EXE"
	echo "Set PYTHON_EXE to a valid interpreter path."
	exit 1
fi

"$PYTHON_EXE" -m PyInstaller --noconfirm --windowed --name "$NAME" native_app/main.py
