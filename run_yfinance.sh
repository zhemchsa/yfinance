#!/usr/bin/env bash
set -eu

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

if [ -d ".myenv3.14" ]; then
  VENV_DIR=".myenv3.14"
elif [ -d ".venv" ]; then
  VENV_DIR=".venv"
else
  python3 -m venv .venv
  VENV_DIR=".venv"
fi

. "$ROOT_DIR/$VENV_DIR/bin/activate"
python -m pip install --upgrade pip >/dev/null
pip install -r requirements.txt >/dev/null
python src/main.py "$@"
