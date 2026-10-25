#!/usr/bin/env bash
# PPT Unify - Quick Launcher
cd "$(dirname "$0")"
if [ -f ".venv/bin/python" ]; then
    .venv/bin/python run.py "$@"
else
    python3 run.py "$@"
fi
