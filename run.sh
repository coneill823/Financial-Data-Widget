#!/usr/bin/env bash
# Install dependencies (first run only) then launch the dashboard.
set -e
pip install -q -r "$(dirname "$0")/requirements.txt"
python "$(dirname "$0")/financial_dashboard.py"
