#!/usr/bin/env bash
set -euo pipefail
cd /workspace

if ! python3 -c 'import venv' 2>/dev/null; then
  sudo apt-get update -qq
  sudo DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends python3.12-venv
fi

if [[ ! -d .venv ]]; then
  python3 -m venv .venv
fi

# shellcheck source=/dev/null
source .venv/bin/activate
pip install -r requirements.txt

export PLAYWRIGHT_BROWSERS_PATH=0
python -m playwright install chromium

# Browser libraries (fonts, libnss3, etc.) — idempotent on Ubuntu.
sudo .venv/bin/python -m playwright install-deps chromium
