#!/usr/bin/env bash
# Linux / macOS build: runs build_macos.sh (.app + .dmg) on macOS, build_linux.sh (binary + .deb) elsewhere.
# The installer (installer.py) is Windows-only.
set -euo pipefail
cd "$(dirname "$0")"
if [ "$(uname)" = "Darwin" ]; then
  exec sh ./build_macos.sh
else
  exec sh ./build_linux.sh
fi
