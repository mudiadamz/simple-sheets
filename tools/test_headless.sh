#!/bin/sh
# Run the test suite in a Linux container on a virtual display (Xvfb): no test window on your screen,
# and the keyboard / mouse stay yours. Needs Docker (on macOS: `brew install colima docker`).
#   tools/test_headless.sh                      all tests
#   tools/test_headless.sh tests/test_grid.py -k wheel
# The code is mounted, not copied: it tests what is on disk now, with ../simple-core (or SIMPLE_CORE=path).
set -e
cd "$(dirname "$0")/.."
CORE=${SIMPLE_CORE:-../simple-core}
[ -d "$CORE/simplecore" ] || { echo "simple-core not found at $CORE (set SIMPLE_CORE)"; exit 2; }
CORE=$(cd "$CORE" && pwd)
if ! docker info >/dev/null 2>&1; then
  if command -v colima >/dev/null 2>&1; then
    echo "Starting Docker (colima)..."
    colima start >/dev/null
  else
    echo "Docker isn't running"; exit 2
  fi
fi
# the image only holds Python + the test packages: rebuilt when the requirements change
TAG=simple-sheets-test:$(cat requirements.txt requirements-dev.txt tools/headless.Dockerfile | cksum | cut -d' ' -f1)
if ! docker image inspect "$TAG" >/dev/null 2>&1; then
  echo "Building the test image (once)..."
  docker build -q -t "$TAG" -f tools/headless.Dockerfile . >/dev/null
fi
# the repo is mounted read-only: test files go to /tmp inside the container
# (Xvfb started by hand: xvfb-run as the container's first process waits forever for Xvfb's ready signal)
exec docker run --rm --init -v "$PWD":/app:ro -v "$CORE":/core:ro "$TAG" sh -c '
  Xvfb :99 -screen 0 1600x1000x24 -nolisten tcp >/dev/null 2>&1 &
  for i in $(seq 50); do [ -e /tmp/.X11-unix/X99 ] && break; sleep 0.1; done
  DISPLAY=:99 exec python -m pytest -p no:cacheprovider "$@"' sh "$@"
