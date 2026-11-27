#!/bin/bash
# Prints the numbers docs/report/build.js repeats (its STATS block): versions,
# code size and test counts. Run after scripts/check_all.sh passes, and copy
# anything that changed into STATS before rebuilding the report.
#
# Usage: scripts/report_stats.sh [check_all output file]
set -uo pipefail
cd "$(dirname "$0")/.."

size() { # label, dirs...
  local label=$1
  shift
  local files
  files=$(git ls-files "$@" | grep '\.dart$')
  printf '%-10s %4s files %7s lines\n' "$label" "$(echo "$files" | wc -l)" "$(echo "$files" | xargs cat | wc -l)"
}

echo "patient $(sed -n 's/^version: //p' app/pubspec.yaml)   doctor $(sed -n 's/^version: //p' doctor/pubspec.yaml)"
size app app/lib
size doctor doctor/lib
size server server/lib server/bin
size protocol packages/lifetrack_protocol/lib
files=$(git ls-files | grep '\.dart$' | grep -E '(^|/)test/')
printf '%-10s %4s files %7s lines\n' tests "$(echo "$files" | wc -l)" "$(echo "$files" | xargs cat | wc -l)"

if [ $# -ge 1 ]; then
  echo 'tests passed, in package order (app, doctor, server, protocol, integration):'
  grep -E 'All tests passed' "$1"
fi
