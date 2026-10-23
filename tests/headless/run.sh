#!/bin/bash
# Runs a CalcMap test macro in LibreOffice without a display.
#
# Usage: tests/headless/run.sh <macro> [spreadsheet.ods]
#   macro       : Sub of tests_calcmap.bas (e.g. TestFullEditor)
#   spreadsheet : tests/test-departements.ods by default (copied, never changed)
#
# Builds a copy of the extension with the test macros added, installs it in
# a throwaway LibreOffice profile, runs the macro and prints the "TEST"
# lines of the log. Image of the map: $WORK/map.png.
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
PROJECT="$(cd "$HERE/../.." && pwd)"
MACRO="${1:?test macro expected}"
SPREADSHEET="${2:-$PROJECT/tests/test-departements.ods}"
WORK="${TMPDIR:-/tmp}/calcmap-tests"
PORT=$((2200 + RANDOM % 500))

rm -rf "$WORK/copy" "$WORK/profile"
mkdir -p "$WORK"
cp -r "$PROJECT" "$WORK/copy"
rm -rf "$WORK/copy/.git"
cat "$HERE/tests_calcmap.bas" >> "$WORK/copy/src/basic/CalcMap.bas"
(cd "$WORK/copy" && python3 build.py > /dev/null)
unopkg add -f "-env:UserInstallation=file://$WORK/profile" "$WORK/copy/dist/CalcMap.oxt" > /dev/null 2>&1
cp "$SPREADSHEET" "$WORK/spreadsheet.ods"
rm -f "$WORK"/.~lock*
export CALCMAP_TESTS="$WORK"
export CALCMAP_FIXTURES="$HERE/fixtures"
timeout 300 python3 "$HERE/driver.py" "$WORK/profile" "$WORK/spreadsheet.ods" "$MACRO" "$WORK/map.png" "$PORT" \
	> /dev/null 2>&1 || true
grep -E "TEST|ERROR" "$WORK/profile/user/calcmap.log" | cut -c22-
