#!/usr/bin/env bash
set -euo pipefail

ROOT="${1:-/home/sprite/outlook_msg}"
FIXTURES_GLOB="${2:-$ROOT/test/fixtures/public_msg/*.msg}"
EML_FIXTURES_GLOB="${3:-$ROOT/test/fixtures/public_eml/*.eml}"
PST_FIXTURES_GLOB="${4:-$ROOT/test/fixtures/public_pst/*.pst}"

echo "[1/9] Running spec-conformance gate..."
"$ROOT/tools/spec_gate.sh" "$ROOT"

echo "[2/9] Running Elixir test suite..."
(
  cd "$ROOT"
  mix test
)

echo "[3/9] Refreshing public fixture semantic snapshot..."
(
  cd "$ROOT"
  mix run tools/update_public_msg_snapshot.exs
)

echo "[4/9] Running semantic gate against ruby-msg baseline..."
"$ROOT/tools/semantic_gate.py" $FIXTURES_GLOB

echo "[5/9] Running baseline diff report (informational)..."
"$ROOT/tools/msg_diff_harness.py" --baseline-parser ruby-msg $FIXTURES_GLOB

echo "[6/9] Running baseline policy gate (CI-fail on unapproved high-severity drift)..."
"$ROOT/tools/baseline_gate.py" $FIXTURES_GLOB

if compgen -G "$EML_FIXTURES_GLOB" > /dev/null; then
  echo "[7/9] Running EML semantic gate against Python stdlib parser..."
  "$ROOT/tools/eml_semantic_gate.py" $EML_FIXTURES_GLOB
else
  echo "[7/9] Skipping EML semantic gate (no EML fixtures found)."
fi

if compgen -G "$PST_FIXTURES_GLOB" > /dev/null; then
  echo "[8/9] Running PST semantic gate..."
  "$ROOT/tools/pst_semantic_gate.py" $PST_FIXTURES_GLOB
else
  echo "[8/9] Skipping PST semantic gate (no PST fixtures found)."
fi

echo "[9/9] Running warning policy gate..."
warn_inputs=()
if compgen -G "$FIXTURES_GLOB" > /dev/null; then
  for f in $FIXTURES_GLOB; do warn_inputs+=("$f"); done
fi
if compgen -G "$EML_FIXTURES_GLOB" > /dev/null; then
  for f in $EML_FIXTURES_GLOB; do warn_inputs+=("$f"); done
fi
if compgen -G "$PST_FIXTURES_GLOB" > /dev/null; then
  for f in $PST_FIXTURES_GLOB; do warn_inputs+=("$f"); done
fi

if [ ${#warn_inputs[@]} -gt 0 ]; then
  "$ROOT/tools/warning_gate.py" "${warn_inputs[@]}"
else
  echo "No files found for warning policy gate; skipping."
fi

echo "Quality gates passed."
