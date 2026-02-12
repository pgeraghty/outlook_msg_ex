#!/usr/bin/env bash
set -euo pipefail

ROOT="${1:-/home/sprite/outlook_msg}"

echo "Running spec-conformance tests (RFC + MS Open Specs)..."
(
  cd "$ROOT"
  mix test --only spec_conformance
)

echo "Spec gate passed."
