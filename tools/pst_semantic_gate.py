#!/usr/bin/env python3
import argparse
import base64
import os
import subprocess
import sys
from typing import Any


def _run(cmd: list[str], cwd: str) -> str:
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if p.returncode != 0:
      raise RuntimeError(f"Command failed: {' '.join(cmd)}\nSTDERR:\n{p.stderr}")
    return p.stdout


def collect(files: list[str], elixir_root: str) -> dict[str, dict[str, Any]]:
    ex = r'''
for f <- System.argv() do
  case OutlookMsg.open_pst_with_report(f) do
    {:ok, pst, warnings} ->
      codes =
        warnings
        |> Enum.map(fn w ->
          case w do
            %OutlookMsg.Warning{code: code} -> to_string(code)
            _ -> "unstructured_warning"
          end
        end)
        |> Enum.uniq()
        |> Enum.sort()
        |> Enum.join(",")

      IO.puts(Enum.join([
        "ROW",
        f,
        "ok",
        Integer.to_string(map_size(pst.index || %{})),
        Integer.to_string(map_size(pst.descriptors || %{})),
        Base.encode64(codes)
      ], "|"))

    {:error, reason} ->
      IO.puts("ROW|" <> f <> "|error|0|0|" <> Base.encode64(inspect(reason)))
  end
end
'''
    out = _run(["mix", "run", "-e", ex, "--", *files], cwd=elixir_root)
    rows: dict[str, dict[str, Any]] = {}
    for line in out.splitlines():
      parts = line.strip().split("|")
      if len(parts) >= 6 and parts[0] == "ROW":
        _, f, status, idx_count, desc_count, payload = parts[:6]
        rows[f] = {
          "status": status,
          "index_count": int(idx_count),
          "descriptor_count": int(desc_count),
          "codes": [c for c in base64.b64decode(payload).decode("utf-8", errors="replace").split(",") if c],
        }
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description="PST semantic gate for recovery guarantees on known corruption classes.")
    ap.add_argument("files", nargs="+", help=".pst files")
    ap.add_argument("--elixir-root", default="/home/sprite/outlook_msg")
    args = ap.parse_args()

    files = [os.path.abspath(f) for f in args.files]
    rows = collect(files, args.elixir_root)
    failures: list[str] = []

    for f in files:
      base = os.path.basename(f)
      row = rows.get(f)
      if not row:
        failures.append(f"{f}: missing output row")
        continue
      if row["status"] != "ok":
        failures.append(f"{f}: parser returned error {row['codes']}")
        continue

      codes = set(row["codes"])

      if base == "minimal_pst97.pst":
        if codes:
          failures.append(f"{f}: expected clean parse, got warning codes {sorted(codes)}")
      elif base == "corrupt_offsets_pst97.pst":
        required = {"pst_index_parse_failed", "pst_descriptor_parse_failed"}
        if not (codes & required):
          failures.append(f"{f}: expected parse-failed warning code, got {sorted(codes)}")
      elif base == "loop_branch_index_pst97.pst":
        if "pst_branch_loop_detected" not in codes:
          failures.append(f"{f}: expected pst_branch_loop_detected, got {sorted(codes)}")
      else:
        # Generic requirement: no hard error and bounded warning set.
        if len(codes) > 20:
          failures.append(f"{f}: excessive warning cardinality ({len(codes)})")

    if failures:
      print("PST SEMANTIC GATE FAILED")
      for f in failures:
        print(f"- {f}")
      return 2

    print("PST semantic gate passed.")
    return 0


if __name__ == "__main__":
  raise SystemExit(main())
