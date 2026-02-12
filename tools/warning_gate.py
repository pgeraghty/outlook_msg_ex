#!/usr/bin/env python3
import argparse
import base64
import json
import os
import subprocess
import sys
from typing import Any


def _run(cmd: list[str], cwd: str) -> str:
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if p.returncode != 0:
      raise RuntimeError(f"Command failed: {' '.join(cmd)}\nSTDERR:\n{p.stderr}")
    return p.stdout


def collect_warnings(files: list[str], elixir_root: str) -> dict[str, list[dict[str, Any]]]:
    ex = r'''
alias OutlookMsg.Warning

enc = fn v -> Base.encode64(to_string(v || "")) end

for f <- System.argv() do
  ext = String.downcase(Path.extname(f))

  result =
    case ext do
      ".msg" -> OutlookMsg.open_with_report(f)
      ".eml" -> OutlookMsg.open_eml_with_report(f)
      ".pst" -> OutlookMsg.open_pst_with_report(f)
      _ -> {:error, :unsupported_extension}
    end

  case result do
    {:ok, _obj, warnings} ->
      if warnings == [] do
        IO.puts("OK|" <> f)
      else
        for w <- warnings do
          case w do
            %Warning{} = ww ->
              IO.puts(Enum.join([
                "WARN",
                f,
                enc.(ww.code),
                enc.(ww.severity),
                enc.(if(ww.recoverable, do: "true", else: "false")),
                enc.(ww.context || ""),
                enc.(ww.message)
              ], "|"))

            txt ->
              IO.puts(Enum.join([
                "WARN",
                f,
                enc.("unstructured_warning"),
                enc.("warn"),
                enc.("true"),
                enc.(""),
                enc.(txt)
              ], "|"))
          end
        end
      end

    {:error, reason} ->
      IO.puts("ERR|" <> f <> "|" <> Base.encode64(inspect(reason)))
  end
end
'''

    out = _run(["mix", "run", "-e", ex, "--", *files], cwd=elixir_root)
    rows: dict[str, list[dict[str, Any]]] = {f: [] for f in files}

    for line in out.splitlines():
      parts = line.strip().split("|")
      if not parts:
        continue
      if parts[0] == "WARN" and len(parts) >= 7:
        _, f, code, sev, rec, ctx, msg = parts[:7]
        dec = lambda s: base64.b64decode(s).decode("utf-8", errors="replace")
        rows.setdefault(f, []).append({
          "code": dec(code),
          "severity": dec(sev),
          "recoverable": dec(rec) == "true",
          "context": dec(ctx),
          "message": dec(msg),
        })
      elif parts[0] == "ERR" and len(parts) >= 3:
        rows.setdefault(parts[1], []).append({
          "code": "parse_error",
          "severity": "error",
          "recoverable": False,
          "context": "",
          "message": base64.b64decode(parts[2]).decode("utf-8", errors="replace"),
        })
      elif parts[0] == "OK" and len(parts) >= 2:
        rows.setdefault(parts[1], [])
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description="Warning policy gate for report-mode parser warnings.")
    ap.add_argument("files", nargs="+", help="input files (.msg/.eml/.pst)")
    ap.add_argument("--policy", default="/home/sprite/outlook_msg/tools/warning_policy.json")
    ap.add_argument("--elixir-root", default="/home/sprite/outlook_msg")
    args = ap.parse_args()

    with open(args.policy, "r", encoding="utf-8") as fh:
      policy = json.load(fh)

    files = [os.path.abspath(f) for f in args.files]
    warnings_by_file = collect_warnings(files, args.elixir_root)

    fail_sev = set(policy.get("fail_on_severity", ["error"]))
    fail_nonrec = bool(policy.get("fail_on_non_recoverable", True))
    max_per_file = int(policy.get("max_warnings_per_file", 50))
    allow_codes = set(policy.get("allow_codes", []))

    failures: list[str] = []

    for f in files:
      rows = warnings_by_file.get(f, [])
      if len(rows) > max_per_file:
        failures.append(f"{f}: too many warnings ({len(rows)} > {max_per_file})")

      for w in rows:
        code = str(w.get("code", ""))
        sev = str(w.get("severity", "warn"))
        rec = bool(w.get("recoverable", True))
        msg = str(w.get("message", ""))

        if allow_codes and code not in allow_codes:
          failures.append(f"{f}: warning code not allowlisted: {code}")
          continue

        if sev in fail_sev:
          failures.append(f"{f}: disallowed severity {sev} code={code} message={msg}")
          continue

        if fail_nonrec and not rec:
          failures.append(f"{f}: non-recoverable warning code={code} message={msg}")

    if failures:
      print("WARNING POLICY GATE FAILED")
      for line in failures:
        print(f"- {line}")
      return 2

    print("Warning policy gate passed.")
    return 0


if __name__ == "__main__":
  raise SystemExit(main())
