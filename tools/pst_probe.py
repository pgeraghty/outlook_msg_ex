#!/usr/bin/env python3
import argparse
import os
import shutil
import subprocess
from typing import Any


def _run(cmd: list[str], cwd: str | None = None) -> str:
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"Command failed: {' '.join(cmd)}\nSTDERR:\n{p.stderr}")
    return p.stdout


def probe_outlook_msg(pst_files: list[str], elixir_root: str) -> dict[str, dict[str, Any]]:
    ex = r'''
for f <- System.argv() do
  case OutlookMsg.open_pst(f) do
    {:ok, pst} ->
      msg_count = pst |> OutlookMsg.Pst.messages() |> Enum.count()
      folder_count = pst |> OutlookMsg.Pst.folders() |> Enum.count()
      item_count = pst |> OutlookMsg.Pst.items() |> Enum.count()
      IO.puts(Enum.join(["ROW", f, Integer.to_string(item_count), Integer.to_string(folder_count), Integer.to_string(msg_count)], "|"))
    {:error, reason} ->
      IO.puts("ERR|" <> f <> "|" <> inspect(reason))
  end
end
'''
    out = _run(["mix", "run", "-e", ex, "--", *pst_files], cwd=elixir_root)
    rows: dict[str, dict[str, Any]] = {}
    for line in out.splitlines():
        parts = line.strip().split("|")
        if not parts:
            continue
        if parts[0] == "ROW" and len(parts) >= 5:
            _, f, items, folders, messages = parts[:5]
            rows[f] = {
                "item_count": int(items),
                "folder_count": int(folders),
                "message_count": int(messages),
            }
        elif parts[0] == "ERR" and len(parts) >= 3:
            rows[parts[1]] = {"error": parts[2]}
    return rows


def probe_readpst(pst_files: list[str]) -> dict[str, dict[str, Any]]:
    if shutil.which("readpst") is None:
        return {}

    rows: dict[str, dict[str, Any]] = {}
    for f in pst_files:
        try:
            # readpst writes files to output dir; use temp dir and count generated eml files.
            import tempfile
            with tempfile.TemporaryDirectory() as td:
                _run(["readpst", "-D", "-M", "-q", "-o", td, f])
                eml = 0
                dirs = 0
                for root, dnames, fnames in os.walk(td):
                    dirs += len(dnames)
                    eml += sum(1 for n in fnames if n.lower().endswith(".eml"))
                rows[f] = {"eml_count": eml, "folder_dirs": dirs}
        except Exception as e:
            rows[f] = {"error": f"{type(e).__name__}: {e}"}
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description="PST structural probe: outlook_msg with optional readpst comparator")
    ap.add_argument("files", nargs="+", help=".pst files")
    ap.add_argument("--elixir-root", default="/home/sprite/outlook_msg")
    args = ap.parse_args()

    files = [os.path.abspath(f) for f in args.files]
    elx = probe_outlook_msg(files, args.elixir_root)
    rp = probe_readpst(files)

    for f in files:
        print(f"\n=== {os.path.basename(f)} ===")
        print("outlook_msg:", elx.get(f, {"error": "no data"}))
        if rp:
            print("readpst:", rp.get(f, {"error": "no data"}))
        else:
            print("readpst: not installed")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
