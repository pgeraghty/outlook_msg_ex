#!/usr/bin/env python3
import argparse
import base64
import email
from email import policy
import json
import os
import subprocess
from typing import Any


def _norm(v: Any) -> str:
    if v is None:
        return ""
    s = str(v)
    s = s.replace("\x00", "")
    return " ".join(s.split())


def _run(cmd: list[str], cwd: str) -> str:
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"Command failed: {' '.join(cmd)}\nSTDERR:\n{p.stderr}")
    return p.stdout


def parse_elixir(files: list[str], elixir_root: str) -> dict[str, dict[str, Any]]:
    ex = r'''
alias OutlookMsg.Mime

norm = fn v ->
  case v do
    nil -> ""
    x -> x |> to_string() |> String.replace("\u0000", "") |> String.replace(~r/\s+/, " ") |> String.trim()
  end
end
b64 = fn v -> Base.encode64(v || "") end

for f <- System.argv() do
  case OutlookMsg.open_eml(f) do
    {:ok, mime} ->
      subject = Mime.get_header(mime, "Subject") || ""
      from = Mime.get_header(mime, "From") || ""
      to = Mime.get_header(mime, "To") || ""
      cc = Mime.get_header(mime, "Cc") || ""
      ctype = Mime.get_header(mime, "Content-Type") || ""
      body = mime.body || ""
      multipart = Mime.multipart?(mime)
      parts_count = length(mime.parts || [])

      IO.puts(Enum.join([
        "ROW",
        f,
        b64.(norm.(subject)),
        b64.(norm.(from)),
        b64.(norm.(to)),
        b64.(norm.(cc)),
        b64.(norm.(ctype)),
        if(multipart, do: "1", else: "0"),
        Integer.to_string(parts_count),
        Integer.to_string(byte_size(body)),
        b64.(norm.(String.slice(body, 0, 160)))
      ], "|"))

    {:error, reason} ->
      IO.puts("ERR|" <> f <> "|" <> Base.encode64(inspect(reason)))
  end
end
'''

    out = _run(["mix", "run", "-e", ex, "--", *files], cwd=elixir_root)
    rows: dict[str, dict[str, Any]] = {}
    for line in out.splitlines():
      parts = line.strip().split("|")
      if not parts:
        continue
      if parts[0] == "ROW" and len(parts) >= 11:
        _, file, subj, frm, to, cc, ctype, multipart, parts_count, blen, prev = parts[:11]
        dec = lambda s: base64.b64decode(s).decode("utf-8", errors="replace")
        rows[file] = {
          "subject": dec(subj),
          "from": dec(frm),
          "to": dec(to),
          "cc": dec(cc),
          "content_type": dec(ctype),
          "multipart": multipart == "1",
          "parts_count": int(parts_count),
          "body_len": int(blen),
          "body_preview": dec(prev),
        }
      elif parts[0] == "ERR" and len(parts) >= 3:
        rows[parts[1]] = {"error": base64.b64decode(parts[2]).decode("utf-8", errors="replace")}
    return rows


def parse_python(files: list[str]) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for f in files:
      try:
        with open(f, "rb") as fh:
          msg = email.message_from_binary_file(fh, policy=policy.default)

        body = ""
        if msg.is_multipart():
          # first text/plain part
          for part in msg.walk():
            if part.get_content_type() == "text/plain":
              try:
                body = part.get_content()
              except Exception:
                body = str(part.get_payload(decode=True) or b"", errors="replace")
              break
        else:
          try:
            body = msg.get_content()
          except Exception:
            body = str(msg.get_payload(decode=True) or b"", errors="replace")

        rows[f] = {
          "subject": _norm(msg.get("Subject", "")),
          "from": _norm(msg.get("From", "")),
          "to": _norm(msg.get("To", "")),
          "cc": _norm(msg.get("Cc", "")),
          "content_type": _norm(msg.get("Content-Type", "")),
          "multipart": bool(msg.is_multipart()),
          "parts_count": len(msg.get_payload()) if msg.is_multipart() and isinstance(msg.get_payload(), list) else 0,
          "body_len": len(body or ""),
          "body_preview": _norm((body or "")[:160]),
        }
      except Exception as e:
        rows[f] = {"error": f"{type(e).__name__}: {e}"}
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description="EML differential harness: outlook_msg vs Python stdlib email parser")
    ap.add_argument("files", nargs="+", help=".eml files")
    ap.add_argument("--elixir-root", default="/home/sprite/outlook_msg")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    files = [os.path.abspath(f) for f in args.files]
    elx = parse_elixir(files, args.elixir_root)
    py = parse_python(files)

    payload = {"files": {}}
    for f in files:
      payload["files"][f] = {"outlook_msg": elx.get(f, {}), "python_email": py.get(f, {})}

    if args.json:
      print(json.dumps(payload, indent=2, sort_keys=True))
      return 0

    fields = ["subject", "from", "to", "cc", "content_type", "multipart", "parts_count", "body_len"]
    for f in files:
      print(f"\n=== {os.path.basename(f)} ===")
      a = elx.get(f, {})
      b = py.get(f, {})
      if "error" in a or "error" in b:
        print("- error")
        print("  outlook_msg:", a.get("error"))
        print("  python_email:", b.get("error"))
        continue
      for k in fields:
        if a.get(k) != b.get(k):
          print(f"- {k}")
          print("  outlook_msg:", a.get(k))
          print("  python_email:", b.get(k))
    return 0


if __name__ == "__main__":
  raise SystemExit(main())
