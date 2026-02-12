#!/usr/bin/env python3
import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
from typing import Any


def _sha(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="replace")).hexdigest()


def _norm(v: Any) -> str:
    if v is None:
        return ""
    s = str(v)
    s = s.replace("\x00", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _preview(v: Any, n: int = 160) -> str:
    return _norm(v)[:n]


def _run(cmd: list[str], cwd: str | None = None) -> str:
    p = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(f"Command failed: {' '.join(cmd)}\nSTDOUT:\n{p.stdout}\nSTDERR:\n{p.stderr}")
    return p.stdout


def parse_ruby(files: list[str], ruby_root: str) -> dict[str, dict[str, Any]]:
    ruby_code = r'''
require "json"
require "digest"
require "mapi/msg"

def norm(v)
  return "" if v.nil?
  v.to_s.gsub(/\x00/, "").gsub(/\s+/, " ").strip
end

ARGV.each do |f|
  begin
    msg = Mapi::Msg.open(f)
    p = msg.props
    body = (p.body || "").to_s
    html = (p.body_html || "").to_s
    rec = msg.recipients.first

    out = {
      parser: "ruby-msg",
      file: f,
      subject: norm(p.subject),
      from: norm(msg.from),
      to: norm(msg.to),
      cc: norm(msg.cc),
      bcc: norm(msg.bcc),
      message_id: norm(p.internet_message_id),
      recipient_count: msg.recipients.length,
      attachment_count: msg.attachments.length,
      first_recipient_email: norm(rec&.email),
      body_len: body.bytesize,
      html_len: html.bytesize,
      body_sha256: Digest::SHA256.hexdigest(body),
      html_sha256: Digest::SHA256.hexdigest(html),
      body_preview: norm(body)[0,160],
      html_preview: norm(html)[0,160]
    }
    puts JSON.generate(out)
    msg.close
  rescue => e
    puts JSON.generate({parser: "ruby-msg", file: f, error: "#{e.class}: #{e.message}"})
  end
end
'''
    out = _run(["ruby", "-Ilib", "-e", ruby_code, *files], cwd=ruby_root)
    results: dict[str, dict[str, Any]] = {}
    for line in out.splitlines():
        if not line.strip().startswith("{"):
            continue
        row = json.loads(line)
        results[row["file"]] = row
    return results


def parse_elixir(files: list[str], elixir_root: str) -> dict[str, dict[str, Any]]:
    ex_code = r'''
alias OutlookMsg.Mapi.PropertySet

norm = fn v ->
  case v do
    nil -> ""
    x -> x |> to_string() |> String.replace("\u0000", "") |> String.replace(~r/\s+/, " ") |> String.trim()
  end
end
b64 = fn v -> Base.encode64(v || "") end

for f <- System.argv() do
  case OutlookMsg.open(f) do
    {:ok, msg} ->
      p = msg.properties
      body = PropertySet.body(p) || ""
      html = PropertySet.body_html(p) || ""
      body = if is_binary(body), do: body, else: inspect(body)
      html = if is_binary(html), do: html, else: inspect(html)
      rec = List.first(msg.recipients)
      subject = norm.(PropertySet.subject(p))
      from = norm.(PropertySet.sender_name(p)) <> " <" <> norm.(PropertySet.sender_email(p)) <> ">"
      to = norm.(PropertySet.display_to(p))
      cc = norm.(PropertySet.display_cc(p))
      bcc = norm.(PropertySet.display_bcc(p))
      message_id = norm.(PropertySet.get(p, :pr_internet_message_id))
      first_rec_email = if(rec, do: norm.(rec.email), else: "")
      body_preview = body |> norm.() |> String.slice(0, 160)
      html_preview = html |> norm.() |> String.slice(0, 160)

      IO.puts(Enum.join([
        "ROW",
        f,
        b64.(subject),
        b64.(from),
        b64.(to),
        b64.(cc),
        b64.(bcc),
        b64.(message_id),
        Integer.to_string(length(msg.recipients)),
        Integer.to_string(length(msg.attachments)),
        b64.(first_rec_email),
        Integer.to_string(byte_size(body)),
        Integer.to_string(byte_size(html)),
        :crypto.hash(:sha256, body) |> Base.encode16(case: :lower),
        :crypto.hash(:sha256, html) |> Base.encode16(case: :lower),
        b64.(body_preview),
        b64.(html_preview)
      ], "|"))

    {:error, reason} ->
      IO.puts("ERR|" <> f <> "|" <> Base.encode64(inspect(reason)))
  end
end
'''
    out = _run(["mix", "run", "-e", ex_code, "--", *files], cwd=elixir_root)
    results: dict[str, dict[str, Any]] = {}
    for line in out.splitlines():
        parts = line.strip().split("|")
        if not parts:
            continue
        tag = parts[0]
        if tag == "ROW" and len(parts) >= 17:
            (
                _,
                file,
                subject_b64,
                from_b64,
                to_b64,
                cc_b64,
                bcc_b64,
                msgid_b64,
                recipient_count,
                attachment_count,
                first_rec_b64,
                body_len,
                html_len,
                body_sha256,
                html_sha256,
                body_prev_b64,
                html_prev_b64,
            ) = parts[:17]
            dec = lambda s: __import__("base64").b64decode(s).decode("utf-8", errors="replace")
            results[file] = {
                "parser": "outlook_msg",
                "file": file,
                "subject": dec(subject_b64),
                "from": dec(from_b64),
                "to": dec(to_b64),
                "cc": dec(cc_b64),
                "bcc": dec(bcc_b64),
                "message_id": dec(msgid_b64),
                "recipient_count": int(recipient_count),
                "attachment_count": int(attachment_count),
                "first_recipient_email": dec(first_rec_b64),
                "body_len": int(body_len),
                "html_len": int(html_len),
                "body_sha256": body_sha256,
                "html_sha256": html_sha256,
                "body_preview": dec(body_prev_b64),
                "html_preview": dec(html_prev_b64),
            }
        elif tag == "ERR" and len(parts) >= 3:
            import base64
            file = parts[1]
            err = base64.b64decode(parts[2]).decode("utf-8", errors="replace")
            results[file] = {"parser": "outlook_msg", "file": file, "error": err}
    return results


def parse_extract_msg(files: list[str]) -> dict[str, dict[str, Any]]:
    import extract_msg  # type: ignore

    out: dict[str, dict[str, Any]] = {}
    for f in files:
        try:
            msg = extract_msg.Message(f)
            body = msg.body or ""
            html = msg.htmlBody or ""
            if isinstance(html, bytes):
                html = html.decode("utf-8", errors="replace")
            recips = getattr(msg, "recipients", []) or []
            atts = getattr(msg, "attachments", []) or []
            first_rec_email = ""
            if recips:
                first = recips[0]
                first_rec_email = _norm(getattr(first, "email", ""))
            row = {
                "parser": "extract-msg",
                "file": f,
                "subject": _norm(getattr(msg, "subject", "")),
                "from": _norm(getattr(msg, "sender", "")),
                "to": _norm(getattr(msg, "to", "")),
                "cc": _norm(getattr(msg, "cc", "")),
                "bcc": _norm(getattr(msg, "bcc", "")),
                "message_id": _norm(getattr(msg, "messageId", "")),
                "recipient_count": len(recips),
                "attachment_count": len(atts),
                "first_recipient_email": first_rec_email,
                "body_len": len(body),
                "html_len": len(html),
                "body_sha256": _sha(body),
                "html_sha256": _sha(html),
                "body_preview": _preview(body),
                "html_preview": _preview(html),
            }
            out[f] = row
            try:
                msg.close()
            except Exception:
                pass
        except Exception as e:
            out[f] = {"parser": "extract-msg", "file": f, "error": f"{type(e).__name__}: {e}"}
    return out


def parse_msg_viewer(files: list[str], viewer_root: str) -> dict[str, dict[str, Any]]:
    if not os.path.isdir(viewer_root):
        return {}

    ts = r'''
import { parse } from "./lib/scripts/msg/msg-parser";
import { readFileSync } from "fs";

const files = process.argv.slice(2);

const norm = (v: any) => {
  if (v == null) return "";
  return String(v).replace(/\u0000/g, "").replace(/\s+/g, " ").trim();
};

for (const f of files) {
  try {
    const buf = readFileSync(f);
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const msg = parse(new DataView(ab));
    const body = msg.content.body ? String(msg.content.body) : "";
    const html = msg.content.bodyHTML ? String(msg.content.bodyHTML) : "";

    const rec = msg.recipients && msg.recipients.length > 0 ? msg.recipients[0] : null;

    const row = {
      parser: "msg-viewer",
      file: f,
      subject: norm(msg.content.subject),
      from: `${norm(msg.content.senderName)} <${norm(msg.content.senderEmail)}>`,
      to: norm(msg.content.toRecipients),
      cc: norm(msg.content.ccRecipients),
      bcc: "",
      message_id: "",
      recipient_count: msg.recipients.length,
      attachment_count: msg.attachments.length,
      first_recipient_email: rec ? norm(rec.email) : "",
      body_len: body.length,
      html_len: html.length,
      body_sha256: "",
      html_sha256: "",
      body_preview: norm(body).slice(0, 160),
      html_preview: norm(html).slice(0, 160),
    };

    console.log(JSON.stringify(row));
  } catch (e: any) {
    console.log(JSON.stringify({ parser: "msg-viewer", file: f, error: String(e) }));
  }
}
'''

    with tempfile.NamedTemporaryFile("w", suffix=".ts", delete=False, dir=viewer_root) as tmp:
        tmp.write(ts)
        script = tmp.name

    try:
        out = _run(["bun", script, *files], cwd=viewer_root)
    finally:
        try:
            os.remove(script)
        except OSError:
            pass

    rows = {}
    for line in out.splitlines():
        if line.strip().startswith("{"):
            r = json.loads(line)
            rows[r["file"]] = r
    return rows


def compare(files: list[str], parsers: dict[str, dict[str, dict[str, Any]]]) -> dict[str, Any]:
    fields = [
        "subject",
        "from",
        "to",
        "recipient_count",
        "attachment_count",
        "first_recipient_email",
        "body_len",
        "html_len",
    ]

    summary: dict[str, Any] = {"files": {}}
    for f in files:
        by_parser = {name: data.get(f, {}) for name, data in parsers.items()}
        mismatches = {}
        for field in fields:
            vals = {name: by_parser[name].get(field) for name in by_parser if by_parser[name]}
            uniq = set(json.dumps(v, sort_keys=True) for v in vals.values())
            if len(uniq) > 1:
                mismatches[field] = vals

        summary["files"][f] = {
            "parsers": by_parser,
            "mismatches": mismatches,
        }
    return summary


def _severity(field: str) -> str:
    if field in {"subject", "recipient_count", "attachment_count", "first_recipient_email"}:
        return "high"
    if field in {"from", "to"}:
        return "medium"
    if field in {"body_len", "html_len"}:
        return "medium"
    return "low"


def baseline_diff(files: list[str], parsers: dict[str, dict[str, dict[str, Any]]], baseline: str) -> dict[str, Any]:
    fields = [
        "subject",
        "from",
        "to",
        "recipient_count",
        "attachment_count",
        "first_recipient_email",
        "body_len",
        "html_len",
    ]

    out: dict[str, Any] = {"files": {}}
    base_data = parsers.get(baseline, {})

    for f in files:
        base = base_data.get(f, {})
        file_rows = {}
        for parser_name, pdata in parsers.items():
            if parser_name == baseline:
                continue
            row = pdata.get(f, {})
            issues = []
            if not base:
                issues.append({"field": "parser", "severity": "high", "baseline": None, "actual": row, "note": "missing baseline row"})
            elif not row:
                issues.append({"field": "parser", "severity": "high", "baseline": base, "actual": None, "note": "missing comparator row"})
            else:
                for field in fields:
                    bv = base.get(field)
                    rv = row.get(field)
                    if bv == rv:
                        continue
                    # Length fields use tolerance band to reduce false alarms.
                    if field == "body_len":
                        tol = max(128, int((bv or 0) * 0.02))
                        if isinstance(bv, int) and isinstance(rv, int) and abs(bv - rv) <= tol:
                            continue
                    if field == "html_len":
                        tol = max(4096, int((bv or 0) * 0.05))
                        if isinstance(bv, int) and isinstance(rv, int) and abs(bv - rv) <= tol:
                            continue
                    issues.append({
                        "field": field,
                        "severity": _severity(field),
                        "baseline": bv,
                        "actual": rv,
                    })
            file_rows[parser_name] = issues
        out["files"][f] = file_rows
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Differential .msg parser harness (no content persisted).")
    ap.add_argument("files", nargs="+", help=".msg files to compare")
    ap.add_argument("--ruby-root", default="/home/sprite/ruby-msg")
    ap.add_argument("--elixir-root", default="/home/sprite/outlook_msg")
    ap.add_argument("--msg-viewer-root", default="/home/sprite/msg-viewer")
    ap.add_argument("--no-ruby", action="store_true")
    ap.add_argument("--no-elixir", action="store_true")
    ap.add_argument("--no-extract-msg", action="store_true")
    ap.add_argument("--with-msg-viewer", action="store_true")
    ap.add_argument(
        "--baseline-parser",
        choices=["ruby-msg", "outlook_msg", "extract-msg", "msg-viewer"],
        help="Show only deviations from a single baseline parser.",
    )
    ap.add_argument("--json", action="store_true", help="Emit full JSON payload")
    args = ap.parse_args()

    files = [os.path.abspath(f) for f in args.files]

    parsers: dict[str, dict[str, dict[str, Any]]] = {}

    if not args.no_ruby:
        parsers["ruby-msg"] = parse_ruby(files, args.ruby_root)
    if not args.no_elixir:
        parsers["outlook_msg"] = parse_elixir(files, args.elixir_root)
    if not args.no_extract_msg:
        parsers["extract-msg"] = parse_extract_msg(files)
    if args.with_msg_viewer:
        parsers["msg-viewer"] = parse_msg_viewer(files, args.msg_viewer_root)

    if args.baseline_parser:
        payload = baseline_diff(files, parsers, args.baseline_parser)
    else:
        payload = compare(files, parsers)

    if args.json:
        print(json.dumps(payload, indent=2, sort_keys=True))
        return 0

    if args.baseline_parser:
        for f, parser_issues in payload["files"].items():
            print(f"\n=== {os.path.basename(f)} (baseline: {args.baseline_parser}) ===")
            total = sum(len(v) for v in parser_issues.values())
            if total == 0:
                print("No material deviations from baseline.")
                continue
            for parser_name, issues in parser_issues.items():
                if not issues:
                    continue
                print(f"- {parser_name}")
                for issue in issues:
                    print(
                        f"  [{issue['severity']}] {issue['field']} "
                        f"baseline={issue.get('baseline')} actual={issue.get('actual')}"
                    )
    else:
        for f, data in payload["files"].items():
            print(f"\n=== {os.path.basename(f)} ===")
            mm = data["mismatches"]
            if not mm:
                print("No semantic mismatches across enabled parsers.")
                continue
            for field, vals in mm.items():
                print(f"- {field}")
                for parser, val in vals.items():
                    print(f"  {parser}: {val}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
