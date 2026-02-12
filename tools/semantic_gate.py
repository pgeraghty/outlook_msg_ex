#!/usr/bin/env python3
import argparse
import os
import sys
from typing import Any


def _load_harness():
    here = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, here)
    import msg_diff_harness as harness  # type: ignore

    return harness


def _norm(v: Any) -> str:
    if v is None:
        return ""
    return " ".join(str(v).replace("\x00", "").split())


def _body_tol(baseline_len: int) -> int:
    return max(128, int(baseline_len * 0.02))


def _html_tol(baseline_len: int) -> int:
    return max(4096, int(baseline_len * 0.05))


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Semantic gate: fail on material drift between outlook_msg and ruby-msg."
    )
    ap.add_argument("files", nargs="+", help=".msg files")
    ap.add_argument("--ruby-root", default="/home/sprite/ruby-msg")
    ap.add_argument("--elixir-root", default="/home/sprite/outlook_msg")
    args = ap.parse_args()

    files = [os.path.abspath(f) for f in args.files]

    harness = _load_harness()
    ruby = harness.parse_ruby(files, args.ruby_root)
    elixir = harness.parse_elixir(files, args.elixir_root)

    failures: list[str] = []

    for f in files:
        r = ruby.get(f, {})
        e = elixir.get(f, {})
        name = os.path.basename(f)

        if "error" in r or "error" in e:
            failures.append(f"[{name}] parse error ruby={r.get('error')} elixir={e.get('error')}")
            continue

        # Must-match semantic invariants
        for field in ["subject", "recipient_count", "attachment_count", "first_recipient_email"]:
            rv = _norm(r.get(field))
            ev = _norm(e.get(field))
            if rv != ev:
                failures.append(f"[{name}] {field} mismatch ruby={rv!r} elixir={ev!r}")

        # Length drift thresholds
        rb = int(r.get("body_len", 0))
        eb = int(e.get("body_len", 0))
        if abs(rb - eb) > _body_tol(rb):
            failures.append(f"[{name}] body_len drift ruby={rb} elixir={eb} tol={_body_tol(rb)}")

        rh = int(r.get("html_len", 0))
        eh = int(e.get("html_len", 0))
        if abs(rh - eh) > _html_tol(rh):
            failures.append(f"[{name}] html_len drift ruby={rh} elixir={eh} tol={_html_tol(rh)}")

        # Only fail hard on HTML loss when ruby has meaningful HTML and no
        # plain-text fallback. If ruby already has body text, html-only loss
        # can be formatting/path-dependent between implementations.
        if rh > 0 and eh == 0 and rb == 0:
            failures.append(f"[{name}] html missing in elixir while ruby has html_len={rh} and no text body")

    if failures:
        print("SEMANTIC GATE FAILED")
        for line in failures:
            print("-", line)
        return 1

    print("Semantic gate passed for all files.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
