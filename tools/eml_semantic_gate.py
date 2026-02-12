#!/usr/bin/env python3
import argparse
from email.utils import getaddresses
import os
import sys

try:
    import eml_diff_harness as harness  # type: ignore
except Exception as e:
    print(f"Failed to import eml_diff_harness.py: {e}", file=sys.stderr)
    raise


def _body_tol(expected_len: int) -> int:
    return max(64, int(expected_len * 0.2))


def _norm_addr(value: str) -> str:
    pairs = getaddresses([value or ""])
    norm: list[str] = []
    for name, addr in pairs:
        name = " ".join((name or "").split())
        addr = " ".join((addr or "").split())
        if name and addr:
            norm.append(f"{name} <{addr}>")
        elif addr:
            norm.append(addr)
    return ", ".join(norm)


def main() -> int:
    ap = argparse.ArgumentParser(description="Semantic gate for EML parsing parity (outlook_msg vs Python stdlib).")
    ap.add_argument("files", nargs="+", help=".eml files")
    ap.add_argument("--elixir-root", default="/home/sprite/outlook_msg")
    args = ap.parse_args()

    files = [os.path.abspath(f) for f in args.files]
    elx = harness.parse_elixir(files, args.elixir_root)
    py = harness.parse_python(files)

    failures: list[str] = []
    for f in files:
        a = elx.get(f, {})
        b = py.get(f, {})

        if "error" in a or "error" in b:
            failures.append(f"{f}: parse error elixir={a.get('error')} python={b.get('error')}")
            continue

        if (a.get("subject") or "") != (b.get("subject") or ""):
            failures.append(f"{f}: subject mismatch elixir={a.get('subject')!r} python={b.get('subject')!r}")

        for field in ("from", "to", "cc"):
            if _norm_addr(a.get(field) or "") != _norm_addr(b.get(field) or ""):
                failures.append(f"{f}: {field} mismatch elixir={a.get(field)!r} python={b.get(field)!r}")

        # Content-Type can differ in parameter ordering/casing. Compare main type only.
        ctype_a = (a.get("content_type") or "").split(";", 1)[0].strip().lower()
        ctype_b = (b.get("content_type") or "").split(";", 1)[0].strip().lower()
        if ctype_a != ctype_b:
            failures.append(f"{f}: content_type mismatch elixir={ctype_a!r} python={ctype_b!r}")

        if bool(a.get("multipart")) != bool(b.get("multipart")):
            failures.append(f"{f}: multipart mismatch elixir={a.get('multipart')} python={b.get('multipart')}")

        if int(a.get("parts_count", 0)) != int(b.get("parts_count", 0)):
            failures.append(f"{f}: parts_count mismatch elixir={a.get('parts_count')} python={b.get('parts_count')}")

        body_a = int(a.get("body_len", 0))
        body_b = int(b.get("body_len", 0))
        if abs(body_a - body_b) > _body_tol(body_b):
            failures.append(f"{f}: body_len drift elixir={body_a} python={body_b}")

    if failures:
        print("EML SEMANTIC GATE FAILED")
        for line in failures:
            print(f"- {line}")
        return 2

    print("EML semantic gate passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
