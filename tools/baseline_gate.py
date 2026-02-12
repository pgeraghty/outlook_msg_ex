#!/usr/bin/env python3
import argparse
import json
import os
import sys
from typing import Any


def _load_harness():
    here = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, here)
    import msg_diff_harness as harness  # type: ignore

    return harness


def _allowed(issue: dict[str, Any], parser_name: str, allow_rules: list[dict[str, Any]]) -> bool:
    for rule in allow_rules:
        if rule.get("parser") and rule["parser"] != parser_name:
            continue
        if rule.get("field") and rule["field"] != issue.get("field"):
            continue
        if rule.get("severity") and rule["severity"] != issue.get("severity"):
            continue
        return True
    return False


def main() -> int:
    ap = argparse.ArgumentParser(description="Policy-based baseline deviation gate.")
    ap.add_argument("files", nargs="+", help=".msg files")
    ap.add_argument("--policy", default="/home/sprite/outlook_msg/tools/baseline_policy.json")
    ap.add_argument("--ruby-root", default="/home/sprite/ruby-msg")
    ap.add_argument("--elixir-root", default="/home/sprite/outlook_msg")
    ap.add_argument("--msg-viewer-root", default="/home/sprite/msg-viewer")
    ap.add_argument("--with-msg-viewer", action="store_true")
    args = ap.parse_args()

    with open(args.policy, "r", encoding="utf-8") as fh:
        policy = json.load(fh)

    baseline = policy.get("baseline_parser", "ruby-msg")
    fail_on = set(policy.get("fail_on_severity", ["high"]))
    allow_rules = policy.get("allow_rules", [])

    files = [os.path.abspath(f) for f in args.files]
    harness = _load_harness()

    parsers: dict[str, dict[str, dict[str, Any]]] = {
        "ruby-msg": harness.parse_ruby(files, args.ruby_root),
        "outlook_msg": harness.parse_elixir(files, args.elixir_root),
        "extract-msg": harness.parse_extract_msg(files),
    }
    if args.with_msg_viewer:
        parsers["msg-viewer"] = harness.parse_msg_viewer(files, args.msg_viewer_root)

    diff = harness.baseline_diff(files, parsers, baseline)

    violations: list[str] = []
    for file, parser_issues in diff["files"].items():
        name = os.path.basename(file)
        for parser_name, issues in parser_issues.items():
            for issue in issues:
                sev = issue.get("severity")
                if sev not in fail_on:
                    continue
                if _allowed(issue, parser_name, allow_rules):
                    continue
                violations.append(
                    f"[{name}] [{parser_name}] [{sev}] {issue.get('field')} "
                    f"baseline={issue.get('baseline')} actual={issue.get('actual')}"
                )

    if violations:
        print("BASELINE POLICY GATE FAILED")
        for v in violations:
            print("-", v)
        return 1

    print("Baseline policy gate passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

