# Spec Conformance Policy

## Primary rule
- The parser should follow published specifications whenever they are available.
- Independent parser implementations are used for differential reference, not as the source of truth.

## Authoritative specs by file type
- `.msg`: MS-OXMSG + MS-CFB + MS-OXPROPS + related Exchange protocol docs.
- `.pst`: MS-PST.
- `.eml`: RFC 5322 + MIME RFCs (2045/2046/2047).

## Practical compatibility rule
- Prefer spec-correct behavior by default.
- Accept real-world deviations where possible without violating safety or breaking core semantics.
- Keep tolerant parsing paths covered by tests so malformed inputs do not crash the parser.
- When recovery is possible, retain recoverable content and emit warnings instead of dropping output.

## Test layers
1. Spec conformance tests
- Tagged `:spec_conformance` and run via `tools/spec_gate.sh`.
- Validate explicit RFC/MS behaviors (header unfolding, RFC2047 decoding, critical MAPI tag mappings, strict PST header handling).

2. Differential reference tests
- Compare semantic output with external parsers (`ruby-msg`, `extract-msg`, Python stdlib `email`, optional `msg-viewer`).
- Used to detect drift and implementation blind spots.

3. Snapshot/regression tests
- Golden snapshots for public fixture semantics.
- Guard against accidental behavior changes.

4. Corruption robustness tests
- Mutation tests ensure parsing APIs do not raise on corrupted payloads.
- Goal: return recoverable content when possible; otherwise return structured error, never crash.
- See `test/robustness_test.exs`.
- PST corruption-class guarantees are enforced by `tools/pst_semantic_gate.py`.

## Warning-aware APIs
- `OutlookMsg.open_with_warnings/1`
- `OutlookMsg.open_eml_with_warnings/1`
- `OutlookMsg.open_pst_with_warnings/1`
- `OutlookMsg.open_with_report/1`
- `OutlookMsg.open_eml_with_report/1`
- `OutlookMsg.open_pst_with_report/1`

These APIs return parsed data plus `warnings` so applications can show/hide parser issues while still rendering recovered content.

## Commands
```bash
# Spec-only gate
/home/sprite/outlook_msg/tools/spec_gate.sh /home/sprite/outlook_msg

# Full quality gate (spec + regression + differential checks)
/home/sprite/outlook_msg/tools/run_quality_gates.sh /home/sprite/outlook_msg
```
