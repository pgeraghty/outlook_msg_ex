# File Type Coverage

## `.msg` (Outlook message file)
Status:
- Primary supported format.
- Parsed through OLE/CFB + MAPI property extraction.
- Supports recipients, attachments, body, RTF decompression, and HTML extraction from encapsulated RTF.

Validation:
- Use `tools/msg_diff_harness.py` against independent parsers.
- Prioritize semantic parity over formatting parity.
- Spec behavior is authoritative; comparator outputs are reference-only.

## `.eml` (RFC2822/MIME message)
Status:
- Supported as MIME parse/serialize workflow.
- APIs:
  - `OutlookMsg.open_eml/1` (path or raw EML text)
  - `OutlookMsg.eml_to_string/1`

Notes:
- `.eml` is structurally simpler than `.msg` because it is already MIME text, not MAPI-in-CFB.
- Main risk area is multipart/header edge cases rather than binary container parsing.

Validation:
- Use `tools/eml_diff_harness.py` to compare Elixir MIME interpretation against Python stdlib email parser.
- Use `tools/eml_semantic_gate.py` for CI/pass-fail semantic parity checks.
- Keep regression fixtures in `test/fixtures/public_eml/` and parse them in ExUnit.
- Enforce RFC behaviors with `tools/spec_gate.sh` (`:spec_conformance` tests).

## `.pst` (Outlook personal store)
Status:
- Supported parser with folder/item traversal:
  - `OutlookMsg.open_pst/1`
  - `OutlookMsg.Pst.items/1`, `messages/1`, `folders/1`

Notes:
- PST is significantly more complex than MSG.
- Differential validation should focus on traversal integrity and item/message counts before deep content parity.

Validation:
- Use `tools/pst_probe.py` for structural probes.
- Optional external comparison (if available): `readpst`/`libpff`.
- Enforce strict error behavior and required header parsing via `:spec_conformance` tests.
- Keep regression fixtures in `test/fixtures/public_pst/` for recovery-path coverage.
- Enforce minimum salvage guarantees via `tools/pst_semantic_gate.py`.
