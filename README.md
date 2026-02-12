# outlook_msg

Elixir library for reading Microsoft Outlook `.msg` and `.pst` files.

## Scope
- OLE/CFB parsing for Outlook storage.
- MSG property extraction (message, recipients, attachments).
- RTF decompression and RTF-derived body/HTML extraction.
- MIME conversion helpers.

Conformance policy:
- Specs are authoritative.
- External parser comparisons are reference checks, not source-of-truth.
- See `docs/spec_conformance.md`.

## Quick Start
```bash
cd /home/sprite/outlook_msg
mix test
```

Parse a message:
```bash
cd /home/sprite/outlook_msg
mix run -e 'alias OutlookMsg.Mapi.PropertySet; {:ok,msg}=OutlookMsg.open("/path/to/file.msg"); IO.puts(PropertySet.subject(msg.properties) || "")'
```

Best-effort parsing with warnings:
```bash
# MSG
{:ok, msg, warnings} = OutlookMsg.open_with_warnings("/path/to/file.msg")

# EML
{:ok, mime, warnings} = OutlookMsg.open_eml_with_warnings("/path/to/file.eml")

# PST
{:ok, pst, warnings} = OutlookMsg.open_pst_with_warnings("/path/to/archive.pst")

# Structured warnings (code/severity/context)
{:ok, msg, warning_structs} = OutlookMsg.open_with_report("/path/to/file.msg")
```

Warning policy:
- Prefer retaining recoverable content over hard failure.
- Attach non-fatal parser issues to `warnings` so callers can decide whether to display or hide them.
- Keep hard failures for unrecoverable container/header conditions.

Public fixture coverage:
- `.msg` fixtures from public sources are stored at `test/fixtures/public_msg/`.
- `.eml` fixtures from public/generic sources are stored at `test/fixtures/public_eml/`.
- Integration tests validate these fixtures directly and run semantic parity checks against `ruby-msg`.
- Golden snapshot for fixture metrics/hashes is stored at `data/snapshots/public_msg_snapshot.exs`.

## Parser Comparison Reference

This repo includes a generalized differential harness for validating `outlook_msg` output against independent parsers.

Tool:
- `tools/msg_diff_harness.py`

Supported comparators:
- `ruby-msg` (Ruby)
- `outlook_msg` (this Elixir lib)
- `extract-msg` (Python)
- `msg-viewer` (optional, TypeScript)

Run:
```bash
/home/sprite/outlook_msg/tools/msg_diff_harness.py \
  --with-msg-viewer \
  /path/to/a.msg /path/to/b.msg

# Baseline-focused report (recommended for triage)
/home/sprite/outlook_msg/tools/msg_diff_harness.py \
  --baseline-parser ruby-msg \
  /path/to/a.msg /path/to/b.msg
```

Additional harnesses:
```bash
# EML parity check (outlook_msg vs Python stdlib email parser)
/home/sprite/outlook_msg/tools/eml_diff_harness.py /path/to/a.eml /path/to/b.eml

# EML semantic gate (CI-style pass/fail)
/home/sprite/outlook_msg/tools/eml_semantic_gate.py /path/to/*.eml

# PST structural probe (and optional readpst comparison if installed)
/home/sprite/outlook_msg/tools/pst_probe.py /path/to/archive.pst

# PST semantic gate (fixture expectations for corruption-recovery behavior)
/home/sprite/outlook_msg/tools/pst_semantic_gate.py /path/to/*.pst

# Semantic gate against ruby-msg baseline (fails on material drift)
/home/sprite/outlook_msg/tools/semantic_gate.py /path/to/*.msg

# Policy gate (CI-oriented, fail on unapproved high-severity deviations)
/home/sprite/outlook_msg/tools/baseline_gate.py /path/to/*.msg

# Refresh golden semantic snapshot for public fixtures
cd /home/sprite/outlook_msg && mix run tools/update_public_msg_snapshot.exs

# Run full local quality gates (tests + snapshot refresh + semantic gate + baseline diff)
/home/sprite/outlook_msg/tools/run_quality_gates.sh /home/sprite/outlook_msg

# Run spec-only conformance gate (RFC + Microsoft Open Specs)
/home/sprite/outlook_msg/tools/spec_gate.sh /home/sprite/outlook_msg
```

### Optional External Comparison Toolchain (Not Project Deps)

These tools are intentionally external to `mix.exs` dependencies and only used for validation/comparison.

Install examples:
```bash
# Python comparator for .msg
python -m pip install extract-msg

# PST external comparator
sudo apt-get update -y && sudo apt-get install -y pst-utils

# Optional TypeScript/browser parser comparator
git clone https://github.com/molotochok/msg-viewer /home/sprite/msg-viewer
```

Used by:
- `tools/msg_diff_harness.py`:
  - `extract-msg` and `msg-viewer` (optional with `--with-msg-viewer`)
- `tools/pst_probe.py`:
  - `readpst` if available on `PATH`
- `tools/eml_semantic_gate.py`:
  - Python stdlib `email` parser (no third-party dependency)
- `tools/baseline_gate.py`:
  - policy file at `tools/baseline_policy.json`

Output:
- Per-file semantic mismatches for key fields:
  - subject/from/to
  - recipient and attachment counts
  - first recipient email
  - body/html lengths

Guidance:
- Treat Microsoft/RFC specs as primary behavior authority.
- Treat `ruby-msg` as a primary baseline for parity checks.
- Ignore formatting-only differences (quoting style, whitespace, timestamp rendering) unless behavior depends on them.
- Treat missing HTML/body extraction as high-severity differences.

## Comparison Checklist

Use this checklist when reviewing parser changes:

1. Correctness (must-pass)
- Subject/from/to resolve to expected semantic values.
- Recipient count and attachment count are stable.
- First recipient SMTP/email resolution is correct (especially EX recipients).
- Body is present when source has body text.
- HTML is present when source has HTML or RTF-encapsulated HTML.

2. Fidelity (high priority)
- Inline attachment metadata (CID/location/content-disposition) is preserved.
- Message-ID and key transport headers are available.
- Body/HTML lengths stay within expected range against baseline parsers.

3. Formatting (lower priority)
- Quoting style, whitespace normalization, and timestamp rendering can differ.
- Prefer semantic equality over byte-for-byte equality unless specifically required.

Recommended baseline order:
1. `ruby-msg`
2. `outlook_msg`
3. `extract-msg`
4. `msg-viewer` (informative, not authoritative for HTML in RTF-heavy samples)

## File Type Coverage

- `.msg`: Primary format; full parser + conversion pipeline.
- `.eml`: Supported via MIME parser/serializer APIs:
  - `OutlookMsg.open_eml/1`
  - `OutlookMsg.eml_to_string/1`
- `.pst`: Parser available (`OutlookMsg.open_pst/1`) with item/folder/message traversal APIs.

See `docs/file_type_coverage.md` for detailed notes and validation approach.
See `docs/warning_taxonomy.md` for warning code semantics and policy gating.

Privacy:
- The harness is path-based and does not write message content into repository files.
- Do not commit private `.msg` fixtures into this repo.

## Deep Reference Material

See:
- `docs/reference_material.md`

Includes:
- MS-OXMSG
- MS-OXPROPS
- MS-OXCMSG
- MS-OXRTFEX
- MS-CFB
- Links to independent parser implementations used for differential verification.
