# Warning Taxonomy

Structured warnings are represented by `%OutlookMsg.Warning{}` with:
- `code` (stable machine-readable atom)
- `severity` (`:info | :warn | :error`)
- `message` (human-readable)
- `context` (optional origin detail)
- `recoverable` (whether parsing continued with useful output)

## Current warning codes
- `:malformed_header_line`
- `:multipart_missing_boundary`
- `:nested_part_warning`
- `:nameid_parse_failed`
- `:attachment_skipped`
- `:property_parse_failed`
- `:pst_index_parse_failed`
- `:pst_descriptor_parse_failed`
- `:pst_branch_loop_detected`

## API surface
- String-compatible warnings:
  - `OutlookMsg.open_with_warnings/1`
  - `OutlookMsg.open_eml_with_warnings/1`
  - `OutlookMsg.open_pst_with_warnings/1`
- Structured warnings:
  - `OutlookMsg.open_with_report/1`
  - `OutlookMsg.open_eml_with_report/1`
  - `OutlookMsg.open_pst_with_report/1`

## Policy gate
- `tools/warning_gate.py` enforces warning policy from `tools/warning_policy.json`.
- Default policy:
  - fails on `error` severity
  - fails on non-recoverable warnings
  - bounds warning volume per file
  - checks warning codes are allowlisted
