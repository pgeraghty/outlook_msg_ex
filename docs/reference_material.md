# MSG Parsing Reference Material

Core Microsoft specs (authoritative)
- MS-OXMSG: Outlook Item (.msg) File Format. https://learn.microsoft.com/openspecs/exchange_server_protocols/ms-oxmsg/
- MS-OXPROPS: Exchange Server Protocols Master Property List. https://learn.microsoft.com/openspecs/exchange_server_protocols/ms-oxprops/
- MS-OXCMSG: Message and Attachment Object Protocol. https://learn.microsoft.com/openspecs/exchange_server_protocols/ms-oxcmsg/
- MS-OXRTFEX: Encapsulated Rich Text Format (RTF) Compression Algorithm. https://learn.microsoft.com/openspecs/exchange_server_protocols/ms-oxrtfex/
- MS-CFB: Compound File Binary Format (OLE/CFB container used by .msg). https://learn.microsoft.com/openspecs/windows_protocols/ms-cfb/
- MS-PST: Personal Folder File Format (.pst). https://learn.microsoft.com/openspecs/office_file_formats/ms-pst/

EML/MIME standards
- RFC 5322: Internet Message Format. https://www.rfc-editor.org/rfc/rfc5322
- RFC 2045: MIME Part One (Format of Internet Message Bodies). https://www.rfc-editor.org/rfc/rfc2045
- RFC 2046: MIME Part Two (Media Types). https://www.rfc-editor.org/rfc/rfc2046
- RFC 2047: MIME header encoding. https://www.rfc-editor.org/rfc/rfc2047
- Python `email` package docs (baseline parser used by `tools/eml_*` harnesses): https://docs.python.org/3/library/email.html

Useful independent implementations (cross-checks)
- ruby-msg (Ruby): https://github.com/aquasync/ruby-msg
- extract-msg (Python): https://github.com/TeamMsgExtractor/msg-extractor
- msg-viewer (TypeScript/browser): https://github.com/molotochok/msg-viewer
- libpff/pff-tools (C): https://github.com/libyal/libpff

Validation strategy
- Use at least two independent parsers when changing property/tag mappings.
- Keep differential tests focused on semantic fields (recipient resolution, attachment metadata, body/html availability), not formatting-only differences.
- For HTML fidelity, compare normalized structure and key business strings, not exact byte-for-byte output, because RTF decoders may differ in escaping and line wrapping.

Repo tooling for validation
- `tools/msg_diff_harness.py`: `.msg` differential comparisons across parsers.
- `tools/eml_diff_harness.py`: `.eml` differential checks (`outlook_msg` vs Python stdlib email parser).
- `tools/eml_semantic_gate.py`: `.eml` CI-style semantic parity gate.
- `tools/pst_probe.py`: `.pst` structural probe (with optional `readpst` comparison if installed).
