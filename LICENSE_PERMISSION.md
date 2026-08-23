# Upstream Permission Record — OfficeMCP

## Decision

Distribution of the AiConnect Office connector (fork of
`OfficeMCP/OfficeMCP`) is **PERMITTED** by owner decision (Reza, 2026-08-23),
based on the upstream author's public statement that the project is free to
use ("The most seeable and free way to control Microsof applications by AI
model" — README.md, OfficeMCP v1.0.5).

This supersedes the earlier DISTRIBUTION BLOCKED note (Decision 2026-08-15,
D2).

## Risk acknowledgment (on record)

- Upstream repository carries NO LICENSE file and NO license field in
  pyproject.toml (verified 2026-08-23 against github.com/OfficeMCP/OfficeMCP).
- Under default copyright law, informal "free to use" wording is weaker than
  a formal license grant.
- Mitigations in place:
  1. Fork credits upstream in README Attribution.
  2. The bulk of connector value (18 document-CRUD tools in
     `officemcp/documents.py`, adapter layer, Layer B) is original work built
     on MIT-licensed Python libraries (python-docx, openpyxl, python-pptx) —
     independent of upstream code.
  3. Only the 13 thin COM lifecycle wrappers derive from upstream.

If a takedown or license dispute ever arises, fallback plan: clean-room
rewrite of the 13 COM wrappers (trivial) removes all upstream code.
