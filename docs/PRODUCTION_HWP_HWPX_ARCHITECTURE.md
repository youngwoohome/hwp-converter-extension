# Production Architecture for HWP/HWPX Support

Status: Proposed  
Date: 2026-03-07  
Owner: `hwp-converter-extension`

## 1. Summary

We will support Korean office documents by shipping a dedicated HWP/HWPX document service that is safe enough for production use.

The key product decision is:

- `HWPX` is the canonical editable format.
- `.hwp` is treated as a legacy import format.
- The app never performs direct in-place mutation of a `.hwp` binary file.
- The production write path uses a versioned JVM sidecar backed by Apache-2.0 libraries.
- Browser-side and lightweight JS parsers are optional read-only helpers, never the source of truth for save operations.

This replaces the current text-only prototype direction with a structure-preserving document architecture.

## 2. Goals

- Open `.hwp` and `.hwpx` files inside the app.
- Provide production-safe read, convert, create, and edit flows.
- Preserve document structure on `.hwpx` saves.
- Prevent accidental corruption of user files.
- Keep app-facing APIs stable even if the underlying parser/writer changes later.
- Use permissive open-source dependencies in the bundled production core.

## 3. Non-Goals

- Full round-trip editing of raw `.hwp` binaries in v1.
- Shipping GPL/AGPL dependencies in the bundled core save path.
- Using a browser parser as the production writer.
- A full Hancom Office clone in the first release.
- Perfect fidelity for every advanced document feature in v1.

## 4. Constraints

- Korean users expect `.hwp` compatibility, but `.hwp` is materially harder to write safely than `.hwpx`.
- The current Node service shape is useful for integration, but not sufficient as the long-term core.
- The existing prototype rewrites text only. That is not safe for production because it drops structure such as tables, styles, and non-text objects.
- The production system must survive malformed files, oversized archives, process crashes, partial writes, and library regressions.

## 5. Recommended Open-Source Stack

## 5.1 Production Core

- `hwplib`
  - Purpose: `.hwp` read/import support.
  - Role: parse metadata, text, tables, and content needed for import.
  - License: Apache-2.0.
  - Repo: <https://github.com/neolord0/hwplib>

- `hwpxlib`
  - Purpose: `.hwpx` read/write support.
  - Role: canonical structure-preserving read/edit/save engine.
  - License: Apache-2.0.
  - Repo: <https://github.com/neolord0/hwpxlib>

- `hwp2hwpx`
  - Purpose: `.hwp` to `.hwpx` conversion.
  - Role: safe import bridge from legacy binary documents into the editable format.
  - License: Apache-2.0.
  - Repo: <https://github.com/neolord0/hwp2hwpx>

## 5.2 Optional Read-Only Helpers

- `@ohah/hwpjs`
  - Purpose: lightweight HWP preview and neutral export helpers.
  - Role: optional browser/Node read-only preview, never authoritative for save.
  - License: MIT.
  - Repo: <https://github.com/ohah/hwpjs>

## 5.3 Evaluation-Only / Optional Workers

- `python-hwpx`
  - Useful for validation experiments and tooling.
  - Not selected as the primary production write engine because the JVM stack covers HWP and HWPX in one family.
  - Repo: <https://github.com/airmang/python-hwpx>

## 5.4 Explicitly Excluded from Bundled Production Core

- `pyhwp`
  - Excluded from bundled core due to AGPL-3.0 licensing concerns.
  - Repo: <https://github.com/mete0r/pyhwp>

- `H2Orestart`
  - Excluded from bundled core due to GPL-3.0 licensing and heavyweight LibreOffice-style operational cost.
  - Can be evaluated later as a separately installed optional export backend.
  - Repo: <https://github.com/ebandal/H2Orestart>

## 6. Canonical Product Model

We define one internal product truth:

- Original file format:
  - `.hwpx` or `.hwp`
- Canonical editable format:
  - `.hwpx`
- Internal edit model:
  - a structured document AST derived from `.hwpx`

This means:

- `.hwpx` opens directly into the canonical path.
- `.hwp` enters the system through an import pipeline and becomes an editable `.hwpx` working copy before any mutation happens.

## 7. High-Level Architecture

```text
Electron App
  -> HWP Extension Gateway (Node/Express)
    -> Hancom Core Sidecar (JVM)
      -> hwplib / hwpxlib / hwp2hwpx
    -> Optional Preview Worker
      -> hwpjs (read-only)
```

### 7.1 Responsibilities

### App / Electron

- Opens tabs.
- Calls stable HTTP or local RPC endpoints.
- Displays editor and preview surfaces.
- Shows save/import/export warnings.

### HWP Extension Gateway

- Keeps the current app-facing contract stable.
- Validates requests.
- Enforces file path rules, size caps, and mode restrictions.
- Supervises sidecar lifecycle.
- Handles health checks, retries, timeouts, checkpoint coordination, and structured errors.

### Hancom Core Sidecar

- Parses documents.
- Converts `.hwp` to `.hwpx`.
- Builds and validates the internal document model.
- Applies edits and saves `.hwpx`.
- Produces neutral exports such as `pdf/html/md/json/docx`.

### Optional Preview Worker

- Provides fast read-only preview.
- Never writes authoritative output.
- Can be disabled without affecting save correctness.

## 8. Why the JVM Sidecar Is the Production Choice

The bundled production core should use the JVM stack because it gives us:

- One dependency family for `.hwp`, `.hwpx`, and import conversion.
- Permissive Apache-2.0 licensing across the critical write path.
- Better separation between app process and document engine.
- Cleaner packaging boundaries for crash containment and upgrades.
- A safer path to structure-preserving save than text-rewrite prototypes.

We keep Node as the gateway because it already fits the app integration model.

## 9. Document Data Model

The save path must not operate on plain text alone.

The internal edit model should represent at least:

- `document`
- `section`
- `paragraph`
- `run`
- `table`
- `row`
- `cell`
- `image`
- `header`
- `footer`
- `footnote`
- `comment`
- `styleRef`
- `metadata`

v1 may edit only a subset safely, but the persisted model must still preserve unedited nodes.

## 10. File Handling Strategy

## 10.1 `.hwpx`

- Open directly.
- Parse into the internal model.
- Save using atomic write and reopen verification.
- Preserve unsupported structures if they were not edited.

## 10.2 `.hwp`

- Open in read-only import mode first.
- When the user requests edit:
  - create a `.hwpx` working copy through `hwp2hwpx`
  - bind the editing session to that working copy
  - keep the original `.hwp` unchanged
- Never offer silent overwrite back into `.hwp` in v1.

## 10.3 Neutral Exports

- `pdf/html/md/json/docx` are generated from the validated internal model or validated `.hwpx`.
- Exporters must be treated as derived outputs, not as canonical working state.

## 11. Save Semantics

All writes must follow the same contract:

1. Read source document or working copy.
2. Parse and validate.
3. Apply requested edits to structured model.
4. Serialize to a temp output path.
5. Reopen temp output and run validation.
6. If validation passes:
   - create checkpoint/backup metadata
   - atomically replace destination
7. Emit success only after the replacement completes.

The service must never:

- stream partial output directly into the final target path
- mutate the only known good copy in place
- report success before a reopen verification step

## 12. Safety and Reliability Requirements

## 12.1 Atomic Writes

- Write to `*.tmp` in the same directory when possible.
- Use rename-based replacement after verification.
- Keep a rollback checkpoint for the last known good version.

## 12.2 Checkpoints and Recovery

- Every write creates a recoverable checkpoint.
- On startup, the gateway checks for incomplete transactions.
- If a temp file exists without a completed commit marker:
  - discard it or recover it into a crash-report folder
  - never auto-promote it over the original

## 12.3 Reopen Verification

After saving:

- reopen the written file with the same authoritative engine
- verify parse success
- verify basic invariants:
  - document opens
  - section count is valid
  - body content is non-empty when original was non-empty
  - edited nodes are present

## 12.4 Resource Limits

The gateway must enforce:

- max file size
- max unzip ratio for archives
- max XML size per part
- sidecar CPU/memory/time limits
- operation timeouts per endpoint

## 12.5 Untrusted Content Handling

The service must never execute:

- embedded scripts
- macros
- OLE payloads
- external resource fetches referenced from document content

Such content is either:

- preserved as opaque unmodified content when safe
- or stripped on neutral export when preservation is unsafe

## 12.6 Strict Input Validation

- Extension checks are not enough.
- Validate magic bytes and archive structure.
- Reject malformed ZIPs, oversized archives, and invalid XML before entering save path.

## 12.7 Capability Gating

If a document contains unsupported editable features:

- open still succeeds in read mode
- editor shows those features as preserved but non-editable
- save path blocks only the unsupported mutation, not the entire document unless integrity is at risk

## 13. Versioned Engine Packaging

The document engine must be versioned independently from the gateway.

Recommended packaging:

- `hancom-doc-core-<version>.jar`
- bundled JRE runtime
- sidecar manifest with:
  - engine version
  - dependency versions
  - supported feature flags
  - migration notes

Benefits:

- independent rollback
- deterministic bug triage
- safe A/B verification during upgrades

## 14. API Contract

The existing prototype endpoints are acceptable as a starting point, but production needs clearer semantics.

### 14.1 Keep

- `GET /healthcheck`
- `GET /open`
- `GET /document`
- `POST /document/save`
- `POST /converter`

### 14.2 Add

- `POST /document/create`
- `POST /document/fork-editable-copy`
- `GET /document/features`
- `GET /document/checkpoint`
- `POST /document/recover`

### 14.3 Save Response Requirements

Every save response should include:

- `success`
- `outputPath`
- `sourceFormat`
- `workingCopy`
- `engineVersion`
- `validationSummary`
- `checkpointId`
- `warnings[]`

## 15. Editing Policy by Format

## 15.1 v1

- `.hwpx`
  - create
  - read
  - edit
  - save
  - export

- `.hwp`
  - read
  - preview
  - convert to `.hwpx`
  - fork editable working copy
  - export through imported `.hwpx` path

## 15.2 Explicit v1 Restriction

- No direct `.hwp` overwrite.
- Any edit request on `.hwp` results in a `.hwpx` working copy.

## 16. Preview Strategy

Preview and edit correctness are different concerns.

Production rule:

- Preview may use lighter tools.
- Save must use only the authoritative engine.

Recommended split:

- `.hwp` fast preview:
  - optional `hwpjs`
- `.hwpx` editor view:
  - gateway asks sidecar for structured content
  - UI renders structured content

If preview and authoritative parse disagree, the authoritative parse wins and the discrepancy is logged.

## 17. Export Strategy

### 17.1 Production Export Path

- Export from validated internal model or validated `.hwpx`.
- Do not export from raw text snapshots.

### 17.2 Format Priorities

- `json`: diagnostic and debug-safe
- `html`: structural review
- `md`: text-first workflows
- `docx`: interoperability
- `pdf`: user-facing deliverable

### 17.3 Fidelity Policy

- `md/json/html` are semantic exports, not full-fidelity targets.
- `pdf/docx` aim for user-facing output but may still lose unsupported features in early phases.

## 18. Observability

Production support requires structured telemetry from the gateway and sidecar.

We should log:

- open/save/convert latency
- parse failures by engine version
- validation failures by feature type
- unsupported feature counts
- rollback count
- crash recovery count
- `.hwp -> .hwpx` import success rate

Logs must exclude document body text by default.

## 19. Security Boundaries

- The sidecar runs as a child process with a minimal environment.
- It receives only explicit file paths or temp copies.
- No arbitrary network access is required for the core engine.
- No dynamic plugin loading in production.
- Temp files must be created in app-controlled directories.
- Path traversal and symlink handling must be validated before writes.

## 20. Recommended Repository Shape

```text
hwp-converter-extension/
  README.md
  docs/
    PRODUCTION_HWP_HWPX_ARCHITECTURE.md
  gateway/
    node service layer
  core/
    jvm sidecar sources or packaged runtime
  preview/
    optional read-only preview helpers
  tests/
    fixture corpus
```

The current `src/` prototype can evolve into the gateway layer, but the authoritative writer should move into the versioned core boundary.

## 21. Test Strategy

No production release should ship without a real fixture corpus.

Required fixture categories:

- simple text-only `.hwpx`
- `.hwpx` with tables
- `.hwpx` with images
- `.hwpx` with comments/headers/footers
- `.hwp` with mixed content
- malformed `.hwpx`
- oversized zip bomb style archive
- password-protected or unsupported `.hwp`

Required test classes:

- parse tests
- round-trip save tests
- save-then-reopen validation tests
- crash-recovery tests
- large document performance tests
- regression fixtures for every prior corruption bug

## 22. Rollout Plan

## Phase 0: Hardening the Contract

- Freeze app-facing API names.
- Add engine version reporting.
- Add explicit working-copy semantics.

## Phase 1: Read + Convert Production Core

- Introduce JVM sidecar.
- Implement authoritative `.hwp` read/import and `.hwpx` read.
- Keep current editor UI read-oriented.

## Phase 2: Safe `.hwpx` Save Path

- Add structured model edits.
- Add atomic save, checkpointing, and reopen verification.
- Enable production edit for `.hwpx`.

## Phase 3: `.hwp` Edit via Working Copy

- Add one-click editable copy flow.
- Open imported `.hwpx` instead of original `.hwp`.
- Make this behavior explicit in the UI.

## Phase 4: Rich Features

- Tables
- images
- comments
- headers/footers
- deeper style preservation

## 23. Rejected Alternatives

### 23.1 Keep the Current Text-Rewrite Prototype

Rejected because it is not safe for production and loses structure on save.

### 23.2 Pure Browser/JS Save Path

Rejected for production because preview-oriented parsers are not the right trust anchor for authoritative writes.

### 23.3 Direct `.hwp` Editing as a First-Class v1 Goal

Rejected because it increases corruption risk and slows delivery of a safe path.

### 23.4 GPL/AGPL Bundled Core

Rejected due to licensing and packaging risk for the primary product path.

## 24. Final Decision

We will build a production-grade HWP/HWPX subsystem as a layered service:

- Node gateway for app integration
- JVM sidecar as the authoritative parser/writer
- `HWPX` as canonical editable format
- `.hwp` as import-first legacy input
- optional lightweight preview helpers outside the write path

This is the safest path that still gives us a real product, not a prototype.
