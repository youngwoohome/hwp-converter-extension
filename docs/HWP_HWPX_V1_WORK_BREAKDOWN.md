# HWP/HWPX V1 Work Breakdown

Status: Proposed  
Date: 2026-03-07  
Owner: `hwp-converter-extension`

This document breaks the production architecture into a v1 delivery plan with milestones, dependencies, risks, and acceptance criteria.

References:

- [Production HWP/HWPX Architecture](./PRODUCTION_HWP_HWPX_ARCHITECTURE.md)
- [HWP/HWPX Service API Specification](./HWP_HWPX_API_SPEC.md)

## 1. V1 Outcome

At the end of v1, the product must safely support:

- opening `.hwp` and `.hwpx`
- reading structured content from both
- creating new `.hwpx`
- editing `.hwpx`
- creating `.hwpx` working copies from `.hwp`
- exporting to `pdf/html/md/json/docx/txt`
- checkpoint-backed recovery for `.hwpx` saves

It must not support:

- direct `.hwp` overwrite
- unrestricted rich feature editing
- browser parser–based authoritative save

## 2. Delivery Principles

- Protect user data before adding feature breadth.
- Finish the safe write path before broadening edit scope.
- Ship read-only support before write support when needed.
- Treat import fidelity and save integrity as separate tracks.

## 3. Milestones

## Milestone 0: Contract Freeze

Goal:

- freeze scope and public API before implementation starts

Tasks:

- finalize architecture decision
- finalize API schema
- define working-copy policy for `.hwp`
- define allowed licenses for bundled dependencies
- define document telemetry and error taxonomy

Deliverables:

- architecture doc
- API spec
- v1 work breakdown

Acceptance:

- team agrees `.hwpx` is the only editable format in v1
- team agrees `.hwp` edits require working-copy conversion

## Milestone 1: Gateway Hardening

Goal:

- turn the current extension into a production gateway rather than a text conversion prototype

Tasks:

- add request ID generation
- add structured error envelope
- add path normalization and path allowlist checks
- add file size and archive ratio guards
- add sidecar health supervision
- add timeout/retry policy
- add engine version reporting

Dependencies:

- Milestone 0

Acceptance:

- malformed requests never crash the process
- engine-down state is surfaced as a structured error
- logs include request IDs and engine version

## Milestone 2: JVM Core Bootstrap

Goal:

- introduce the authoritative production engine boundary

Tasks:

- package JVM runtime strategy
- create sidecar bootstrap process
- define gateway-to-sidecar RPC or HTTP protocol
- integrate `hwplib`
- integrate `hwpxlib`
- integrate `hwp2hwpx`
- add startup self-test and compatibility matrix logging

Dependencies:

- Milestone 1

Acceptance:

- sidecar starts deterministically on supported environments
- gateway can query sidecar health and engine version
- `.hwp` and `.hwpx` sample files parse via authoritative engine

## Milestone 3: Read Path

Goal:

- production-safe open and read flow for both formats

Tasks:

- implement `POST /document/open`
- implement `GET /document`
- implement `GET /document/features`
- map source document to normalized internal model
- identify unsupported editable features
- add read-only mode for `.hwp`
- optionally wire fast preview helper behind a feature flag

Dependencies:

- Milestone 2

Acceptance:

- `.hwpx` opens as editable
- `.hwp` opens as read-only or import-required
- normalized body contains paragraph/table structure for supported fixtures

## Milestone 4: Working-Copy Import Flow

Goal:

- safely convert `.hwp` into editable `.hwpx`

Tasks:

- implement `POST /document/fork-editable-copy`
- create destination naming rules for working copies
- preserve original path vs working-copy path in session
- emit explicit warnings in responses
- verify imported `.hwpx` can reopen
- add failure handling when import succeeds partially

Dependencies:

- Milestone 3

Acceptance:

- editing a raw `.hwp` always creates a `.hwpx` working copy first
- original `.hwp` never changes
- imported working copy opens as editable

## Milestone 5: Canonical `.hwpx` Create Flow

Goal:

- support new document creation without legacy baggage

Tasks:

- implement `POST /document/create`
- generate blank `.hwpx`
- assign default metadata and structure
- open created document as editable session

Dependencies:

- Milestone 2

Acceptance:

- new `.hwpx` can be created, opened, saved, and reopened

## Milestone 6: Safe `.hwpx` Save Path

Goal:

- deliver the first production-grade mutation and save path

Tasks:

- implement session checkpoint store
- implement atomic temp-write + rename strategy
- implement reopen verification
- implement `POST /document/save`
- support v1 mutation set:
  - replace paragraph text
  - insert paragraph
  - delete paragraph
  - replace table cell text
  - insert table row
  - delete table row
- implement `GET /document/checkpoints`
- implement `POST /document/recover`

Dependencies:

- Milestone 3
- Milestone 5

Acceptance:

- save never mutates file in place
- save produces checkpoint ID
- recover restores a prior checkpoint
- failed save leaves last good file intact

## Milestone 7: Export Pipeline

Goal:

- provide user-facing derived outputs from validated documents

Tasks:

- implement `POST /converter` on top of authoritative model
- support `pdf/html/md/json/docx/txt`
- validate output file existence and readability where possible
- add warnings for semantic-loss exports such as `md`

Dependencies:

- Milestone 3
- Milestone 6 for editable session-based exports

Acceptance:

- all v1 export targets work from `.hwpx`
- `.hwp` exports route through validated import path when necessary

## Milestone 8: UI Integration

Goal:

- expose the safe backend semantics clearly in the app

Tasks:

- add `.hwp` read-only badges
- add “Create editable HWPX copy” action
- show original path vs working-copy path
- surface save warnings and recovery actions
- show unsupported-editable-feature warnings

Dependencies:

- Milestone 4
- Milestone 6

Acceptance:

- user can tell whether they are editing original vs working copy
- user can recover prior checkpoints from UI

## Milestone 9: Production Test Corpus and Release Gates

Goal:

- prevent corruption regressions before rollout

Tasks:

- build fixture corpus
- add parse/save/reopen tests
- add crash recovery tests
- add large file and malformed archive tests
- add golden regression fixtures
- define release blockers

Dependencies:

- Milestones 3 through 7

Acceptance:

- no release goes out without green corruption and recovery suite

## 4. Workstreams

## 4.1 Core Engine Workstream

- sidecar packaging
- parser integration
- internal model mapping
- mutation engine
- serializer

## 4.2 Gateway Workstream

- API schemas
- process supervision
- request validation
- error handling
- telemetry

## 4.3 App Integration Workstream

- open/save/fork flows
- tab state
- working-copy messaging
- recovery UI

## 4.4 QA and Corpus Workstream

- fixtures
- corruption regressions
- performance baselines
- upgrade compatibility checks

## 5. Critical Risks

## Risk 1: `.hwp` Import Fidelity Is Lower Than Expected

Impact:

- imported `.hwpx` may not preserve all layout semantics

Mitigation:

- treat `.hwp` as import-first, not round-trip authoritative
- preserve original `.hwp`
- clearly label working-copy conversion in UI
- build fixture corpus from real Korean business documents

## Risk 2: `.hwpx` Save Drops Unsupported Structures

Impact:

- corruption or destructive edits

Mitigation:

- preserve unsupported nodes untouched
- block unsupported mutations instead of flattening them
- require reopen verification after every save

## Risk 3: Sidecar Operational Complexity

Impact:

- startup failures, runtime crashes, packaging issues

Mitigation:

- explicit engine versioning
- sidecar self-test at startup
- healthcheck endpoint
- deterministic logs and crash diagnostics

## Risk 4: License Drift

Impact:

- legal or distribution risk

Mitigation:

- allow only approved licenses in bundled core
- track dependency licenses in release checklist
- keep excluded tools outside the bundled production path

## Risk 5: Exports Misunderstood as Full-Fidelity

Impact:

- user expectations mismatch, especially for Markdown and JSON

Mitigation:

- label `md/html/json` as semantic exports
- keep `.hwpx` as canonical source of truth

## 6. Release Sequence

Recommended release order:

1. hidden internal read path
2. gated `.hwp` read-only support
3. gated `.hwp -> .hwpx` working-copy flow
4. gated `.hwpx` safe save
5. gated export pipeline
6. broader UI exposure

## 7. Staffing Assumptions

Minimum implementation ownership:

- one engineer on gateway and Electron integration
- one engineer on JVM sidecar and document model
- one engineer or dedicated QA owner for fixture corpus and corruption testing

## 8. Operational Readiness Checklist

Before production launch:

- sidecar version is pinned
- dependency licenses reviewed
- save path has rollback coverage
- working-copy semantics are visible in UI
- malformed file corpus passes safely
- crash recovery path tested
- telemetry dashboards exist for failures and rollbacks

## 9. V1 Definition of Done

V1 is done when all of the following are true:

- raw `.hwp` cannot be overwritten by the app
- `.hwpx` save is atomic and reopen-verified
- `.hwp` can become an editable `.hwpx` working copy
- checkpoints exist and can be restored
- API responses are structured and versioned
- export targets work from the authoritative engine
- fixture-based corruption regression suite is passing

## 10. Post-V1 Candidates

- image editing
- comments and tracked annotations
- header/footer editing
- deeper style preservation
- improved PDF fidelity
- richer preview engine
- collaborative change review
