# ooeditor-hwp Product Architecture

## Goal

`ooeditor-hwp` is a dedicated HWP/HWPX extension product, distributed separately from `oo-editors`.

It is **not** a fork of `oo-editors` and **not** a replacement for Hancom Office.
It is a production extension focused on:

- opening `hwp` and `hwpx` files inside the app
- deterministic document automation
- safe `.hwpx` working-copy editing
- high-fidelity export for review and downstream editing
- handoff back to native Hancom workflows on Windows

## Why Split It Out

`oo-editors` is optimized around ONLYOFFICE and OOXML-family editing.
HWP/HWPX needs a different product shape:

- dedicated import pipeline for `.hwp`
- canonical editable format as `.hwpx`
- deterministic template filling instead of freeform editor-first behavior
- HWP-specific table, header/footer, and image mutation rules
- separate runtime dependencies such as a JVM core

Trying to force this into `oo-editors` creates the wrong incentives:

- HWP support becomes tied to OOXML/editor assumptions
- release cadence becomes coupled to unrelated doc/xls/ppt changes
- bundle size and runtime requirements get mixed together
- host integration logic becomes harder to reason about

## Product Position

`ooeditor-hwp` should be positioned as:

- a Korean document automation and review extension
- a safe HWP/HWPX workflow bridge for macOS and agent-driven automation
- a system that gets the document to 80-90% completion and hands off to native Hancom editing when needed

It should **not** be positioned as:

- a drop-in replacement for the Hancom editor
- a full WYSIWYG Hancom clone
- a dependency hidden inside `oo-editors`

## Primary User Flows

### 1. Open and Inspect

- user opens `.hwp` or `.hwpx`
- extension parses the source
- if `.hwp`, the source is read-only until promoted into `.hwpx`
- extension shows structured preview, analysis, and available mutation targets

### 2. Working Copy Automation

- user or agent creates an editable `.hwpx` working copy
- agent fills placeholders, repeats rows, inserts images, and updates headers/footers
- extension saves through atomic write + reopen verification

### 3. Review and Handoff

- user reviews the output inside the app
- output is exported as `.hwpx`, `pdf`, `html`, `md`, or `json`
- final polish is done in the native Hancom editor if needed

## Product Boundary

### In Scope

- `hwp` import
- `hwpx` canonical edit/save
- deterministic template fill
- merged-table safe repeat
- image/header/footer structured mutations
- markdown fidelity export
- checkpoints and recovery
- host-app integration over localhost HTTP or IPC proxy

### Out of Scope

- direct authoritative `.hwp` round-trip editing
- full Hancom-like ribbon UI
- arbitrary drag-and-drop layout editing
- perfect WYSIWYG reproduction of Hancom Office

## Runtime Architecture

### Host App

The host app is responsible for:

- installing the extension bundle
- launching the extension process
- healthchecking the extension
- routing `.hwp/.hwpx` files into the extension viewer
- exposing automation tools to the agent

### ooeditor-hwp Service

The extension service is responsible for:

- document session management
- working-copy lifecycle
- safe save/recover/checkpoints
- structured HWPX mutation
- deterministic automation APIs
- export and download endpoints

### Authoritative Core

The authoritative core is responsible for:

- `.hwp -> .hwpx` import
- low-level HWPX package reads/writes
- structure-preserving operations
- validation and reopen verification

This layer can stay embedded inside the extension repo, but it should remain a clearly separated subsystem.

## Canonical Format Policy

- authoritative editable format: `.hwpx`
- legacy input format: `.hwp`
- review and handoff output: `.hwpx`

Rules:

- raw `.hwp` is never overwritten
- editing starts only after a managed `.hwpx` working copy is created
- all automation APIs target the working `.hwpx`

## Packaging Model

`ooeditor-hwp` should be released as a standalone GitHub release bundle, similar to `oo-editors`.

Recommended bundle contents:

- `dist/server.js`
- `package.json`
- production `node_modules`
- `jvm-core/target/*.jar`
- static assets for preview/open pages
- version metadata

Recommended install layout in the host app:

- `userData/ooeditor-hwp/`

Recommended release artifact names:

- `ooeditor-hwp-darwin-arm64.zip`
- `ooeditor-hwp-darwin-x64.zip`
- `ooeditor-hwp-win32-x64.zip`
- `ooeditor-hwp-linux-x64.zip`

## Repo Structure

Recommended standalone repo structure:

```text
ooeditor-hwp/
  src/
    core/
    converters/
    routes/
    utils/
  jvm-core/
  docs/
  fixtures/
  scripts/
  package.json
  README.md
```

Recommended logical layers:

- `src/core`: structured document engine
- `src/converters`: import/export adapters
- `src/routes`: HTTP API surface
- `src/utils`: runtime, filesystem, validation helpers
- `fixtures`: real-world HWP/HWPX regression files

## Host Integration Model

The host app should treat `ooeditor-hwp` exactly as an installable external extension:

- separate install directory
- separate version tracking
- separate update policy
- separate process lifecycle
- separate logs

Do **not** hide `ooeditor-hwp` under the `oo-editors` install tree.

## Release and Versioning

Recommended versioning:

- semver for the extension bundle
- explicit compatibility metadata for the host app

Example:

```json
{
  "name": "ooeditor-hwp",
  "version": "0.4.0",
  "apiVersion": "1",
  "compatibleHost": ">=0.6.0"
}
```

## Safety Requirements

Production requirements for every release:

- real fixture corpus regression tests
- save -> reopen verification
- atomic write only
- checkpoint restore path validated
- `.hwp` import fixtures from real customer-style forms
- merged-table mutation fixtures
- image/header/footer fixtures

## Relationship to oo-editors

`ooeditor-hwp` and `oo-editors` should coexist.

- `oo-editors`: docx/xlsx/pptx/pdf/md and OOXML-centric review/edit flows
- `ooeditor-hwp`: hwp/hwpx open, automate, export, handoff

The host app should decide which extension handles which file family.

## Recommended Next Step

The next implementation step is to formalize a generic extension manager in the host app so both `oo-editors` and `ooeditor-hwp` use the same install/update/runtime contract.
