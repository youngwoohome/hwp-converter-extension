# ooeditor-hwp Host Integration Plan

## Goal

Integrate `ooeditor-hwp` into a host Electron app as a first-class downloadable extension, parallel to `oo-editors`.

This document defines:

- install/update behavior
- runtime process model
- file routing
- review and automation UX expectations

## Integration Principles

- `ooeditor-hwp` is a separate extension product
- HWP/HWPX should not be routed through `oo-editors` by default
- the host app should expose deterministic automation tools, not a fake Hancom editor
- review UI can be lightweight; authoritative save path remains the HWP extension

## Host Responsibilities

The host app must provide:

- release discovery
- download and extraction
- versioned install directories
- extension healthcheck
- process startup and shutdown
- per-extension logging
- UI routing by file extension

## File Routing

Recommended routing policy:

- `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`, `.pdf`, `.md`: `oo-editors`
- `.hwp`, `.hwpx`: `ooeditor-hwp`

There should be no fallback from `.hwp/.hwpx` into `oo-editors` unless explicitly marked experimental.

## Extension Install Layout

Recommended host-side paths:

- `userData/oo-editors/`
- `userData/ooeditor-hwp/`

Recommended temporary/staging paths:

- `userData/oo-editors-staging/`
- `userData/ooeditor-hwp-staging/`

Recommended logs:

- `logs/oo-editors-*.log`
- `logs/ooeditor-hwp-*.log`

## GitHub Release Model

The host should fetch latest releases from a dedicated GitHub repo for `ooeditor-hwp`.

Required release asset behavior:

- per-architecture zip assets
- embedded version file
- consistent top-level folder structure
- startup entrypoint at a stable path

Example:

- repo: `openinterpreter/ooeditor-hwp`
- asset: `ooeditor-hwp-darwin-arm64.zip`

## Runtime Process Model

The host launches `ooeditor-hwp` as a child process, separate from `oo-editors`.

Recommended behavior:

- start on first HWP/HWPX open
- keep warm while the app session is active
- stop on app quit
- restart on crash
- surface healthcheck and last error state in logs

## Health Contract

Minimum startup contract:

- process spawned
- `GET /healthcheck` returns `true`
- optional `GET /healthcheck?format=json` returns version/capability details

The host should not route files into the viewer until healthcheck passes.

## Viewer Model

Recommended host viewer behavior:

- iframe or webview for `/open?filepath=...`
- receive structured document state via `postMessage`
- allow host-side action panels to call deterministic APIs

The host should treat the embedded UI as:

- a structured preview and operation surface
- not a full WYSIWYG editor

## Automation Surface

The host should expose a tool layer that maps cleanly onto the extension API.

Recommended operations:

- `open_hwp_document`
- `fork_hwp_editable_copy`
- `analyze_hwp_reference`
- `analyze_hwp_repeat_region`
- `fill_hwp_template`
- `save_hwp_document`
- `convert_hwp_document`
- `recover_hwp_document`
- `replace_hwp_image`
- `insert_hwp_image`
- `delete_hwp_image`
- `update_hwp_image_placement`

These are naming examples; they can still route to the existing HTTP endpoints.

## UX Model

Recommended user messaging:

- “Open and automate HWP/HWPX documents”
- “Create editable working copy”
- “Review generated result”
- “Export HWPX for final editing in Hancom Office”

Avoid messaging that implies:

- perfect Hancom compatibility
- full replacement of native Hancom editing

## Review Surface Policy

For HWP/HWPX, the default review surface should come from `ooeditor-hwp`, not `oo-editors`.

If `oo-editors` review remains available, it should be marked:

- experimental
- non-authoritative
- unsupported for save

## Update Policy

Recommended host policy:

- check for updates on app start
- download in background
- swap in only after full extraction and validation
- never mutate a running extension install in place

## Failure Handling

If `ooeditor-hwp` fails to start:

- show a clear extension startup error
- keep the file unopened rather than falling back to the wrong engine
- link to logs or copyable diagnostics

If a save fails:

- preserve checkpoint
- surface recover action
- never silently overwrite the source

## Implementation Plan

### Phase 1

- dedicated `hwp-extension.ts` style service stays separate from `office-extension.ts`
- GitHub release discovery for `ooeditor-hwp`
- install/update/start/stop lifecycle
- `.hwp/.hwpx` routing only to HWP extension

### Phase 2

- generic extension manager abstraction
- unified install/update UI for both extensions
- shared logging and health dashboard

### Phase 3

- tool registry integration for deterministic HWP operations
- richer preview panel and structured action UI

## Recommendation

Do not keep growing HWP support inside the `oo-editors` mental model.
Promote the HWP path into a fully separate product and make the host treat it as such.
