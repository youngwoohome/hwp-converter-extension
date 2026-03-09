# Hancom-Primary Integration Architecture

Status: Proposed  
Date: 2026-03-09  
Owner: `hwp-converter-extension`

## 1. Summary

The product goal is not just to "open HWP files somehow".
The target is:

- HWP/HWPX editing that feels close to native Hancom usage
- document fidelity high enough for real production workflows
- Markdown export that preserves tables, images, and structure as much as possible

The recommended production architecture is:

- use **Hancom Web Hwp editor** as the primary interactive editing surface
- use **Hancom Hwp SDK** as the authoritative programmatic engine for conversion, extraction, and server-side document operations
- keep **HWPX** as the canonical editable storage format
- keep the current open-source `hwp-converter-extension` as the integration shell, session manager, and fallback path

This architecture is a better fit than forking `oo-editors` as the primary HWP solution.

## 2. Why We Are Choosing Hancom as the Primary Engine

The requirement is closer to "Hancom-like document behavior" than to "generic office viewing".

Official Hancom material describes:

- **Hwp SDK**
  - open HWP/HWPX without installing the Hwp desktop program
  - create HWPX documents
  - convert HWP/HWPX to HTML and PDF
  - extract and edit text, image, and table data
  - provide compare/merge style document operations  
  Source: [Hancom Hwp SDK](https://download.hancom.com/product/sdk/hwpSdk)

- **Web Hwp editor**
  - HTML5 browser-based editor
  - HWP-like UI and shortcuts
  - high compatibility on import/export
  - HWP, HWPX, PDF, ODT support
  - browser support including Chrome, Edge, Whale, Firefox, Safari
  - developer guide and API surface through Hancom Developer  
  Sources: [Web Hwp editor product page](https://download.hancom.com/en/product/solution/officesolution/hangulgiangi), [Hancom product page (KR)](https://www.hancom.com/product/solution/officesolution/hangulgiangi), [Hancom Developer Web Hwp](https://developer.hancom.com/webhwp)

Official Hancom documentation also states that:

- legacy `.hwp` is binary and harder to analyze directly
- `HWPX` is the current default format
- `HWPX` is based on OWPML, an open XML standard  
  Source: [Hancom FAQ on HWP vs HWPX](https://www.hancom.com/support/faqCenter/faq/detail/3131)

These points align directly with our product goals.

## 3. Why We Are Not Making `oo-editors` the Primary HWP Strategy

`oo-editors` is a good host pattern but not the right authoritative HWP engine.

Reasons:

- `oo-editors` is optimized around the OOXML/ONLYOFFICE world
- even when HWP/HWPX can be opened in that ecosystem, the behavior is still centered on conversion compatibility rather than Hancom-native semantics
- the user requirement is "close to native Hancom behavior", not just "a document opens in a browser editor"
- a fork of `oo-editors` would still leave us with weaker HWP-specific save fidelity, automation semantics, and format-aware extraction

Decision:

- keep the **distribution and host integration pattern** similar to `oo-editors`
- do **not** make an `oo-editors` fork the primary HWP document engine

## 4. Canonical Product Model

We define the following internal truths:

- original source formats:
  - `.hwp`
  - `.hwpx`
- canonical editable format:
  - `.hwpx`
- canonical editor surface:
  - Hancom Web Hwp editor
- authoritative backend document engine:
  - Hancom Hwp SDK
- host shell and session manager:
  - `hwp-converter-extension`

This means:

- `.hwpx` opens directly in the primary path
- `.hwp` is imported into a managed `.hwpx` working copy before mutation
- the original `.hwp` is never mutated in place by the app

## 5. Target Architecture

```text
host Electron app
  -> HWP Extension Manager
    -> hwp-converter-extension (Node gateway)
      -> Hancom Web Hwp editor runtime (interactive editor UI)
      -> Hancom Hwp SDK worker (authoritative backend operations)
      -> Optional OSS fallback engine (read-only / degraded mode)
```

## 5.1 Layer Responsibilities

### Host Electron app

- chooses the HWP viewer/editor route for `.hwp/.hwpx`
- installs, starts, updates, and stops the HWP extension
- hosts the iframe or webview surface
- manages tabs, refresh, checkpoints UI, and user messaging

### `hwp-converter-extension`

- keeps a stable localhost API for the host app
- manages document sessions and working copies
- brokers requests between the host app and Hancom runtimes
- enforces file safety rules and recovery semantics
- exposes uniform export APIs such as Markdown, HTML, JSON, PDF

### Hancom Web Hwp editor runtime

- provides the main editing UI
- provides Hwp-like commands, editing semantics, layout behavior, and keyboard patterns
- is the only user-facing rich editor surface in production mode

### Hancom Hwp SDK worker

- performs authoritative document open/convert/save operations
- extracts structured content for Markdown export
- performs non-UI server-side tasks such as preview conversion, compare/merge, and asset extraction

### OSS fallback engine

- current DOM/ZIP/JVM prototype remains useful for:
  - local development
  - internal testing
  - degraded read-only operation
  - emergency fallback when Hancom runtime is unavailable
- it is not the primary production path

## 6. Editing UX Strategy

## 6.1 Preferred UX

For users, `.hwp/.hwpx` should feel like a first-class document tab, not a textarea projection.

So the viewer should move from the current lightweight `/open` page to:

- a dedicated HWP iframe surface
- backed by Hancom Web Hwp editor
- with host-level integration for:
  - tab title
  - dirty state
  - autosave status
  - save-as / export actions
  - checkpoint recovery
  - file refresh events

## 6.2 Save Model

- open `.hwpx` directly
- open `.hwp` as source + imported `.hwpx` working copy
- autosave always targets the managed working copy
- explicit save can:
  - overwrite the working `.hwpx`
  - export to a user-chosen location
  - optionally export a `.hwp` copy if and only if Hancom runtime supports a safe export path

## 6.3 Safety Rules

- raw `.hwp` source is immutable in place
- every editable session has a working copy and checkpoints
- every save is:
  - atomic
  - versioned
  - verified after write
- failed vendor operations never silently replace the last known good file

## 7. Markdown Export Strategy

Markdown is not the system of record.
It is a derived format for AI, search, knowledge workflows, and sharing.

Therefore the export path should be:

```text
HWP/HWPX
  -> Hancom-authoritative extraction
  -> internal structured AST
  -> HTML + assets + semantic metadata
  -> Markdown emitter
```

## 7.1 Why Not Direct `HWP -> Markdown`

Direct Markdown conversion is too lossy for:

- merged or nested tables
- floating images
- text boxes and shape objects
- footnotes/endnotes
- comments / revisions
- page layout semantics
- headers and footers

So we should treat Markdown as a projection from a richer internal representation.

## 7.2 Export Output Contract

The recommended export output is a bundle, not a single `.md` file only.

Example:

```text
report.md
report.assets/
  image-001.png
  image-002.png
report.metadata.json
report.diagnostics.json
```

## 7.3 Markdown Emission Rules

### Tables

- simple rectangular tables -> GitHub Flavored Markdown tables
- merged cells / nested structures -> embedded HTML tables inside Markdown
- table captions and metadata -> preserved in `metadata.json`

### Images

- export binary assets to `assets/`
- emit Markdown image references
- preserve alt text and captions when available
- preserve placement metadata in `metadata.json`

### Formatting

- headings, lists, emphasis, inline code -> native Markdown
- unsupported inline styles -> HTML spans only when necessary
- page-only formatting -> omitted from Markdown, preserved in metadata

### Comments / Revisions / Fields

- not forced into lossy inline text
- stored in sidecar metadata
- surface a fidelity warning to the caller

## 7.4 Quality Modes

Support at least two export modes:

- `markdown_mode=clean`
  - cleaner Markdown
  - more aggressive simplification
- `markdown_mode=fidelity`
  - preserves complex blocks using embedded HTML and metadata sidecars

Default should be `fidelity`.

## 8. Integration Contract with the Host App

The host app should continue to treat HWP support as an external extension, similar in lifecycle shape to `oo-editors`.

## 8.1 What Stays the Same

- GitHub/private artifact based extension distribution
- install into app-managed runtime directory
- start a localhost service
- check readiness through `/healthcheck`
- open the editor through `/open` or a session URL

## 8.2 What Changes

- HWP extension becomes a **premium multi-runtime wrapper**, not just a small Node converter
- the editor surface should point to the Hancom editor session, not the current textarea-based fallback page
- The host app should understand:
  - HWP extension install state
  - HWP session state
  - imported working-copy state
  - vendor-runtime health

## 8.3 Recommended Session API Additions

The current API is a good starting point, but a Hancom-primary integration needs explicit editor-session endpoints.

Recommended additions:

- `POST /editor/session/create`
- `GET /editor/session/:id`
- `POST /editor/session/:id/save`
- `POST /editor/session/:id/export`
- `POST /editor/session/:id/checkpoint`
- `POST /editor/session/:id/close`

And keep:

- `GET /healthcheck`
- `POST /document/open`
- `POST /document/fork-editable-copy`
- `POST /converter`

## 9. Packaging and Distribution

This architecture crosses an open-source/proprietary boundary.

That boundary should be explicit.

## 9.1 Open-Source Repository

Public repo contents should include:

- the Node gateway
- host integration code
- OSS fallback engine
- session/recovery/checkpoint logic
- build glue
- documentation

It should **not** include Hancom proprietary binaries or license files.

## 9.2 Proprietary Runtime Packaging

Hancom runtimes should be distributed separately:

- private artifact store
- customer-installed package
- enterprise deployment channel

The extension should support:

- `HANCOM_WEBHWP_HOME`
- `HANCOM_HWP_SDK_HOME`
- `HANCOM_LICENSE_PATH`
- `HANCOM_FONTS_DIR`

## 9.3 Install Modes

Support three install modes:

1. `oss-fallback`
   - no Hancom runtime
   - degraded functionality

2. `hancom-edit`
   - Web Hwp editor available
   - rich edit UI enabled

3. `hancom-full`
   - Web Hwp editor + Hwp SDK available
   - full edit + export + Markdown fidelity path enabled

## 10. Operational Requirements

## 10.1 Fonts

Font fidelity is a major risk for HWP rendering and export quality.

Production requirements:

- configurable customer font directory
- startup validation for required font packs
- export diagnostics when fallback fonts are used
- optional per-document font manifest

## 10.2 Isolation

- vendor runtimes must run out of process
- per-session working directories must be isolated
- no direct access from renderer to proprietary file system paths
- localhost endpoints only
- authenticated session tokens between host and extension

## 10.3 Observability

Track:

- vendor runtime start failures
- import duration
- save duration
- export duration
- fidelity downgrade events
- font substitution events
- checkpoint recovery events

## 11. Decision on HWP vs HWPX

This remains unchanged from the earlier production architecture:

- `HWPX` is the canonical editable format
- `.hwp` is legacy input

The difference is that now the **editing UX** becomes Hancom-native-ish rather than our own custom lightweight editor.

## 12. Proposed Rollout

## Phase 0: Commercial Validation

- confirm commercial terms and deployment rights for Web Hwp editor and Hwp SDK
- validate local embedding, automation, export APIs, and license activation
- validate whether vendor-supported HWP export-back is safe enough for optional `.hwpx -> .hwp` export

## Phase 1: Viewer/Editor Surface

- replace current fallback `/open` production route with Hancom editor session route
- wire session lifecycle into the host app
- support `.hwp -> .hwpx` managed working-copy import

## Phase 2: Safe Save Path

- autosave
- save verification
- checkpointing
- crash recovery
- dirty-state integration

## Phase 3: Markdown Fidelity Path

- authoritative structured extraction
- asset extraction
- HTML intermediary
- fidelity-mode Markdown export
- metadata sidecars

## Phase 4: Tooling and Agent Integration

- `read_hwp`
- `export_hwp_markdown`
- `open_hwp_editor`
- `recover_hwp_session`
- richer semantic extraction for AI workflows

## 13. Final Recommendation

For the stated product goal, the recommended path is:

- **primary editing UX**: Hancom Web Hwp editor
- **primary conversion/extraction engine**: Hancom Hwp SDK
- **canonical editable format**: HWPX
- **host distribution model**: keep the current downloadable-extension model used by the host app
- **Markdown export model**: authoritative extraction -> structured AST -> Markdown bundle with assets and metadata

This gives us the best chance of reaching:

- HWP/HWPX editing that feels close to native Hancom usage
- materially better fidelity than an `oo-editors` fork
- Markdown output that is useful for AI workflows without pretending to be lossless
