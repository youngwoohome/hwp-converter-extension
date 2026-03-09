# HWP Converter Extension

`hwp-converter-extension` is a standalone local service for working with `HWP` and `HWPX` documents.

It is designed for host apps, agent workflows, and document automation systems that need to:

- open Korean office documents outside Hancom Office
- create safe editable `.hwpx` working copies from `.hwp`
- analyze existing forms and templates
- fill placeholders and tables deterministically
- replace, insert, move, and delete image instances
- export structured Markdown, HTML, JSON, PDF, and DOCX derivatives
- keep checkpoints and recover from failed writes

This project is best thought of as an **HWP/HWPX automation and working-copy engine**, not a full native Hancom Office replacement.

## What You Can Build With It

- an AI agent that fills Korean government or enterprise forms and exports `.hwpx`
- a Mac or browser-based app that can open `.hwp/.hwpx` and hand off the final file to Hancom Office on Windows
- a template engine for batch-generating notices, reports, minutes, or tax forms
- a Markdown extraction pipeline for knowledge systems and LLM workflows
- a host-side extension product such as `ooeditor-hwp`

## Typical Workflow

1. Open a `.hwp` or `.hwpx` document.
2. If the source is `.hwp`, create a managed editable `.hwpx` working copy.
3. Analyze the document or a reference template.
4. Apply deterministic mutations:
   - placeholder fill
   - explicit block or cell updates
   - repeated table row regions
   - image insert, replace, delete, or placement updates
5. Save with checkpoints and verification.
6. Export `.hwpx`, `.pdf`, `Markdown + assets + diagnostics`, or other derivatives.
7. Optionally finish the last manual edits in Hancom Office.

## Current Capabilities

### Documents

- `POST /document/open`
- `POST /document/fork-editable-copy`
- `POST /document/create`
- `POST /document/save`
- `GET /document?documentId=...`
- `GET /document/features?documentId=...`
- `GET /document/checkpoints?documentId=...`
- `POST /document/recover`

### Template Automation

- `GET /document/templates`
- `POST /document/analyze-reference`
- `POST /document/analyze-repeat-region`
- `POST /document/fill-template`

### Conversion

- `POST /converter`
- `GET /download/:token`

### Embedded Viewer Surface

- `GET /open?filepath=...`
- `GET /healthcheck`

## What Makes This Different

- `.hwp` is treated as a source format, not as an unsafe in-place write target
- `.hwpx` is the canonical editable format
- edits are deterministic and fail-closed
- complex tables are handled through structured row-band analysis instead of naive text replacement
- image edits happen at instance or asset scope with explicit locators
- Markdown export is structure-aware and can preserve complex tables as HTML when needed

## What This Is Not

- not a native Hancom Office clone
- not a promise of perfect round-trip fidelity for every document feature
- not a generic WYSIWYG office suite

The intended product model is:

- automate 80-90% of repetitive HWP/HWPX work here
- export a clean `.hwpx`
- finish the last manual polish in Hancom Office when needed

## API Example

### Convert to Markdown with diagnostics

```json
{
  "filetype": "hwpx",
  "outputtype": "md",
  "filePath": "/absolute/path/to/input.hwpx",
  "outputPath": "/absolute/path/to/output.md",
  "markdownMode": "fidelity",
  "includeDiagnostics": true
}
```

### Fill a template deterministically

```json
{
  "documentId": "doc_123",
  "requiredPlaceholders": ["[Meeting Title]"],
  "values": {
    "[Meeting Title]": "Q1 Planning Meeting"
  },
  "fills": [
    { "target": "0:3", "value": "Alice, Bob, Carol" }
  ],
  "tableRepeats": [
    {
      "tableBlockId": "0:1",
      "templateRowIndex": 6,
      "templateEndRowIndex": 8,
      "boundaryPolicy": "split_boundary_merges",
      "rows": [
        ["A", "Seoul", "1000"],
        ["B", "Busan", "2000"]
      ]
    }
  ]
}
```

## Local Development

```bash
pnpm install
pnpm build
pnpm start
```

Default server port is `38124`.

If you are developing the JVM import core too:

```bash
cd jvm-core
mvn package
```

## Release Bundles

This repo includes a GitHub Actions workflow that builds downloadable release bundles for:

- `darwin-arm64`
- `darwin-x64`
- `windows-x64`

Release flow:

1. push your changes
2. create and push a tag such as `v0.1.0`
3. GitHub Actions builds the service bundle
4. GitHub Release assets are published as:
   - `hwp-converter-extension-darwin-arm64.zip`
   - `hwp-converter-extension-darwin-x64.zip`
   - `hwp-converter-extension-windows-x64.zip`

Those assets are suitable for host apps that download and install the extension from GitHub Releases.

## Public Repo Notes

This repository contains the extension and its document engine only.
It does **not** need to include any host app code.

Typical integration model:

- host app downloads a release bundle
- host app starts the service as a local child process
- host app calls the localhost API for open, fill, save, convert, and recovery flows

## Documentation

- [Production HWP/HWPX Architecture](./docs/PRODUCTION_HWP_HWPX_ARCHITECTURE.md)
- [HWP/HWPX Service API Specification](./docs/HWP_HWPX_API_SPEC.md)
- [HWP/HWPX V1 Work Breakdown](./docs/HWP_HWPX_V1_WORK_BREAKDOWN.md)
- [Hancom-Primary Integration Architecture](./docs/HANCOM_PRIMARY_INTEGRATION_ARCHITECTURE.md)
- [HWPX-Rekian Adoption Notes](./docs/HWPX_REKIAN_ADOPTION_NOTES.md)
- [Merged Table Editing Strategy](./docs/MERGED_TABLE_EDITING_STRATEGY.md)
- [ooeditor-hwp Product Architecture](./docs/OOEDITOR_HWP_PRODUCT_ARCHITECTURE.md)
- [ooeditor-hwp Host Integration Plan](./docs/OOEDITOR_HWP_HOST_INTEGRATION_PLAN.md)
