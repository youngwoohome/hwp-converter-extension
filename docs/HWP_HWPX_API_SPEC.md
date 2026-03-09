# HWP/HWPX Service API Specification

Status: Proposed  
Date: 2026-03-07  
Owner: `hwp-converter-extension`

This document defines the production API contract for the HWP/HWPX service described in [PRODUCTION_HWP_HWPX_ARCHITECTURE.md](./PRODUCTION_HWP_HWPX_ARCHITECTURE.md).

The primary goal is to give the app a stable, production-safe contract while allowing the underlying engine implementation to change.

## 1. Design Principles

- The API must separate read, import, edit, save, and export concerns.
- `.hwp` and `.hwpx` must not be treated identically in mutation flows.
- Every write operation must be checkpoint-aware and verification-backed.
- The client must be able to distinguish:
  - original file
  - editable working copy
  - derived export
- Errors must be structured and machine-readable.

## 2. Terminology

- `sourcePath`
  - path to the user-selected source document
- `workingCopyPath`
  - path to the canonical editable `.hwpx` copy used for mutation
- `documentId`
  - stable per-session identifier for an opened document context
- `checkpointId`
  - identifier for a recoverable saved state
- `engineVersion`
  - version string for the authoritative parser/writer sidecar
- `documentMode`
  - one of: `read_only`, `editable`, `import_required`

## 3. Transport

The service remains HTTP-based for compatibility with the current extension shape.

- Base URL:
  - `http://127.0.0.1:<port>`
- Content type:
  - request: `application/json`
  - response: `application/json`
- Large file upload is out of scope for v1.
- Local file paths are the primary transport for source documents.

## 4. Common Response Envelope

All JSON responses should use this base shape:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "warnings": []
}
```

On failure:

```json
{
  "success": false,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "error": {
    "code": "UNSUPPORTED_FEATURE",
    "message": "The document contains an unsupported editable feature.",
    "details": {
      "feature": "embedded-ole-object"
    },
    "retryable": false
  },
  "warnings": []
}
```

## 5. Error Codes

The service must only emit known error codes.

- `INVALID_REQUEST`
- `FILE_NOT_FOUND`
- `UNSUPPORTED_EXTENSION`
- `UNSUPPORTED_FILE_STRUCTURE`
- `DOCUMENT_TOO_LARGE`
- `ZIP_BOMB_DETECTED`
- `ENGINE_UNAVAILABLE`
- `ENGINE_TIMEOUT`
- `PARSE_FAILED`
- `IMPORT_FAILED`
- `SAVE_FAILED`
- `VALIDATION_FAILED`
- `UNSUPPORTED_FEATURE`
- `READ_ONLY_SOURCE`
- `CHECKPOINT_NOT_FOUND`
- `RECOVERY_FAILED`
- `PERMISSION_DENIED`
- `INTERNAL_ERROR`

## 6. Format Policy Exposed to Clients

- `.hwpx`
  - may open as `editable`
- `.hwp`
  - opens as `read_only` or `import_required`
- `.hwp` edit request
  - requires explicit working-copy creation

The service must never return an editable session for `.hwp` unless it is bound to a `.hwpx` working copy.

## 7. Endpoints

## 7.1 `GET /healthcheck`

Purpose:

- verify gateway liveness
- verify sidecar availability

Response:

```json
{
  "success": true,
  "engineVersion": "1.0.0",
  "status": "ok",
  "sidecar": {
    "healthy": true,
    "ready": true
  }
}
```

## 7.2 `POST /document/open`

Purpose:

- open a source document and return its mode and metadata

Request:

```json
{
  "path": "/absolute/path/to/file.hwp"
}
```

Response:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "document": {
    "documentId": "doc_123",
    "sourcePath": "/absolute/path/to/file.hwp",
    "workingCopyPath": null,
    "sourceFormat": "hwp",
    "canonicalFormat": "hwpx",
    "documentMode": "import_required",
    "readOnly": true,
    "title": "file.hwp",
    "features": {
      "hasTables": true,
      "hasImages": true,
      "hasComments": false,
      "hasUnsupportedEditableFeatures": false
    }
  },
  "warnings": []
}
```

Rules:

- `.hwpx` may return `documentMode=editable`
- `.hwp` returns `documentMode=read_only` or `import_required`

## 7.3 `GET /document`

Purpose:

- fetch the normalized, read-safe document representation for viewing

Query:

- `documentId`

Response:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "document": {
    "documentId": "doc_123",
    "sourceFormat": "hwpx",
    "documentMode": "editable",
    "body": {
      "sections": [
        {
          "blocks": [
            {
              "type": "paragraph",
              "runs": [
                { "text": "Hello", "styleRef": "Body" }
              ]
            }
          ]
        }
      ]
    }
  },
  "warnings": []
}
```

Notes:

- This endpoint returns normalized structure, not raw file internals.
- The returned body must be sufficient for UI rendering and targeted edits.

## 7.4 `GET /document/features`

Purpose:

- inspect document feature support and editing capability

Query:

- `documentId`

Response:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "features": {
    "editable": ["paragraph", "run", "table", "cell"],
    "preservedReadOnly": ["image", "comment", "header", "footer"],
    "unsupported": ["oleObject"]
  },
  "warnings": []
}
```

## 7.5 `POST /document/fork-editable-copy`

Purpose:

- create an editable `.hwpx` working copy from a source document

Request:

```json
{
  "documentId": "doc_123",
  "outputPath": "/absolute/path/to/file.edit.hwpx"
}
```

Response:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "document": {
    "documentId": "doc_123",
    "sourcePath": "/absolute/path/to/file.hwp",
    "workingCopyPath": "/absolute/path/to/file.edit.hwpx",
    "sourceFormat": "hwp",
    "canonicalFormat": "hwpx",
    "documentMode": "editable",
    "readOnly": false
  },
  "warnings": [
    {
      "code": "WORKING_COPY_CREATED",
      "message": "The original .hwp file remains unchanged. Edits apply to the .hwpx working copy."
    }
  ]
}
```

Rules:

- Required before any edit on `.hwp`
- Optional for `.hwpx`
- Must not overwrite the original `.hwp`

## 7.6 `POST /document/create`

Purpose:

- create a blank canonical document

Request:

```json
{
  "format": "hwpx",
  "outputPath": "/absolute/path/to/new-document.hwpx",
  "title": "New document"
}
```

Response:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "document": {
    "documentId": "doc_456",
    "sourcePath": "/absolute/path/to/new-document.hwpx",
    "workingCopyPath": "/absolute/path/to/new-document.hwpx",
    "sourceFormat": "hwpx",
    "canonicalFormat": "hwpx",
    "documentMode": "editable",
    "readOnly": false
  },
  "warnings": []
}
```

Rules:

- v1 supports only `format=hwpx`
- `format=hwp` is not supported in v1

## 7.7 `POST /document/save`

Purpose:

- apply mutations to an editable `.hwpx` target

Request:

```json
{
  "documentId": "doc_123",
  "baseCheckpointId": "ckpt_001",
  "mutations": [
    {
      "op": "replace_text_in_paragraph",
      "sectionIndex": 0,
      "blockIndex": 0,
      "text": "Updated text"
    }
  ]
}
```

Response:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "document": {
    "documentId": "doc_123",
    "sourcePath": "/absolute/path/to/file.hwpx",
    "workingCopyPath": "/absolute/path/to/file.hwpx",
    "sourceFormat": "hwpx",
    "canonicalFormat": "hwpx",
    "documentMode": "editable",
    "readOnly": false
  },
  "save": {
    "checkpointId": "ckpt_002",
    "validationSummary": {
      "reopenVerified": true,
      "preservedUnsupportedNodes": true
    }
  },
  "warnings": []
}
```

Rules:

- Save is only valid for an editable `.hwpx` target.
- `baseCheckpointId` is required when the session already has a checkpoint.
- Save must fail with `READ_ONLY_SOURCE` for a raw `.hwp` document.

## 7.8 `GET /document/checkpoints`

Purpose:

- list available checkpoints for a document

Query:

- `documentId`

Response:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "checkpoints": [
    {
      "checkpointId": "ckpt_001",
      "createdAt": "2026-03-07T10:00:00.000Z",
      "path": "/absolute/path/to/checkpoints/file.hwpx.ckpt1"
    }
  ],
  "warnings": []
}
```

## 7.9 `POST /document/recover`

Purpose:

- restore a prior checkpoint into the current working copy

Request:

```json
{
  "documentId": "doc_123",
  "checkpointId": "ckpt_001"
}
```

Response:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "restored": {
    "checkpointId": "ckpt_001",
    "outputPath": "/absolute/path/to/file.hwpx"
  },
  "warnings": []
}
```

## 7.10 `POST /converter`

Purpose:

- export a source or working copy to another format

Request:

```json
{
  "path": "/absolute/path/to/file.hwpx",
  "targetFormat": "pdf",
  "outputPath": "/absolute/path/to/file.pdf"
}
```

Response:

```json
{
  "success": true,
  "requestId": "req_123",
  "engineVersion": "1.0.0",
  "conversion": {
    "sourcePath": "/absolute/path/to/file.hwpx",
    "sourceFormat": "hwpx",
    "targetFormat": "pdf",
    "outputPath": "/absolute/path/to/file.pdf"
  },
  "warnings": []
}
```

Supported v1 target formats:

- `pdf`
- `html`
- `md`
- `json`
- `docx`
- `txt`

Rules:

- Exports are derived outputs only.
- `targetFormat=hwp` is not supported.

## 7.11 `GET /open`

Purpose:

- open the document in the extension-provided web UI

Query:

- `filepath`

Rules:

- Read-only for raw `.hwp`
- Editable only for `.hwpx` or working-copy-backed sessions

## 8. Mutation Model

The service should not accept raw full-document text replacement as the long-term write API.

v1 mutation set:

- `replace_text_in_paragraph`
- `insert_paragraph_after`
- `delete_paragraph`
- `replace_table_cell_text`
- `insert_table_row`
- `delete_table_row`

Deferred mutations:

- image insertion
- comment insertion
- header/footer editing
- style editing
- embedded object manipulation

## 9. State Transitions

### 9.1 `.hwpx`

```text
open -> editable -> save -> checkpointed
```

### 9.2 `.hwp`

```text
open -> read_only/import_required -> fork-editable-copy -> editable(.hwpx) -> save -> checkpointed
```

Invalid transition:

```text
open raw .hwp -> save directly to same .hwp
```

This must return `READ_ONLY_SOURCE`.

## 10. Validation Contract

Every save must produce a validation summary that covers:

- parse success after write
- canonical structure present
- edited nodes present
- no structural corruption detected

Every export must validate:

- output file exists
- output file is readable by the relevant validator where possible

## 11. Logging Contract

Every request should emit:

- `requestId`
- `documentId` if present
- operation
- source format
- target format if present
- latency
- engine version
- result

Document text must not be logged by default.

## 12. Backward Compatibility

The current prototype endpoints:

- `GET /document?filepath=...`
- `POST /document/save`
- `POST /converter`
- `GET /open?filepath=...`

may continue to exist in v1, but they should internally route to the new document session model.

Compatibility policy:

- Old endpoints remain supported for one release train.
- New code should use the session-oriented endpoints.

## 13. Acceptance Criteria

The API is ready for v1 implementation when:

- all routes in this document have request/response schemas
- `.hwp` cannot be accidentally overwritten
- `.hwpx` save path is checkpointed and reopen-verified
- the app can tell source vs working copy vs export
- errors are structured and stable
