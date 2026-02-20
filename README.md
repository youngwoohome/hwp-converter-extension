# HWP Converter Extension

Standalone converter service for `hwp` and `hwpx` documents.

This project follows the existing OfficeExtension request shape used in the app:

- `POST /converter`
- `GET /healthcheck`
- `GET /download/:token`

## Why this exists

The current app flow relies on OfficeExtension/oo-editors for many office conversions.
This repository provides a dedicated converter path for HWP/HWPX so we can evolve Korean document support independently.

## API Contract

### `GET /healthcheck`

Returns plain text `true`.

### `POST /converter`

Request body (OfficeExtension-compatible):

```json
{
  "filetype": "hwpx",
  "outputtype": "pdf",
  "filePath": "/absolute/path/to/input.hwpx",
  "outputPath": "/absolute/path/to/output.pdf"
}
```

Fields:

- `filetype`: source type (`hwp` or `hwpx`)
- `outputtype`: target type (`txt`, `md`, `html`, `json`, `docx`, `pdf`)
- `filePath` or `url`: source input
- `outputPath` (optional): when omitted, a temp file is created and `url` is returned

Response:

```json
{
  "success": true,
  "outputPath": "/tmp/hwp-converter-...pdf",
  "url": "http://localhost:8090/download/<token>"
}
```

### `GET /download/:token`

Downloads a temp output file generated from `/converter` when `outputPath` was omitted.

## Conversion Support

- `hwpx`:
  - Extracts text from zipped XML sections under `Contents/section*.xml`
  - Converts to `txt/md/html/json/docx/pdf`

- `hwp`:
  - Uses `hwp5txt` (from `pyhwp`) for text extraction
  - Then converts extracted text to `txt/md/html/json/docx/pdf`

If `hwp5txt` is missing, `.hwp` conversions fail with an actionable installation error.

## Run

```bash
pnpm install
pnpm dev
```

Default port: `8090` (override with `PORT=8080` if you want drop-in replacement behavior).

## Notes

- This repo currently provides conversion only.
- `/open` editor endpoint is intentionally not implemented yet.
