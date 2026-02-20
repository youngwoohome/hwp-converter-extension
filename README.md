# HWP Converter Extension

Standalone converter/editor service for `hwp` and `hwpx` documents.

This project follows the OfficeExtension-like API shape used in the app:

- `GET /healthcheck`
- `POST /converter`
- `GET /download/:token`
- `GET /open?filepath=...`
- `GET /document?filepath=...`
- `POST /document/save`

## Why this exists

The app currently relies on OfficeExtension/oo-editors for office flows.
This repository adds a dedicated HWP/HWPX path so Korean document support can evolve independently.

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

### `GET /document?filepath=...`

Reads a `.hwp` or `.hwpx` file and returns extracted text + paragraphs.

### `POST /document/save`

Writes edited text back to a document.

Request:

```json
{
  "filePath": "/absolute/path/to/file.hwpx",
  "text": "updated content"
}
```

Behavior:

- `.hwpx`: saved in place (or to `outputPath` if provided)
- `.hwp`: binary write is not implemented; saves `.hwpx` working copy and returns warning
- If `.hwp` save receives non-`.hwpx` `outputPath`, it is normalized to `.hwpx`

### `GET /open?filepath=...`

Serves a simple browser editor UI that loads/saves via `/document` APIs.

## Conversion Support

- `hwpx`:
  - Extracts text from zipped XML sections under `Contents/section*.xml`
  - Converts to `txt/md/html/json/docx/pdf`
  - Save-back for editing supported

- `hwp`:
  - Uses `hwp5txt` (from `pyhwp`) for text extraction
  - Converts extracted text to `txt/md/html/json/docx/pdf`
  - Save-back currently writes `.hwpx` working copy

If `hwp5txt` is missing, `.hwp` reads/conversions fail with an actionable installation error.

## Run

```bash
pnpm install
pnpm dev
```

Default port: `8090` (override with `PORT=8080` if needed).

## Notes

- This service is designed for external/manual installation and run.
- For app integration, point your app-side HWP extension client to `http://localhost:8090`.
