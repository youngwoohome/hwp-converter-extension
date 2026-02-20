import express, { type Request } from 'express';
import { unlink } from 'node:fs/promises';
import { extname } from 'node:path';
import { nanoid } from 'nanoid';
import type { OfficeExtensionConvertRequest, OfficeExtensionConvertResponse, TargetFormat } from './types.js';
import { convertWithSource } from './converters/index.js';
import { extractHwpText, saveHwpTextAsHwpx } from './converters/hwp.js';
import { extractHwpxText, saveHwpxText } from './converters/hwpx.js';
import {
  fileExists,
  guessSourceFormat,
  normalizeAbsolutePath,
  normalizeTargetFormat,
  writeTempFile,
} from './utils/files.js';
import { EphemeralFileStore } from './utils/fileStore.js';

const app = express();
const port = Number(process.env.PORT || 8090);
const fileStore = new EphemeralFileStore(15 * 60 * 1000);

const SUPPORTED_TARGETS = new Set<TargetFormat>(['txt', 'md', 'html', 'json', 'docx', 'pdf']);

app.use(express.json({ limit: '25mb' }));

function baseUrl(req: Request): string {
  const forwardedProto = req.headers['x-forwarded-proto'];
  const proto = typeof forwardedProto === 'string' ? forwardedProto.split(',')[0] : req.protocol;
  return `${proto}://${req.get('host')}`;
}

function downloadContentType(path: string): string {
  const ext = extname(path).toLowerCase();
  if (ext === '.pdf') return 'application/pdf';
  if (ext === '.docx') return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  if (ext === '.json') return 'application/json; charset=utf-8';
  if (ext === '.html') return 'text/html; charset=utf-8';
  if (ext === '.md') return 'text/markdown; charset=utf-8';
  return 'text/plain; charset=utf-8';
}

async function resolveSourceFile(request: OfficeExtensionConvertRequest): Promise<string> {
  if (request.filePath && request.filePath.trim().length > 0) {
    return normalizeAbsolutePath(request.filePath);
  }

  if (!request.url || request.url.trim().length === 0) {
    throw new Error('Either filePath or url is required.');
  }

  const response = await fetch(request.url);
  if (!response.ok) {
    throw new Error(`Failed to download source file from url: ${response.status} ${response.statusText}`);
  }

  const sourceHint = request.filetype?.toLowerCase() || extname(new URL(request.url).pathname).replace('.', '').toLowerCase() || 'bin';
  const extension = sourceHint.startsWith('.') ? sourceHint : `.${sourceHint}`;

  const data = Buffer.from(await response.arrayBuffer());
  return writeTempFile(data, extension);
}

function renderOpenPage(filePath: string): string {
  const safePath = JSON.stringify(filePath);
  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>HWP Extension Editor</title>
  <style>
    :root { color-scheme: light dark; }
    body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #111; color: #eee; }
    .wrap { display: flex; flex-direction: column; height: 100vh; }
    .toolbar { display: flex; align-items: center; gap: 8px; padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.15); background: rgba(0,0,0,0.2); }
    .path { font-size: 12px; opacity: 0.8; flex: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    button { border: 1px solid rgba(255,255,255,0.3); border-radius: 6px; padding: 6px 10px; background: rgba(255,255,255,0.06); color: inherit; cursor: pointer; }
    button:hover { background: rgba(255,255,255,0.14); }
    .status { font-size: 12px; opacity: 0.85; }
    textarea { flex: 1; width: 100%; box-sizing: border-box; border: 0; outline: none; padding: 14px; resize: none; background: #161616; color: #f4f4f4; font: 14px/1.6 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="toolbar">
      <div class="path" id="path"></div>
      <button id="reloadBtn">Reload</button>
      <button id="saveBtn">Save</button>
      <div class="status" id="status">Loading...</div>
    </div>
    <textarea id="editor" spellcheck="false"></textarea>
  </div>

  <script>
    const filePath = ${safePath};
    const pathEl = document.getElementById('path');
    const statusEl = document.getElementById('status');
    const editorEl = document.getElementById('editor');
    const reloadBtn = document.getElementById('reloadBtn');
    const saveBtn = document.getElementById('saveBtn');

    pathEl.textContent = filePath;

    function setStatus(msg) {
      statusEl.textContent = msg;
    }

    async function loadDocument() {
      setStatus('Loading...');
      try {
        const res = await fetch('/document?filepath=' + encodeURIComponent(filePath));
        const json = await res.json();
        if (!res.ok || !json.success) {
          throw new Error(json.error || 'Failed to load document');
        }
        editorEl.value = json.text || '';
        setStatus('Loaded');
        window.parent?.postMessage({ type: 'ONLYOFFICE_DOCUMENT_READY' }, '*');
      } catch (error) {
        setStatus('Load failed');
        alert(error.message || 'Failed to load document');
      }
    }

    async function saveDocument() {
      setStatus('Saving...');
      try {
        const res = await fetch('/document/save', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ filePath, text: editorEl.value }),
        });
        const json = await res.json();
        if (!res.ok || !json.success) {
          throw new Error(json.error || 'Failed to save document');
        }

        if (json.warning) {
          setStatus('Saved with warning');
          alert(json.warning + (json.outputPath ? '\n\nOutput: ' + json.outputPath : ''));
        } else {
          setStatus('Saved');
        }

        window.parent?.postMessage({
          type: 'HWP_EXTENSION_DOCUMENT_SAVED',
          filePath: json.outputPath || filePath,
          originalPath: filePath,
        }, '*');
      } catch (error) {
        setStatus('Save failed');
        alert(error.message || 'Failed to save document');
      }
    }

    reloadBtn.addEventListener('click', () => { void loadDocument(); });
    saveBtn.addEventListener('click', () => { void saveDocument(); });

    void loadDocument();
  </script>
</body>
</html>`;
}

app.get('/healthcheck', (_req, res) => {
  res.status(200).send('true');
});

app.get('/document', async (req, res) => {
  try {
    const filePathRaw = req.query.filepath;
    const filePath = typeof filePathRaw === 'string' ? normalizeAbsolutePath(filePathRaw) : '';
    if (!filePath) {
      res.status(400).json({ success: false, error: 'filepath query is required.' });
      return;
    }

    const exists = await fileExists(filePath);
    if (!exists) {
      res.status(404).json({ success: false, error: `File not found: ${filePath}` });
      return;
    }

    const sourceFormat = guessSourceFormat(undefined, filePath);
    if (!sourceFormat) {
      res.status(400).json({ success: false, error: 'Only .hwp/.hwpx files are supported.' });
      return;
    }

    const extracted = sourceFormat === 'hwpx'
      ? await extractHwpxText(filePath)
      : await extractHwpText(filePath);

    res.json({
      success: true,
      sourceFormat,
      text: extracted.rawText,
      paragraphs: extracted.paragraphs,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Failed to read document';
    res.status(500).json({ success: false, error: message });
  }
});

app.post('/document/save', async (req, res) => {
  try {
    const filePathRaw = req.body?.filePath;
    const textRaw = req.body?.text;
    const outputPathRaw = req.body?.outputPath;

    if (typeof filePathRaw !== 'string' || filePathRaw.trim().length === 0) {
      res.status(400).json({ success: false, error: 'filePath is required.' });
      return;
    }

    const filePath = normalizeAbsolutePath(filePathRaw);
    const text = typeof textRaw === 'string' ? textRaw : '';
    const outputPath = typeof outputPathRaw === 'string' && outputPathRaw.trim().length > 0
      ? normalizeAbsolutePath(outputPathRaw)
      : undefined;

    const sourceFormat = guessSourceFormat(undefined, filePath);
    if (!sourceFormat) {
      res.status(400).json({ success: false, error: 'Only .hwp/.hwpx files are supported.' });
      return;
    }

    if (sourceFormat === 'hwpx') {
      const savedPath = await saveHwpxText(filePath, text, outputPath || filePath);
      res.json({ success: true, outputPath: savedPath });
      return;
    }

    const savedPath = await saveHwpTextAsHwpx(filePath, text, outputPath);
    const outputChanged = !!outputPath && outputPath !== savedPath;
    res.json({
      success: true,
      outputPath: savedPath,
      warning: outputChanged
        ? '.hwp binary write is not supported. outputPath was normalized to a .hwpx working copy.'
        : '.hwp binary write is not supported yet. Saved as .hwpx working copy instead.',
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Failed to save document';
    res.status(500).json({ success: false, error: message });
  }
});

app.post('/converter', async (req, res) => {
  const request = req.body as Partial<OfficeExtensionConvertRequest>;

  try {
    if (!request.filetype || !request.outputtype) {
      const badRequest: OfficeExtensionConvertResponse = {
        success: false,
        error: 'filetype and outputtype are required.',
      };
      res.status(400).json(badRequest);
      return;
    }

    const sourcePath = await resolveSourceFile(request as OfficeExtensionConvertRequest);
    const sourceFormat = guessSourceFormat(request.filetype, sourcePath);
    if (!sourceFormat) {
      throw new Error(`Unsupported source format: ${request.filetype}. Only hwp/hwpx are supported.`);
    }

    const exists = await fileExists(sourcePath);
    if (!exists) {
      throw new Error(`Source file does not exist: ${sourcePath}`);
    }

    const target = normalizeTargetFormat(request.outputtype);
    if (!SUPPORTED_TARGETS.has(target as TargetFormat)) {
      throw new Error(`Unsupported target format: ${request.outputtype}. Supported: ${Array.from(SUPPORTED_TARGETS).join(', ')}`);
    }

    const result = await convertWithSource({
      sourcePath,
      sourceFormat,
      targetFormat: target as TargetFormat,
      requestedOutputPath: request.outputPath,
    });

    if (request.outputPath && request.outputPath.trim().length > 0) {
      const response: OfficeExtensionConvertResponse = {
        success: true,
        outputPath: result.outputPath,
      };
      res.json(response);
      return;
    }

    const token = nanoid(16);
    fileStore.put(token, result.outputPath);

    const response: OfficeExtensionConvertResponse = {
      success: true,
      outputPath: result.outputPath,
      url: `${baseUrl(req)}/download/${token}`,
    };

    res.json(response);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown conversion failure';
    const response: OfficeExtensionConvertResponse = {
      success: false,
      error: message,
      details: message,
    };
    res.status(500).json(response);
  }
});

app.get('/download/:token', async (req, res) => {
  try {
    const token = req.params.token;
    const stored = await fileStore.consume(token);

    if (!stored) {
      res.status(404).json({ success: false, error: 'File token expired or does not exist.' });
      return;
    }

    res.status(200);
    res.setHeader('Content-Type', downloadContentType(stored.path));
    res.setHeader('Content-Disposition', `attachment; filename="${stored.path.split('/').pop() || 'converted'}"`);

    stored.stream.on('close', async () => {
      try {
        await unlink(stored.path);
      } catch {
        // Ignore cleanup errors
      }
    });

    stored.stream.pipe(res);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Download failed';
    res.status(500).json({ success: false, error: message });
  }
});

app.get('/open', async (req, res) => {
  const filePathRaw = req.query.filepath;
  const filePath = typeof filePathRaw === 'string' ? normalizeAbsolutePath(filePathRaw) : '';

  if (!filePath) {
    res.status(400).send('Missing filepath query parameter.');
    return;
  }

  const exists = await fileExists(filePath);
  if (!exists) {
    res.status(404).send(`File not found: ${filePath}`);
    return;
  }

  res.status(200).setHeader('Content-Type', 'text/html; charset=utf-8').send(renderOpenPage(filePath));
});

setInterval(() => {
  void fileStore.cleanupExpired();
}, 60_000).unref();

app.listen(port, () => {
  console.log(`[hwp-converter-extension] listening on http://localhost:${port}`);
});
