import express, { type Request } from 'express';
import { unlink } from 'node:fs/promises';
import { extname } from 'node:path';
import { nanoid } from 'nanoid';
import type { OfficeExtensionConvertRequest, OfficeExtensionConvertResponse, TargetFormat } from './types.js';
import { convertWithSource } from './converters/index.js';
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

app.get('/healthcheck', (_req, res) => {
  res.status(200).send('true');
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

app.get('/open', (_req, res) => {
  res.status(501).json({
    success: false,
    error: 'This repository currently provides conversion only. /open editor view is not implemented.',
  });
});

setInterval(() => {
  void fileStore.cleanupExpired();
}, 60_000).unref();

app.listen(port, () => {
  console.log(`[hwp-converter-extension] listening on http://localhost:${port}`);
});
