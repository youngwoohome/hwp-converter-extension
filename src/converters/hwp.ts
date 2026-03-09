import { execFile } from 'node:child_process';
import { unlink } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import { promisify } from 'node:util';
import type { ConvertContext, ConvertedArtifact, TextExtractionResult } from '../types.js';
import { writeConvertedOutput } from './common.js';
import { saveHwpxText } from './hwpx.js';
import { convertHwpToHwpxWithJvm } from '../core/jvmCore.js';
import { convertStructuredHwpxToMarkdown } from '../core/markdownExport.js';
import { writeTempFile } from '../utils/files.js';

const execFileAsync = promisify(execFile);

async function commandExists(command: string): Promise<boolean> {
  const locator = process.platform === 'win32' ? 'where' : 'which';
  try {
    await execFileAsync(locator, [command]);
    return true;
  } catch {
    return false;
  }
}

async function extractWithHwp5txt(sourcePath: string): Promise<string> {
  const installed = await commandExists('hwp5txt');
  if (!installed) {
    throw new Error(
      'hwp5txt is required for .hwp extraction. Install pyhwp (pip install pyhwp) and make sure hwp5txt is on PATH.'
    );
  }

  const { stdout } = await execFileAsync('hwp5txt', [sourcePath], {
    maxBuffer: 64 * 1024 * 1024,
  });

  if (!stdout || stdout.trim().length === 0) {
    throw new Error('hwp5txt returned empty output. The file may be encrypted or unsupported.');
  }

  return stdout;
}

export async function extractHwpText(sourcePath: string): Promise<TextExtractionResult> {
  const raw = await extractWithHwp5txt(sourcePath);
  const normalized = raw.replace(/\r\n/g, '\n').trim();

  const paragraphs = normalized
    .split(/\n\s*\n/g)
    .map((p) => p.replace(/\s+/g, ' ').trim())
    .filter((p) => p.length > 0);

  return {
    paragraphs,
    rawText: normalized,
  };
}

export async function saveHwpTextAsHwpx(sourcePath: string, text: string, outputPath?: string): Promise<string> {
  const defaultPath = join(dirname(sourcePath), `${basename(sourcePath, '.hwp')}.hwpx`);
  const requestedPath = outputPath || defaultPath;
  const targetPath = requestedPath.toLowerCase().endsWith('.hwpx')
    ? requestedPath
    : `${requestedPath.replace(/\.[^/.]+$/i, '')}.hwpx`;
  return saveHwpxText(targetPath, text, targetPath);
}

export async function convertHwp(context: ConvertContext): Promise<ConvertedArtifact> {
  if (context.targetFormat === 'md') {
    const tempHwpxPath = await writeTempFile(Buffer.alloc(0), '.hwpx');
    try {
      await convertHwpToHwpxWithJvm(context.sourcePath, tempHwpxPath);
      const converted = await convertStructuredHwpxToMarkdown({
        ...context,
        sourcePath: tempHwpxPath,
      });
      await unlink(tempHwpxPath).catch(() => undefined);
      return {
        ...converted,
        warnings: [
          ...(converted.warnings ?? []),
          {
            code: 'HWP_IMPORTED_TO_HWPX_FOR_MARKDOWN',
            message: 'Markdown was exported from an imported HWPX working model to preserve structure such as tables where possible.',
          },
        ],
      };
    } catch {
      await unlink(tempHwpxPath).catch(() => undefined);
    }
  }

  const extracted = await extractHwpText(context.sourcePath);
  return writeConvertedOutput({
    context,
    paragraphs: extracted.paragraphs,
    rawText: extracted.rawText,
    metadata: {
      extractor: 'hwp5txt',
      source: 'hwp',
    },
  });
}
