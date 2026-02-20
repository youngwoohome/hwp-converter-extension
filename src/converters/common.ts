import { createWriteStream } from 'node:fs';
import { writeFile } from 'node:fs/promises';
import { once } from 'node:events';
import { extname, join } from 'node:path';
import { tmpdir } from 'node:os';
import { nanoid } from 'nanoid';
import { Document, Packer, Paragraph } from 'docx';
import PDFDocument from 'pdfkit';
import type { ConvertContext, ConvertedArtifact, TargetFormat } from '../types.js';
import { ensureParentDirectory, normalizeAbsolutePath } from '../utils/files.js';

const TARGET_EXTENSIONS: Record<TargetFormat, string> = {
  txt: '.txt',
  md: '.md',
  html: '.html',
  json: '.json',
  docx: '.docx',
  pdf: '.pdf',
};

export interface WriteOutputInput {
  context: ConvertContext;
  paragraphs: string[];
  rawText: string;
  metadata?: Record<string, unknown>;
}

function resolveOutputPath(context: ConvertContext): string {
  const targetExt = TARGET_EXTENSIONS[context.targetFormat];

  if (context.requestedOutputPath && context.requestedOutputPath.trim().length > 0) {
    const normalized = normalizeAbsolutePath(context.requestedOutputPath);
    if (!extname(normalized)) {
      return `${normalized}${targetExt}`;
    }
    return normalized;
  }

  return join(tmpdir(), `hwp-converter-${Date.now()}-${nanoid(8)}${targetExt}`);
}

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function sanitizeParagraphs(paragraphs: string[]): string[] {
  const cleaned = paragraphs
    .map((line) => line.replace(/\s+/g, ' ').trim())
    .filter((line) => line.length > 0);

  if (cleaned.length > 0) return cleaned;
  return [''];
}

async function writeDocx(paragraphs: string[], outputPath: string): Promise<void> {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: sanitizeParagraphs(paragraphs).map((text) => new Paragraph(text)),
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  await writeFile(outputPath, buffer);
}

async function writePdf(paragraphs: string[], outputPath: string): Promise<void> {
  const doc = new PDFDocument({ margin: 48, size: 'A4' });
  const stream = createWriteStream(outputPath);

  doc.pipe(stream);

  const lines = sanitizeParagraphs(paragraphs);
  for (const paragraph of lines) {
    doc.text(paragraph, {
      width: 500,
      align: 'left',
      lineGap: 2,
    });
    doc.moveDown(0.75);
  }

  doc.end();
  await once(stream, 'finish');
}

export async function writeConvertedOutput(input: WriteOutputInput): Promise<ConvertedArtifact> {
  const { context, rawText, metadata } = input;
  const paragraphs = sanitizeParagraphs(input.paragraphs);

  const outputPath = resolveOutputPath(context);
  await ensureParentDirectory(outputPath);

  switch (context.targetFormat) {
    case 'txt': {
      await writeFile(outputPath, `${paragraphs.join('\n\n')}\n`, 'utf-8');
      break;
    }
    case 'md': {
      await writeFile(outputPath, `${paragraphs.join('\n\n')}\n`, 'utf-8');
      break;
    }
    case 'html': {
      const html = [
        '<!doctype html>',
        '<html lang="ko">',
        '<head>',
        '  <meta charset="utf-8" />',
        '  <meta name="viewport" content="width=device-width, initial-scale=1" />',
        '  <title>Converted from HWP/HWPX</title>',
        '  <style>body{font-family:AppleSDGothicNeo,"Noto Sans KR",sans-serif;line-height:1.6;max-width:840px;margin:40px auto;padding:0 16px;}p{margin:0 0 1em;}</style>',
        '</head>',
        '<body>',
        ...paragraphs.map((p) => `  <p>${escapeHtml(p)}</p>`),
        '</body>',
        '</html>',
      ].join('\n');
      await writeFile(outputPath, `${html}\n`, 'utf-8');
      break;
    }
    case 'json': {
      const payload = {
        sourceFormat: context.sourceFormat,
        targetFormat: context.targetFormat,
        paragraphCount: paragraphs.length,
        text: rawText,
        paragraphs,
        metadata: metadata || {},
      };
      await writeFile(outputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf-8');
      break;
    }
    case 'docx': {
      await writeDocx(paragraphs, outputPath);
      break;
    }
    case 'pdf': {
      await writePdf(paragraphs, outputPath);
      break;
    }
    default: {
      const exhaustiveCheck: never = context.targetFormat;
      throw new Error(`Unsupported target format: ${exhaustiveCheck}`);
    }
  }

  return {
    outputPath,
    targetFormat: context.targetFormat,
  };
}
