import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { XMLValidator } from 'fast-xml-parser';
import type { ConvertContext, ConvertedArtifact, TextExtractionResult } from '../types.js';
import { writeConvertedOutput } from './common.js';

function decodeEntities(value: string): string {
  return value
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&#10;', '\n')
    .replaceAll('&#13;', '\r')
    .replaceAll('&#9;', '\t')
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)));
}

function cleanInlineText(xmlChunk: string): string {
  return decodeEntities(
    xmlChunk
      .replace(/<hp:lineBreak\s*\/?>/gi, '\n')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
  );
}

function extractParagraphsFromSectionXml(xml: string): string[] {
  const hpParagraphs = Array.from(xml.matchAll(/<hp:p\b[\s\S]*?<\/hp:p>/gi)).map((m) => cleanInlineText(m[0]));
  const genericParagraphs = Array.from(xml.matchAll(/<p\b[\s\S]*?<\/p>/gi)).map((m) => cleanInlineText(m[0]));

  const candidate = hpParagraphs.length > 0 ? hpParagraphs : genericParagraphs;
  const filtered = candidate.filter((line) => line.length > 0);
  if (filtered.length > 0) return filtered;

  const fallback = cleanInlineText(xml);
  return fallback ? [fallback] : [];
}

function parseSectionOrder(name: string): number {
  const match = name.match(/section(\d+)\.xml$/i);
  return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
}

export async function extractHwpxText(sourcePath: string): Promise<TextExtractionResult> {
  const buffer = await readFile(sourcePath);
  const zip = await JSZip.loadAsync(buffer);

  const sectionFiles = Object.keys(zip.files)
    .filter((name) => /(^|\/)contents\//i.test(name))
    .filter((name) => /section\d+\.xml$/i.test(name))
    .sort((a, b) => parseSectionOrder(a) - parseSectionOrder(b));

  if (sectionFiles.length === 0) {
    throw new Error('No HWPX section XML files were found under Contents/.');
  }

  const paragraphs: string[] = [];

  for (const sectionName of sectionFiles) {
    const sectionFile = zip.file(sectionName);
    if (!sectionFile) continue;

    const xml = await sectionFile.async('string');
    const validated = XMLValidator.validate(xml);
    if (validated !== true) {
      throw new Error(`Invalid XML in section: ${sectionName}`);
    }

    paragraphs.push(...extractParagraphsFromSectionXml(xml));
  }

  const cleaned = paragraphs
    .map((p) => p.replace(/\s+/g, ' ').trim())
    .filter((p) => p.length > 0);

  return {
    paragraphs: cleaned,
    rawText: cleaned.join('\n\n'),
  };
}

export async function convertHwpx(context: ConvertContext): Promise<ConvertedArtifact> {
  const extracted = await extractHwpxText(context.sourcePath);
  return writeConvertedOutput({
    context,
    paragraphs: extracted.paragraphs,
    rawText: extracted.rawText,
    metadata: {
      extractor: 'zip+xml',
      source: 'hwpx',
    },
  });
}
