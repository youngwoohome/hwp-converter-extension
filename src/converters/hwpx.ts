import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { XMLValidator } from 'fast-xml-parser';
import type { ConvertContext, ConvertedArtifact, TextExtractionResult } from '../types.js';
import { writeConvertedOutput } from './common.js';
import { convertStructuredHwpxToMarkdown } from '../core/markdownExport.js';
import { writeValidatedHwpxPackage } from '../core/hwpxPackage.js';

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

function normalizeParagraphs(text: string): string[] {
  const normalized = text.replace(/\r\n/g, '\n');
  return normalized
    .split(/\n\s*\n/g)
    .map((line) => line.replace(/\s+/g, ' ').trim())
    .filter((line) => line.length > 0);
}

function escapeXml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function buildSectionXml(paragraphs: string[]): string {
  const body = paragraphs.length > 0 ? paragraphs : [''];

  const hpParagraphs = body.map((line) => {
    const escaped = escapeXml(line);
    return `<hp:p><hp:run><hp:t>${escaped}</hp:t></hp:run></hp:p>`;
  }).join('');

  return `<?xml version="1.0" encoding="UTF-8"?>\n<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">${hpParagraphs}</hp:section>\n`;
}

async function loadOrCreateZip(sourcePath: string): Promise<JSZip> {
  try {
    const buffer = await readFile(sourcePath);
    return await JSZip.loadAsync(buffer);
  } catch {
    const zip = new JSZip();
    const timestamp = new Date().toISOString();
    zip.file('mimetype', 'application/hwp+zip');
    zip.file(
      'META-INF/container.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
      + '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">'
      + '<ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/></ocf:rootfiles>'
      + '</ocf:container>'
    );
    zip.file(
      'Contents/content.hpf',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
      + '<opf:package xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" version="" unique-identifier="" id="">'
      + '<opf:metadata><opf:title/><opf:language>ko</opf:language>'
      + `<opf:meta name="CreatedDate" content="text">${timestamp}</opf:meta>`
      + `<opf:meta name="ModifiedDate" content="text">${timestamp}</opf:meta>`
      + '</opf:metadata>'
      + '<opf:manifest><opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/></opf:manifest>'
      + '<opf:spine><opf:itemref idref="section0"/></opf:spine>'
      + '</opf:package>'
    );
    zip.file(
      'version.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
      + '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="0" micro="5" buildNumber="0" xmlVersion="1.4" application="hwp-converter-extension" appVersion="0.1.0"/>'
    );
    zip.file(
      'settings.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
      + '<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
      + '<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>'
      + '</ha:HWPApplicationSetting>'
    );
    return zip;
  }
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

export async function saveHwpxText(sourcePath: string, text: string, outputPath?: string): Promise<string> {
  const targetPath = outputPath || sourcePath;
  const zip = await loadOrCreateZip(sourcePath);

  const sectionNames = Object.keys(zip.files).filter((name) => /(^|\/)contents\/section\d+\.xml$/i.test(name));
  for (const sectionName of sectionNames) {
    zip.remove(sectionName);
  }

  const paragraphs = normalizeParagraphs(text);
  const sectionXml = buildSectionXml(paragraphs);
  zip.file('Contents/section0.xml', sectionXml);

  await writeValidatedHwpxPackage(targetPath, zip);
  return targetPath;
}

export async function convertHwpx(context: ConvertContext): Promise<ConvertedArtifact> {
  if (context.targetFormat === 'md') {
    return convertStructuredHwpxToMarkdown(context);
  }

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
