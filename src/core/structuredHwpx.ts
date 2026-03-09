import { copyFile, readFile } from 'node:fs/promises';
import { extname, posix as pathPosix } from 'node:path';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { XMLValidator } from 'fast-xml-parser';
import JSZip from 'jszip';
import type {
  ApiWarning,
  DocumentBlock,
  DocumentFeatureSummary,
  EmbeddedImageReference,
  EmbeddedImageKind,
  NormalizedDocumentBody,
  ParagraphBlock,
  SupplementalTextBlock,
  SupplementalTextKind,
  TableBlock,
  TableCellBlock,
  TextMutationOperation,
} from '../types.js';
import type { HwpxPackageValidationSummary } from './hwpxPackage.js';
import { loadValidatedHwpxPackage, writeValidatedHwpxPackage } from './hwpxPackage.js';

const HP_NS = 'http://www.hancom.co.kr/hwpml/2011/paragraph';
const HC_NS = 'http://www.hancom.co.kr/hwpml/2011/core';
const OPF_NS = 'http://www.idpf.org/2007/opf';
const MAIN_CONTENT_PATH = 'Contents/content.hpf';
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>';

type SectionContext = {
  fileName: string;
  document: Document;
  root: Element;
};

type ParagraphBlockRef = {
  type: 'paragraph';
  paragraphElement: Element;
  editable: boolean;
  block: ParagraphBlock;
};

type TableCellRef = {
  rowIndex: number;
  columnIndex: number;
  rowSpan: number;
  colSpan: number;
  cellElement: Element;
  subListElement: Element;
  paragraphs: Element[];
  block: TableCellBlock;
};

type TableBlockRef = {
  type: 'table';
  anchorParagraphElement: Element;
  tableElement: Element;
  editable: boolean;
  block: TableBlock;
  rows: Element[];
  cells: TableCellRef[];
  hasMergedCells: boolean;
};

type BlockRef = ParagraphBlockRef | TableBlockRef;

type SupplementalTextPartContext = {
  kind: SupplementalTextKind;
  fileName: string;
  document: Document;
  root: Element;
};

type SupplementalParagraphRef = {
  kind: SupplementalTextKind;
  paragraphIndex: number;
  paragraphElement: Element;
  editable: boolean;
  block: SupplementalTextBlock;
};

type ManifestItemRef = {
  id: string;
  href: string;
  resolvedPath: string;
  mediaType: string | null;
};

export type ParsedStructuredHwpx = {
  zip: JSZip;
  sections: SectionContext[];
  supplementalParts: SupplementalTextPartContext[];
  supplementalParagraphRefs: SupplementalParagraphRef[];
  blockRefs: BlockRef[][];
  body: NormalizedDocumentBody;
  supplementalText: SupplementalTextBlock[];
  imageReferences: EmbeddedImageReference[];
  paragraphs: string[];
  rawText: string;
  headerTexts: string[];
  footerTexts: string[];
  features: DocumentFeatureSummary;
  warnings: ApiWarning[];
  packageValidation: HwpxPackageValidationSummary;
};

function localNameOf(node: Node | null): string {
  if (!node) return '';
  if ('localName' in node && typeof node.localName === 'string' && node.localName.length > 0) {
    return node.localName;
  }
  if ('nodeName' in node && typeof node.nodeName === 'string') {
    const parts = node.nodeName.split(':');
    return parts[parts.length - 1] ?? node.nodeName;
  }
  return '';
}

function elementChildren(parent: Element): Element[] {
  const children: Element[] = [];
  for (let child = parent.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === child.ELEMENT_NODE) {
      children.push(child as Element);
    }
  }
  return children;
}

function descendantElements(parent: Element, localName?: string): Element[] {
  const items: Element[] = [];
  const visit = (node: Node): void => {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== child.ELEMENT_NODE) continue;
      const element = child as Element;
      if (!localName || localNameOf(element) === localName) {
        items.push(element);
      }
      visit(element);
    }
  };
  visit(parent);
  return items;
}

function parseInteger(value: string | null, fallback: number): number {
  if (typeof value !== 'string' || value.trim().length === 0) return fallback;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function parseNullableInteger(value: string | null): number | null {
  const parsed = parseInteger(value, NaN);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeWhitespace(text: string): string {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

function splitParagraphText(text: string): string[] {
  return normalizeWhitespace(text)
    .split(/\n\s*\n/g)
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);
}

function normalizePackagePath(value: string): string {
  const normalized = pathPosix.normalize(value.replace(/\\/g, '/'));
  return normalized.replace(/^\/+/, '').replace(/^\.\//, '');
}

function resolveManifestHref(zip: JSZip, manifestPath: string, href: string): string {
  if (href.trim().length === 0) {
    return '';
  }
  const direct = normalizePackagePath(href);
  if (zip.file(direct)) {
    return direct;
  }
  if (href.startsWith('/')) {
    return direct;
  }
  const relative = normalizePackagePath(pathPosix.join(pathPosix.dirname(manifestPath), href));
  if (zip.file(relative)) {
    return relative;
  }
  return direct;
}

async function parseManifestItems(zip: JSZip, manifestPath: string): Promise<ManifestItemRef[]> {
  const file = zip.file(manifestPath);
  if (!file) {
    return [];
  }

  const xml = await file.async('string');
  const validated = XMLValidator.validate(xml);
  if (validated !== true) {
    return [];
  }

  const parser = new DOMParser();
  const document = parser.parseFromString(xml, 'application/xml');
  const root = document.documentElement;
  if (!root) {
    return [];
  }

  return descendantElements(root, 'item')
    .map((element) => {
      const id = element.getAttribute('id')?.trim() ?? '';
      const href = element.getAttribute('href')?.trim() ?? '';
      if (!id || !href) {
        return null;
      }
      return {
        id,
        href,
        resolvedPath: resolveManifestHref(zip, manifestPath, href),
        mediaType: element.getAttribute('media-type'),
      } satisfies ManifestItemRef;
    })
    .filter((entry): entry is ManifestItemRef => entry !== null);
}

function buildImageAssetLookup(
  assetEntries: string[],
  manifestItems: ManifestItemRef[]
): {
  byItemId: Map<string, ManifestItemRef>;
  byAssetPath: Map<string, ManifestItemRef>;
  byFallbackId: Map<string, ManifestItemRef>;
} {
  const byItemId = new Map<string, ManifestItemRef>();
  const byAssetPath = new Map<string, ManifestItemRef>();
  const byFallbackId = new Map<string, ManifestItemRef>();

  manifestItems.forEach((item) => {
    if (!assetEntries.includes(item.resolvedPath)) {
      return;
    }
    byItemId.set(item.id, item);
    byAssetPath.set(item.resolvedPath, item);

    const baseName = pathPosix.basename(item.resolvedPath);
    const stem = baseName.includes('.') ? baseName.slice(0, baseName.lastIndexOf('.')) : baseName;
    if (!byFallbackId.has(baseName)) {
      byFallbackId.set(baseName, item);
    }
    if (!byFallbackId.has(stem)) {
      byFallbackId.set(stem, item);
    }
  });

  assetEntries.forEach((assetPath) => {
    if (byAssetPath.has(assetPath)) {
      return;
    }
    const baseName = pathPosix.basename(assetPath);
    const stem = baseName.includes('.') ? baseName.slice(0, baseName.lastIndexOf('.')) : baseName;
    const synthetic: ManifestItemRef = {
      id: stem,
      href: assetPath,
      resolvedPath: assetPath,
      mediaType: null,
    };
    byAssetPath.set(assetPath, synthetic);
    if (!byFallbackId.has(baseName)) {
      byFallbackId.set(baseName, synthetic);
    }
    if (!byFallbackId.has(stem)) {
      byFallbackId.set(stem, synthetic);
    }
  });

  return {
    byItemId,
    byAssetPath,
    byFallbackId,
  };
}

const EMPTY_IMAGE_ASSET_LOOKUP = buildImageAssetLookup([], []);

const IMAGE_MEDIA_TYPES = new Map<string, string>([
  ['.png', 'image/png'],
  ['.jpg', 'image/jpeg'],
  ['.jpeg', 'image/jpeg'],
  ['.gif', 'image/gif'],
  ['.bmp', 'image/bmp'],
  ['.svg', 'image/svg+xml'],
  ['.webp', 'image/webp'],
  ['.tif', 'image/tiff'],
  ['.tiff', 'image/tiff'],
]);

const DEFAULT_INSERTED_IMAGE_WIDTH = 7200;
const DEFAULT_INSERTED_IMAGE_HEIGHT = 5400;

function extractTextFromElement(element: Element, options?: { skipTables?: boolean }): string {
  const fragments: string[] = [];
  const skipTables = options?.skipTables ?? false;

  const walk = (node: Node): void => {
    if (node.nodeType === node.TEXT_NODE) {
      fragments.push(node.nodeValue ?? '');
      return;
    }

    if (node.nodeType !== node.ELEMENT_NODE) return;
    const el = node as Element;
    const name = localNameOf(el);
    if (skipTables && name === 'tbl') {
      return;
    }
    if (name === 'lineBreak') {
      fragments.push('\n');
      return;
    }
    if (name === 'tab') {
      fragments.push('\t');
      return;
    }
    for (let child = el.firstChild; child; child = child.nextSibling) {
      walk(child);
    }
  };

  walk(element);
  return normalizeWhitespace(fragments.join(''));
}

function directRunChildren(paragraphElement: Element): Element[] {
  return elementChildren(paragraphElement).filter((child) => localNameOf(child) === 'run');
}

function paragraphHasUnsupportedObjects(paragraphElement: Element): boolean {
  const unsupported = new Set([
    'tbl',
    'pic',
    'picture',
    'container',
    'ole',
    'equation',
    'line',
    'rect',
    'ellipse',
    'arc',
    'polygon',
    'curve',
    'connectLine',
    'textart',
    'video',
    'chart',
    'hiddenComment',
    'footNote',
    'endNote',
  ]);

  return descendantElements(paragraphElement).some((element) => unsupported.has(localNameOf(element)));
}

function findFirstTableInParagraph(paragraphElement: Element): Element | null {
  for (const run of directRunChildren(paragraphElement)) {
    const table = elementChildren(run).find((child) => localNameOf(child) === 'tbl');
    if (table) {
      return table;
    }
  }
  return null;
}

function paragraphCanBeRepresentedAsTable(paragraphElement: Element, tableElement: Element): boolean {
  const paragraphText = extractTextFromElement(paragraphElement, { skipTables: true });
  if (paragraphText.length > 0) return false;

  for (const run of directRunChildren(paragraphElement)) {
    for (const child of elementChildren(run)) {
      const name = localNameOf(child);
      if (!['secPr', 'ctrl', 'tbl', 't'].includes(name)) {
        return false;
      }
      if (name === 't' && normalizeWhitespace(child.textContent ?? '').length > 0) {
        return false;
      }
      if (name === 'tbl' && child !== tableElement) {
        return false;
      }
    }
  }

  return true;
}

function findDirectDescendantByName(parent: Element, localName: string): Element | null {
  return descendantElements(parent, localName)[0] ?? null;
}

function parseBoxEdges(
  element: Element | null,
  attributeNames: { left: string; right: string; top: string; bottom: string }
): { left: number | null; right: number | null; top: number | null; bottom: number | null } {
  return {
    left: element ? parseNullableInteger(element.getAttribute(attributeNames.left)) : null,
    right: element ? parseNullableInteger(element.getAttribute(attributeNames.right)) : null,
    top: element ? parseNullableInteger(element.getAttribute(attributeNames.top)) : null,
    bottom: element ? parseNullableInteger(element.getAttribute(attributeNames.bottom)) : null,
  };
}

function parsePicturePlacement(pictureElement: Element) {
  const size = findDirectDescendantByName(pictureElement, 'sz');
  const inMargin = findDirectDescendantByName(pictureElement, 'inMargin');
  const imgClip = findDirectDescendantByName(pictureElement, 'imgClip');

  return {
    objectId: pictureElement.getAttribute('id')?.trim() ?? null,
    instanceId: pictureElement.getAttribute('instid')?.trim() ?? null,
    zOrder: parseNullableInteger(pictureElement.getAttribute('zOrder')),
    textWrap: pictureElement.getAttribute('textWrap')?.trim() ?? null,
    textFlow: pictureElement.getAttribute('textFlow')?.trim() ?? null,
    numberingType: pictureElement.getAttribute('numberingType')?.trim() ?? null,
    widthRelTo: size?.getAttribute('widthRelTo')?.trim() ?? null,
    heightRelTo: size?.getAttribute('heightRelTo')?.trim() ?? null,
    margins: parseBoxEdges(inMargin, { left: 'left', right: 'right', top: 'top', bottom: 'bottom' }),
    clip: parseBoxEdges(imgClip, { left: 'left', right: 'right', top: 'top', bottom: 'bottom' }),
  };
}

function probePngSize(bytes: Buffer): { width: number; height: number } | null {
  if (bytes.length < 24 || bytes.toString('ascii', 1, 4) !== 'PNG') {
    return null;
  }
  return {
    width: bytes.readUInt32BE(16),
    height: bytes.readUInt32BE(20),
  };
}

function probeGifSize(bytes: Buffer): { width: number; height: number } | null {
  if (bytes.length < 10 || (bytes.toString('ascii', 0, 6) !== 'GIF87a' && bytes.toString('ascii', 0, 6) !== 'GIF89a')) {
    return null;
  }
  return {
    width: bytes.readUInt16LE(6),
    height: bytes.readUInt16LE(8),
  };
}

function probeBmpSize(bytes: Buffer): { width: number; height: number } | null {
  if (bytes.length < 26 || bytes.toString('ascii', 0, 2) !== 'BM') {
    return null;
  }
  return {
    width: Math.abs(bytes.readInt32LE(18)),
    height: Math.abs(bytes.readInt32LE(22)),
  };
}

function probeJpegSize(bytes: Buffer): { width: number; height: number } | null {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) {
    return null;
  }

  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    const marker = bytes[offset + 1];
    if (marker === 0xd9 || marker === 0xda) {
      break;
    }
    const segmentLength = bytes.readUInt16BE(offset + 2);
    if (segmentLength < 2 || offset + 2 + segmentLength > bytes.length) {
      break;
    }
    if (
      (marker >= 0xc0 && marker <= 0xc3)
      || (marker >= 0xc5 && marker <= 0xc7)
      || (marker >= 0xc9 && marker <= 0xcb)
      || (marker >= 0xcd && marker <= 0xcf)
    ) {
      return {
        height: bytes.readUInt16BE(offset + 5),
        width: bytes.readUInt16BE(offset + 7),
      };
    }
    offset += 2 + segmentLength;
  }

  return null;
}

function probeWebpSize(bytes: Buffer): { width: number; height: number } | null {
  if (bytes.length < 30 || bytes.toString('ascii', 0, 4) !== 'RIFF' || bytes.toString('ascii', 8, 12) !== 'WEBP') {
    return null;
  }

  const chunkType = bytes.toString('ascii', 12, 16);
  if (chunkType === 'VP8X' && bytes.length >= 30) {
    return {
      width: 1 + bytes.readUIntLE(24, 3),
      height: 1 + bytes.readUIntLE(27, 3),
    };
  }
  if (chunkType === 'VP8 ' && bytes.length >= 30) {
    return {
      width: bytes.readUInt16LE(26) & 0x3fff,
      height: bytes.readUInt16LE(28) & 0x3fff,
    };
  }
  if (chunkType === 'VP8L' && bytes.length >= 25) {
    const value = bytes.readUInt32LE(21);
    return {
      width: (value & 0x3fff) + 1,
      height: ((value >> 14) & 0x3fff) + 1,
    };
  }

  return null;
}

function probeSvgSize(bytes: Buffer): { width: number; height: number } | null {
  const text = bytes.toString('utf8');
  if (!text.includes('<svg')) {
    return null;
  }

  const widthMatch = text.match(/\bwidth=["']([0-9.]+)(px)?["']/i);
  const heightMatch = text.match(/\bheight=["']([0-9.]+)(px)?["']/i);
  if (widthMatch && heightMatch) {
    return {
      width: Math.max(1, Math.round(Number.parseFloat(widthMatch[1] ?? '0'))),
      height: Math.max(1, Math.round(Number.parseFloat(heightMatch[1] ?? '0'))),
    };
  }

  const viewBoxMatch = text.match(/\bviewBox=["'][^"']*?([0-9.]+)\s+([0-9.]+)["']/i);
  if (viewBoxMatch) {
    return {
      width: Math.max(1, Math.round(Number.parseFloat(viewBoxMatch[1] ?? '0'))),
      height: Math.max(1, Math.round(Number.parseFloat(viewBoxMatch[2] ?? '0'))),
    };
  }

  return null;
}

function probeIntrinsicImageSize(imagePath: string, bytes: Buffer): { width: number; height: number } | null {
  const extension = extname(imagePath).toLowerCase();
  if (extension === '.png') return probePngSize(bytes);
  if (extension === '.gif') return probeGifSize(bytes);
  if (extension === '.jpg' || extension === '.jpeg') return probeJpegSize(bytes);
  if (extension === '.bmp') return probeBmpSize(bytes);
  if (extension === '.webp') return probeWebpSize(bytes);
  if (extension === '.svg') return probeSvgSize(bytes);
  return null;
}

function inferInsertedPictureSize(
  imagePath: string,
  bytes: Buffer,
  patch?: Extract<TextMutationOperation, { op: 'insert_image_from_prototype' }>['patch']
): { width: number; height: number } {
  const intrinsic = probeIntrinsicImageSize(imagePath, bytes);
  const ratio = intrinsic && intrinsic.width > 0 && intrinsic.height > 0
    ? intrinsic.width / intrinsic.height
    : 4 / 3;

  if (typeof patch?.width === 'number' && typeof patch?.height === 'number') {
    return { width: patch.width, height: patch.height };
  }
  if (typeof patch?.width === 'number') {
    return {
      width: patch.width,
      height: Math.max(1, Math.round(patch.width / ratio)),
    };
  }
  if (typeof patch?.height === 'number') {
    return {
      width: Math.max(1, Math.round(patch.height * ratio)),
      height: patch.height,
    };
  }
  if (intrinsic) {
    return {
      width: DEFAULT_INSERTED_IMAGE_WIDTH,
      height: Math.max(1, Math.round(DEFAULT_INSERTED_IMAGE_WIDTH / ratio)),
    };
  }
  return {
    width: DEFAULT_INSERTED_IMAGE_WIDTH,
    height: DEFAULT_INSERTED_IMAGE_HEIGHT,
  };
}

function ensureDirectChildElement(parent: Element, namespaceUri: string, localName: string, prefix: string): Element {
  const existing = findDirectChild(parent, localName);
  if (existing) {
    return existing;
  }
  const created = parent.ownerDocument.createElementNS(namespaceUri, `${prefix}:${localName}`);
  parent.appendChild(created);
  return created;
}

function setOptionalNumericAttribute(element: Element, attributeName: string, value: number | undefined): void {
  if (typeof value !== 'number') {
    return;
  }
  element.setAttribute(attributeName, String(value));
}

function maybeUpdatePictureClipForResizedImage(
  pictureElement: Element,
  previousWidth: number | null,
  previousHeight: number | null,
  nextWidth: number | null,
  nextHeight: number | null,
  explicitClipPatch: { left?: number; right?: number; top?: number; bottom?: number } | undefined
): void {
  if (!nextWidth || !nextHeight || explicitClipPatch) {
    return;
  }
  const clip = findDirectChild(pictureElement, 'imgClip');
  if (!clip || !previousWidth || !previousHeight) {
    return;
  }

  const left = parseNullableInteger(clip.getAttribute('left'));
  const right = parseNullableInteger(clip.getAttribute('right'));
  const top = parseNullableInteger(clip.getAttribute('top'));
  const bottom = parseNullableInteger(clip.getAttribute('bottom'));

  const coversFullImage = left === 0 && top === 0 && right === previousWidth && bottom === previousHeight;
  if (!coversFullImage) {
    return;
  }

  clip.setAttribute('right', String(nextWidth));
  clip.setAttribute('bottom', String(nextHeight));
}

function syncPictureRectangle(pictureElement: Element, width: number | null, height: number | null): void {
  if (!width || !height) {
    return;
  }
  const imgRect = findDirectChild(pictureElement, 'imgRect');
  if (!imgRect) {
    return;
  }

  const pt0 = ensureDirectChildElement(imgRect, HC_NS, 'pt0', 'hc');
  const pt1 = ensureDirectChildElement(imgRect, HC_NS, 'pt1', 'hc');
  const pt2 = ensureDirectChildElement(imgRect, HC_NS, 'pt2', 'hc');
  const pt3 = ensureDirectChildElement(imgRect, HC_NS, 'pt3', 'hc');

  pt0.setAttribute('x', '0');
  pt0.setAttribute('y', '0');
  pt1.setAttribute('x', String(width));
  pt1.setAttribute('y', '0');
  pt2.setAttribute('x', String(width));
  pt2.setAttribute('y', String(height));
  pt3.setAttribute('x', '0');
  pt3.setAttribute('y', String(height));
}

function extractImageReferencesFromParagraph(
  paragraphElement: Element,
  target: {
    kind: EmbeddedImageKind;
    targetId: string;
    containerId: string;
    sectionIndex: number;
    blockIndex: number;
    paragraphIndex?: number;
    rowIndex?: number;
    columnIndex?: number;
  },
  assetLookup: ReturnType<typeof buildImageAssetLookup>
): EmbeddedImageReference[] {
  const references: EmbeddedImageReference[] = [];

  descendantElements(paragraphElement, 'pic').forEach((pictureElement) => {
    const imageElement = findDirectDescendantByName(pictureElement, 'img');
    const binaryItemId = imageElement?.getAttribute('binaryItemIDRef')?.trim() ?? '';
    if (!binaryItemId) {
      return;
    }

    const manifestItem = assetLookup.byItemId.get(binaryItemId) ?? assetLookup.byFallbackId.get(binaryItemId) ?? null;
    const imageDim = findDirectDescendantByName(pictureElement, 'imgDim');
    const size = findDirectDescendantByName(pictureElement, 'sz');
    const parsedWidth = parseInteger(imageDim?.getAttribute('dimwidth') ?? size?.getAttribute('width') ?? null, NaN);
    const parsedHeight = parseInteger(imageDim?.getAttribute('dimheight') ?? size?.getAttribute('height') ?? null, NaN);

    references.push({
      kind: target.kind,
      targetId: target.targetId,
      containerId: target.containerId,
      sectionIndex: target.sectionIndex,
      blockIndex: target.blockIndex,
      paragraphIndex: target.paragraphIndex,
      rowIndex: target.rowIndex,
      columnIndex: target.columnIndex,
      binaryItemId,
      assetPath: manifestItem?.resolvedPath ?? null,
      assetFileName: manifestItem ? pathPosix.basename(manifestItem.resolvedPath) : null,
      mediaType: manifestItem?.mediaType ?? null,
      width: Number.isFinite(parsedWidth) ? parsedWidth : null,
      height: Number.isFinite(parsedHeight) ? parsedHeight : null,
      placement: parsePicturePlacement(pictureElement),
    });
  });

  return references;
}

function createParagraphBlock(blockId: string, paragraphElement: Element): ParagraphBlock {
  const text = extractTextFromElement(paragraphElement, { skipTables: true });
  const containsObjects = paragraphHasUnsupportedObjects(paragraphElement);
  return {
    type: 'paragraph',
    blockId,
    text,
    styleRef: paragraphElement.getAttribute('styleIDRef'),
    editable: !containsObjects,
    containsObjects,
  };
}

function findDirectChild(parent: Element, localName: string): Element | null {
  for (const child of elementChildren(parent)) {
    if (localNameOf(child) === localName) {
      return child;
    }
  }
  return null;
}

function createTableBlock(
  blockId: string,
  tableElement: Element,
  physicalRows: Element[],
  assetLookup: ReturnType<typeof buildImageAssetLookup>,
  sectionIndex: number,
  blockIndex: number
): { block: TableBlock; cellRefs: TableCellRef[]; hasMergedCells: boolean; imageReferences: EmbeddedImageReference[] } {
  const cellRefs: TableCellRef[] = [];
  const imageReferences: EmbeddedImageReference[] = [];
  let hasMergedCells = false;

  physicalRows.forEach((rowElement, physicalRowIndex) => {
    const cells = elementChildren(rowElement).filter((child) => localNameOf(child) === 'tc');
    cells.forEach((cellElement, physicalCellIndex) => {
      const cellAddr = findDirectChild(cellElement, 'cellAddr');
      const cellSpan = findDirectChild(cellElement, 'cellSpan');
      const subList = findDirectChild(cellElement, 'subList');
      if (!subList) {
        return;
      }

      const rowIndex = parseInteger(cellAddr?.getAttribute('rowAddr') ?? null, physicalRowIndex);
      const columnIndex = parseInteger(cellAddr?.getAttribute('colAddr') ?? null, physicalCellIndex);
      const rowSpan = parseInteger(cellSpan?.getAttribute('rowSpan') ?? null, 1);
      const colSpan = parseInteger(cellSpan?.getAttribute('colSpan') ?? null, 1);
      if (rowSpan > 1 || colSpan > 1) {
        hasMergedCells = true;
      }

      const paragraphs = elementChildren(subList).filter((child) => localNameOf(child) === 'p');
      const paragraphTexts = paragraphs
        .map((paragraph) => extractTextFromElement(paragraph, { skipTables: true }))
        .filter((text) => text.length > 0);

      paragraphs.forEach((paragraph, paragraphIndex) => {
        imageReferences.push(...extractImageReferencesFromParagraph(
          paragraph,
          {
            kind: 'table_cell',
            targetId: `${blockId}:r${rowIndex}:c${columnIndex}`,
            containerId: blockId,
            sectionIndex,
            blockIndex,
            rowIndex,
            columnIndex,
            paragraphIndex,
          },
          assetLookup
        ));
      });

      const block: TableCellBlock = {
        cellId: `${blockId}:r${rowIndex}:c${columnIndex}`,
        rowIndex,
        columnIndex,
        rowSpan,
        colSpan,
        text: paragraphTexts.join('\n'),
        paragraphs: paragraphTexts,
        editable: true,
      };

      cellRefs.push({
        rowIndex,
        columnIndex,
        rowSpan,
        colSpan,
        cellElement,
        subListElement: subList,
        paragraphs,
        block,
      });
    });
  });

  const block: TableBlock = {
    type: 'table',
    blockId,
    tableId: tableElement.getAttribute('id'),
    rowCount: parseInteger(tableElement.getAttribute('rowCnt'), physicalRows.length),
    columnCount: parseInteger(tableElement.getAttribute('colCnt'), 0),
    editable: true,
    cells: cellRefs.map((cellRef) => cellRef.block),
  };

  return { block, cellRefs, hasMergedCells, imageReferences };
}

function createEmptyParagraph(document: Document, templateParagraph?: Element | null): Element {
  const paragraph = templateParagraph
    ? (templateParagraph.cloneNode(false) as Element)
    : document.createElementNS(HP_NS, 'hp:p');

  if (!paragraph.getAttribute('id')) {
    paragraph.setAttribute('id', '0');
  }
  if (!paragraph.getAttribute('paraPrIDRef')) {
    paragraph.setAttribute('paraPrIDRef', templateParagraph?.getAttribute('paraPrIDRef') || '0');
  }
  if (!paragraph.getAttribute('styleIDRef')) {
    paragraph.setAttribute('styleIDRef', templateParagraph?.getAttribute('styleIDRef') || '0');
  }
  paragraph.setAttribute('pageBreak', paragraph.getAttribute('pageBreak') || '0');
  paragraph.setAttribute('columnBreak', paragraph.getAttribute('columnBreak') || '0');
  paragraph.setAttribute('merged', paragraph.getAttribute('merged') || '0');

  return paragraph;
}

function stripParagraphTextContent(paragraphElement: Element): { charPrIDRef: string } {
  let charPrIDRef = '0';
  const runs = elementChildren(paragraphElement).filter((child) => localNameOf(child) === 'run');

  for (const run of runs) {
    if (run.getAttribute('charPrIDRef')) {
      charPrIDRef = run.getAttribute('charPrIDRef') || charPrIDRef;
    }

    for (let child = run.firstChild; child;) {
      const next = child.nextSibling;
      if (child.nodeType === child.ELEMENT_NODE) {
        const name = localNameOf(child as Element);
        if (name === 't') {
          run.removeChild(child);
        }
      } else if (child.nodeType === child.TEXT_NODE) {
        run.removeChild(child);
      }
      child = next;
    }
  }

  for (let child = paragraphElement.firstChild; child;) {
    const next = child.nextSibling;
    if (child.nodeType === child.ELEMENT_NODE && localNameOf(child as Element) === 'linesegarray') {
      paragraphElement.removeChild(child);
    }
    child = next;
  }

  return { charPrIDRef };
}

function appendPlainTextRun(paragraphElement: Element, text: string, charPrIDRef: string): void {
  const document = paragraphElement.ownerDocument;
  const run = document.createElementNS(HP_NS, 'hp:run');
  run.setAttribute('charPrIDRef', charPrIDRef);
  const t = document.createElementNS(HP_NS, 'hp:t');
  t.appendChild(document.createTextNode(text));
  run.appendChild(t);
  paragraphElement.appendChild(run);
}

function setParagraphPlainText(paragraphElement: Element, text: string): void {
  const { charPrIDRef } = stripParagraphTextContent(paragraphElement);
  appendPlainTextRun(paragraphElement, text, charPrIDRef);
}

function ensureSectionHasParagraph(sectionRoot: Element): void {
  const paragraphs = elementChildren(sectionRoot).filter((child) => localNameOf(child) === 'p');
  if (paragraphs.length > 0) return;
  const paragraph = createEmptyParagraph(sectionRoot.ownerDocument);
  setParagraphPlainText(paragraph, '');
  sectionRoot.appendChild(paragraph);
}

function getAnchorParagraph(ref: BlockRef): Element {
  return ref.type === 'paragraph' ? ref.paragraphElement : ref.anchorParagraphElement;
}

function createParagraphFromTemplate(afterParagraph: Element, text: string): Element {
  const paragraph = createEmptyParagraph(afterParagraph.ownerDocument, afterParagraph);
  const clonedChildren = elementChildren(afterParagraph)
    .filter((child) => localNameOf(child) === 'run')
    .map((child) => child.cloneNode(true) as Element);

  for (const child of clonedChildren) {
    paragraph.appendChild(child);
  }

  const { charPrIDRef } = stripParagraphTextContent(paragraph);
  appendPlainTextRun(paragraph, text, charPrIDRef);
  return paragraph;
}

function createCellParagraphFromTemplate(cellRef: TableCellRef): Element {
  const template = cellRef.paragraphs[0] ?? null;
  const paragraph = createEmptyParagraph(cellRef.subListElement.ownerDocument, template);
  if (template) {
    const runs = elementChildren(template)
      .filter((child) => localNameOf(child) === 'run')
      .map((child) => child.cloneNode(true) as Element);
    for (const run of runs) {
      paragraph.appendChild(run);
    }
  }
  return paragraph;
}

function setCellText(cellRef: TableCellRef, text: string): void {
  const values = splitParagraphText(text);
  const paragraphs = values.length > 0 ? values : [''];
  const existing = [...cellRef.paragraphs];

  for (let index = 0; index < paragraphs.length; index += 1) {
    const paragraph = existing[index] ?? createCellParagraphFromTemplate(cellRef);
    if (!existing[index]) {
      cellRef.subListElement.appendChild(paragraph);
      cellRef.paragraphs.push(paragraph);
    }
    setParagraphPlainText(paragraph, paragraphs[index] ?? '');
  }

  for (let index = cellRef.paragraphs.length - 1; index >= paragraphs.length; index -= 1) {
    const paragraph = cellRef.paragraphs[index];
    cellRef.subListElement.removeChild(paragraph);
    cellRef.paragraphs.splice(index, 1);
  }
}

function ensureCellAddrElement(cell: Element): Element {
  let cellAddr = findDirectChild(cell, 'cellAddr');
  if (!cellAddr) {
    cellAddr = cell.ownerDocument.createElementNS(HP_NS, 'hp:cellAddr');
    cell.appendChild(cellAddr);
  }
  return cellAddr;
}

function ensureCellSpanElement(cell: Element): Element {
  let cellSpan = findDirectChild(cell, 'cellSpan');
  if (!cellSpan) {
    cellSpan = cell.ownerDocument.createElementNS(HP_NS, 'hp:cellSpan');
    cell.appendChild(cellSpan);
  }
  return cellSpan;
}

function rowCells(row: Element): Element[] {
  return elementChildren(row).filter((child) => localNameOf(child) === 'tc');
}

function setRowCount(tableRef: TableBlockRef): void {
  tableRef.rows = elementChildren(tableRef.tableElement).filter((child) => localNameOf(child) === 'tr');
  tableRef.tableElement.setAttribute('rowCnt', String(tableRef.rows.length));
}

function cloneTableRowPreserveSpans(templateRow: Element, insertedRowIndex: number): Element {
  const clonedRow = templateRow.cloneNode(true) as Element;

  rowCells(clonedRow).forEach((cell) => {
    const cellAddr = ensureCellAddrElement(cell);
    cellAddr.setAttribute('rowAddr', String(insertedRowIndex));

    const cellSpan = ensureCellSpanElement(cell);
    if (!cellSpan.getAttribute('rowSpan')) {
      cellSpan.setAttribute('rowSpan', '1');
    }
    if (!cellSpan.getAttribute('colSpan')) {
      cellSpan.setAttribute('colSpan', '1');
    }

    const subList = findDirectChild(cell, 'subList');
    if (!subList) return;
    const paragraphs = elementChildren(subList).filter((child) => localNameOf(child) === 'p');
    paragraphs.forEach((paragraph, index) => {
      if (index === 0) {
        setParagraphPlainText(paragraph, '');
      } else {
        subList.removeChild(paragraph);
      }
    });
  });

  return clonedRow;
}

function shiftDownstreamRowAddresses(tableRef: TableBlockRef, startRowIndex: number, delta: number): void {
  tableRef.rows.slice(startRowIndex).forEach((row) => {
    rowCells(row).forEach((cell) => {
      const cellAddr = ensureCellAddrElement(cell);
      const current = parseInteger(cellAddr.getAttribute('rowAddr'), 0);
      cellAddr.setAttribute('rowAddr', String(current + delta));
    });
  });
}

function collectBoundaryCrossingCells(tableRef: TableBlockRef, splitRowIndex: number): TableCellRef[] {
  return tableRef.cells
    .filter((cell) => cell.rowIndex < splitRowIndex && cell.rowIndex + cell.rowSpan > splitRowIndex)
    .sort((left, right) => left.columnIndex - right.columnIndex);
}

function cellRefKey(cell: TableCellRef): string {
  return `${cell.rowIndex}:${cell.columnIndex}`;
}

function insertCellIntoRowByColumn(rowElement: Element, cellElement: Element, columnIndex: number): void {
  const cells = rowCells(rowElement);
  const nextCell = cells.find((cell) => {
    const cellAddr = findDirectChild(cell, 'cellAddr');
    const currentColumn = parseInteger(cellAddr?.getAttribute('colAddr') ?? null, Number.MAX_SAFE_INTEGER);
    return currentColumn > columnIndex;
  });

  if (nextCell) {
    rowElement.insertBefore(cellElement, nextCell);
    return;
  }

  rowElement.appendChild(cellElement);
}

function setCellSpanAndAddress(cellElement: Element, rowIndex: number, columnIndex: number, rowSpan: number, colSpan: number): void {
  const cellAddr = ensureCellAddrElement(cellElement);
  cellAddr.setAttribute('rowAddr', String(rowIndex));
  cellAddr.setAttribute('colAddr', String(columnIndex));

  const cellSpan = ensureCellSpanElement(cellElement);
  cellSpan.setAttribute('rowSpan', String(rowSpan));
  cellSpan.setAttribute('colSpan', String(colSpan));
}

function extendCellRowSpan(cell: TableCellRef, delta: number): void {
  if (delta <= 0) return;
  setCellSpanAndAddress(cell.cellElement, cell.rowIndex, cell.columnIndex, cell.rowSpan + delta, cell.colSpan);
}

function splitBoundaryCrossingCells(tableRef: TableBlockRef, splitRowIndex: number): boolean {
  const targetRow = tableRef.rows[splitRowIndex];
  if (!targetRow) {
    throw new Error(`Cannot split merged cells at row ${splitRowIndex}: target row does not exist.`);
  }

  const crossings = collectBoundaryCrossingCells(tableRef, splitRowIndex);
  if (crossings.length === 0) {
    return false;
  }

  crossings.forEach((cell) => {
    const upperSpan = splitRowIndex - cell.rowIndex;
    const lowerSpan = cell.rowIndex + cell.rowSpan - splitRowIndex;
    if (upperSpan <= 0 || lowerSpan <= 0) {
      throw new Error(`Cannot split merged cell r${cell.rowIndex} c${cell.columnIndex} at boundary ${splitRowIndex}.`);
    }

    setCellSpanAndAddress(cell.cellElement, cell.rowIndex, cell.columnIndex, upperSpan, cell.colSpan);

    const lowerClone = cell.cellElement.cloneNode(true) as Element;
    setCellSpanAndAddress(lowerClone, splitRowIndex, cell.columnIndex, lowerSpan, cell.colSpan);
    insertCellIntoRowByColumn(targetRow, lowerClone, cell.columnIndex);
  });

  refreshTableBlockRef(tableRef);
  return true;
}

function analyzeRegionInsertionBoundary(tableRef: TableBlockRef, startRowIndex: number, endRowIndex: number): {
  safe: boolean;
  reason?: string;
  topCrossings: TableCellRef[];
  bottomCrossings: TableCellRef[];
} {
  const topCrossings = collectBoundaryCrossingCells(tableRef, startRowIndex);
  const bottomCrossings = collectBoundaryCrossingCells(tableRef, endRowIndex + 1);
  let reason: string | undefined;
  if (topCrossings.length > 0) {
    reason = startRowIndex === endRowIndex
      ? `Row ${startRowIndex} is covered by a merged cell anchored above the selected repeat row.`
      : `Rows ${startRowIndex}-${endRowIndex} are covered by a merged cell anchored above the selected repeat region.`;
  } else if (bottomCrossings.length > 0) {
    reason = startRowIndex === endRowIndex
      ? `Row ${startRowIndex} contains or touches a vertical merge that crosses the insertion boundary.`
      : `Rows ${startRowIndex}-${endRowIndex} contain or touch a vertical merge that crosses the insertion boundary.`;
  }

  return {
    safe: topCrossings.length === 0 && bottomCrossings.length === 0,
    reason,
    topCrossings,
    bottomCrossings,
  };
}

function prepareRegionInsertionBoundary(
  tableRef: TableBlockRef,
  startRowIndex: number,
  endRowIndex: number,
  boundaryPolicy: 'reject' | 'split_boundary_merges' = 'reject'
): { safe: boolean; reason?: string } {
  let boundary = analyzeRegionInsertionBoundary(tableRef, startRowIndex, endRowIndex);
  if (boundary.safe || boundaryPolicy !== 'split_boundary_merges') {
    return boundary;
  }

  const regionRowCount = endRowIndex - startRowIndex + 1;
  const bottomKeys = new Set(boundary.bottomCrossings.map(cellRefKey));
  const coveringCells = boundary.topCrossings.filter((cell) => bottomKeys.has(cellRefKey(cell)));

  if (coveringCells.length > 0) {
    coveringCells.forEach((cell) => {
      extendCellRowSpan(cell, regionRowCount);
    });
    refreshTableBlockRef(tableRef);
  }

  boundary = analyzeRegionInsertionBoundary(tableRef, startRowIndex, endRowIndex);
  const topOnly = boundary.topCrossings.filter((cell) => !boundary.bottomCrossings.some((other) => cellRefKey(other) === cellRefKey(cell)));
  if (topOnly.length > 0) {
    splitBoundaryCrossingCells(tableRef, startRowIndex);
  }

  boundary = analyzeRegionInsertionBoundary(tableRef, startRowIndex, endRowIndex);
  const bottomOnly = boundary.bottomCrossings.filter((cell) => !boundary.topCrossings.some((other) => cellRefKey(other) === cellRefKey(cell)));
  if (bottomOnly.length > 0) {
    splitBoundaryCrossingCells(tableRef, endRowIndex + 1);
  }

  boundary = analyzeRegionInsertionBoundary(tableRef, startRowIndex, endRowIndex);
  const finalBottomKeys = new Set(boundary.bottomCrossings.map(cellRefKey));
  const unsupportedTop = boundary.topCrossings.filter((cell) => !finalBottomKeys.has(cellRefKey(cell)));
  const finalTopKeys = new Set(boundary.topCrossings.map(cellRefKey));
  const unsupportedBottom = boundary.bottomCrossings.filter((cell) => !finalTopKeys.has(cellRefKey(cell)));

  if (unsupportedTop.length === 0 && unsupportedBottom.length === 0) {
    return { safe: true };
  }

  return boundary;
}

function prepareRowInsertionBoundary(
  tableRef: TableBlockRef,
  rowIndex: number,
  boundaryPolicy: 'reject' | 'split_boundary_merges' = 'reject'
): { safe: boolean; reason?: string } {
  return prepareRegionInsertionBoundary(tableRef, rowIndex, rowIndex, boundaryPolicy);
}

function refreshTableBlockRef(tableRef: TableBlockRef): void {
  const rows = elementChildren(tableRef.tableElement).filter((child) => localNameOf(child) === 'tr');
  const refreshed = createTableBlock(tableRef.block.blockId, tableRef.tableElement, rows, EMPTY_IMAGE_ASSET_LOOKUP, 0, 0);
  tableRef.rows = rows;
  tableRef.cells = refreshed.cellRefs;
  tableRef.hasMergedCells = refreshed.hasMergedCells;
  tableRef.block.tableId = refreshed.block.tableId;
  tableRef.block.rowCount = refreshed.block.rowCount;
  tableRef.block.columnCount = refreshed.block.columnCount;
  tableRef.block.editable = refreshed.block.editable;
  tableRef.block.cells = refreshed.block.cells;
}

function insertTableRow(
  tableRef: TableBlockRef,
  rowIndex: number,
  boundaryPolicy: 'reject' | 'split_boundary_merges' = 'reject'
): void {
  if (rowIndex < 0 || rowIndex >= tableRef.rows.length) {
    throw new Error(`rowIndex ${rowIndex} is out of range.`);
  }

  const boundary = prepareRowInsertionBoundary(tableRef, rowIndex, boundaryPolicy);
  if (!boundary.safe) {
    throw new Error(boundary.reason || 'insert_table_row is not safe for the selected merged-table row.');
  }

  const insertAt = rowIndex + 1;
  shiftDownstreamRowAddresses(tableRef, insertAt, 1);

  const templateRow = tableRef.rows[rowIndex];
  const clonedRow = cloneTableRowPreserveSpans(templateRow, insertAt);
  const referenceNode = templateRow.nextSibling;
  tableRef.tableElement.insertBefore(clonedRow, referenceNode);
  setRowCount(tableRef);
  refreshTableBlockRef(tableRef);
}

function cloneTableRegion(
  tableRef: TableBlockRef,
  templateStartRowIndex: number,
  templateEndRowIndex: number,
  insertAfterRowIndex: number,
  boundaryPolicy: 'reject' | 'split_boundary_merges' = 'reject'
): void {
  if (templateStartRowIndex < 0 || templateEndRowIndex >= tableRef.rows.length || templateStartRowIndex > templateEndRowIndex) {
    throw new Error(`Template row region ${templateStartRowIndex}-${templateEndRowIndex} is out of range.`);
  }
  if (insertAfterRowIndex < 0 || insertAfterRowIndex >= tableRef.rows.length) {
    throw new Error(`insertAfterRowIndex ${insertAfterRowIndex} is out of range.`);
  }

  const boundary = prepareRegionInsertionBoundary(tableRef, templateStartRowIndex, templateEndRowIndex, boundaryPolicy);
  if (!boundary.safe) {
    throw new Error(boundary.reason || 'clone_table_region is not safe for the selected merged-table region.');
  }

  const insertAt = insertAfterRowIndex + 1;
  const templateRows = tableRef.rows.slice(templateStartRowIndex, templateEndRowIndex + 1);
  const regionRowCount = templateRows.length;
  shiftDownstreamRowAddresses(tableRef, insertAt, regionRowCount);

  const referenceNode = tableRef.rows[insertAfterRowIndex]?.nextSibling ?? null;
  templateRows.forEach((templateRow, offset) => {
    const clonedRow = cloneTableRowPreserveSpans(templateRow, insertAt + offset);
    tableRef.tableElement.insertBefore(clonedRow, referenceNode);
  });

  setRowCount(tableRef);
  refreshTableBlockRef(tableRef);
}

function deleteTableRow(
  tableRef: TableBlockRef,
  rowIndex: number,
  boundaryPolicy: 'reject' | 'split_boundary_merges' = 'reject'
): void {
  if (tableRef.rows.length <= 1) {
    throw new Error('Cannot delete the last remaining row in a table.');
  }
  if (rowIndex < 0 || rowIndex >= tableRef.rows.length) {
    throw new Error(`rowIndex ${rowIndex} is out of range.`);
  }

  const boundary = prepareRowInsertionBoundary(tableRef, rowIndex, boundaryPolicy);
  if (!boundary.safe) {
    throw new Error(boundary.reason || 'delete_table_row is not safe for the selected merged-table row.');
  }

  tableRef.tableElement.removeChild(tableRef.rows[rowIndex]);
  setRowCount(tableRef);
  shiftDownstreamRowAddresses(tableRef, rowIndex, -1);
  refreshTableBlockRef(tableRef);
}

function serializeSection(section: SectionContext): string {
  const serializer = new XMLSerializer();
  const xml = serializer.serializeToString(section.document);
  const normalized = xml.startsWith('<?xml') ? xml : `${XML_DECLARATION}${xml}`;
  const validated = XMLValidator.validate(normalized);
  if (validated !== true) {
    throw new Error(`Generated section XML is invalid for ${section.fileName}.`);
  }
  return normalized;
}

function serializeSupplementalPart(part: SupplementalTextPartContext): string {
  const serializer = new XMLSerializer();
  const xml = serializer.serializeToString(part.document);
  const normalized = xml.startsWith('<?xml') ? xml : `${XML_DECLARATION}${xml}`;
  const validated = XMLValidator.validate(normalized);
  if (validated !== true) {
    throw new Error(`Generated supplemental XML is invalid for ${part.fileName}.`);
  }
  return normalized;
}

function serializeXmlDocument(document: Document, label: string): string {
  const serializer = new XMLSerializer();
  const xml = serializer.serializeToString(document);
  const normalized = xml.startsWith('<?xml') ? xml : `${XML_DECLARATION}${xml}`;
  const validated = XMLValidator.validate(normalized);
  if (validated !== true) {
    throw new Error(`Generated XML is invalid for ${label}.`);
  }
  return normalized;
}

async function loadZipAndSections(filePath: string): Promise<{ zip: JSZip; sections: SectionContext[]; validation: HwpxPackageValidationSummary }> {
  const { zip, validation } = await loadValidatedHwpxPackage(filePath);
  const sectionNames = validation.sectionEntries;

  const parser = new DOMParser();
  const sections: SectionContext[] = [];

  for (const fileName of sectionNames) {
    const file = zip.file(fileName);
    if (!file) continue;
    const xml = await file.async('string');
    const validated = XMLValidator.validate(xml);
    if (validated !== true) {
      throw new Error(`Invalid XML in section: ${fileName}`);
    }

    const document = parser.parseFromString(xml, 'application/xml');
    const root = document.documentElement;
    if (!root) {
      throw new Error(`Missing root element in section: ${fileName}`);
    }
    sections.push({ fileName, document, root });
  }

  return { zip, sections, validation };
}

async function loadOptionalTextPart(
  zip: JSZip,
  fileName: string,
  kind: SupplementalTextKind
): Promise<{
  context: SupplementalTextPartContext | null;
  paragraphRefs: SupplementalParagraphRef[];
  texts: string[];
}> {
  const file = zip.file(fileName);
  if (!file) {
    return {
      context: null,
      paragraphRefs: [],
      texts: [],
    };
  }

  const xml = await file.async('string');
  const validated = XMLValidator.validate(xml);
  if (validated !== true) {
    return {
      context: null,
      paragraphRefs: [],
      texts: [],
    };
  }

  const parser = new DOMParser();
  const document = parser.parseFromString(xml, 'application/xml');
  const root = document.documentElement;
  if (!root) {
    return {
      context: null,
      paragraphRefs: [],
      texts: [],
    };
  }

  const paragraphRefs = descendantElements(root, 'p').map((paragraphElement, paragraphIndex) => {
    const text = extractTextFromElement(paragraphElement, { skipTables: true }).trim();
    const containsObjects = paragraphHasUnsupportedObjects(paragraphElement) || findFirstTableInParagraph(paragraphElement) !== null;
    const targetId = `${kind}:${paragraphIndex}`;

    return {
      kind,
      paragraphIndex,
      paragraphElement,
      editable: !containsObjects,
      block: {
        kind,
        targetId,
        paragraphIndex,
        text,
        editable: !containsObjects,
        containsObjects,
      },
    };
  });

  return {
    context: {
      kind,
      fileName,
      document,
      root,
    },
    paragraphRefs,
    texts: paragraphRefs
      .map((paragraphRef) => paragraphRef.block.text)
      .filter((text) => text.length > 0),
  };
}

function summarizeFeatures(
  sections: SectionContext[],
  blockRefs: BlockRef[][],
  supplemental: {
    supplementalParts: SupplementalTextPartContext[];
    headerTexts: string[];
    footerTexts: string[];
    headerPartPresent: boolean;
    footerPartPresent: boolean;
    headerEditableCount: number;
    footerEditableCount: number;
  }
): DocumentFeatureSummary {
  const editable = new Set<string>(['paragraph']);
  const preservedReadOnly = new Set<string>();
  const unsupported = new Set<string>();

  if (blockRefs.some((section) => section.some((block) => block.type === 'table'))) {
    editable.add('table');
  }

  if (blockRefs.some((section) => section.some((block) => block.type === 'paragraph' && block.block.containsObjects))) {
    preservedReadOnly.add('embedded_object_runs');
  }
  if (supplemental.headerEditableCount > 0) {
    editable.add('header');
  }
  if (supplemental.headerPartPresent) {
    preservedReadOnly.add('header');
  }
  if (supplemental.footerEditableCount > 0) {
    editable.add('footer');
  }
  if (supplemental.footerPartPresent) {
    preservedReadOnly.add('footer');
  }

  const tagToFeature = new Map<string, string>([
    ['picture', 'image'],
    ['pic', 'image'],
    ['container', 'container'],
    ['ole', 'ole'],
    ['equation', 'equation'],
    ['hiddenComment', 'comment'],
    ['footNote', 'footnote'],
    ['endNote', 'endnote'],
    ['header', 'header'],
    ['footer', 'footer'],
    ['chart', 'chart'],
    ['video', 'video'],
    ['textart', 'textart'],
  ]);

  const rootsToScan = [
    ...sections.map((section) => section.root),
    ...supplemental.supplementalParts.map((part) => part.root),
  ];

  rootsToScan.forEach((root) => {
    descendantElements(root).forEach((element) => {
      const feature = tagToFeature.get(localNameOf(element));
      if (feature) unsupported.add(feature);
    });
  });

  if (unsupported.has('image')) {
    editable.add('image_asset');
  }

  const allowPlainTextSave = unsupported.size === 0
    && blockRefs.every((section) => section.every((block) => block.type === 'paragraph' && block.editable));

  return {
    projectionMode: 'structured_hwpx',
    authoritative: true,
    editable: [...editable],
    preservedReadOnly: [...preservedReadOnly],
    unsupported: [...unsupported],
    hasUnsupportedEditableFeatures: unsupported.size > 0 || preservedReadOnly.size > 0,
    allowPlainTextSave,
  };
}

function buildWarnings(features: DocumentFeatureSummary, packageValidation: HwpxPackageValidationSummary): ApiWarning[] {
  const warnings: ApiWarning[] = [];
  if (!features.allowPlainTextSave) {
    warnings.push({
      code: 'STRUCTURED_MUTATIONS_REQUIRED',
      message: 'This document contains structured content. Use block/table mutations instead of replacing the entire plain-text projection.',
    });
  }

  const partiallyEditable = new Set<string>();
  if (features.editable.includes('image_asset') && features.unsupported.includes('image')) {
    partiallyEditable.add('image');
  }

  if (features.unsupported.length > 0 || features.preservedReadOnly.length > 0) {
    const preserved = [...features.preservedReadOnly, ...features.unsupported]
      .filter((feature) => !partiallyEditable.has(feature));
    if (preserved.length > 0) {
      warnings.push({
        code: 'READ_ONLY_FEATURES_PRESENT',
        message: `This document includes read-only preserved features: ${preserved.join(', ')}.`,
      });
    }
  }
  if (partiallyEditable.has('image')) {
    warnings.push({
      code: 'PARTIAL_IMAGE_EDITING_ONLY',
      message: 'Embedded images support deterministic asset replacement and exact placement patches, but freeform anchor/layout editing remains read-only.',
    });
  }
  if (packageValidation.missingRecommendedEntries.length > 0) {
    warnings.push({
      code: 'PACKAGE_RECOMMENDED_PARTS_MISSING',
      message: `This HWPX package is missing recommended parts: ${packageValidation.missingRecommendedEntries.join(', ')}.`,
    });
  }
  return warnings;
}

export async function parseStructuredHwpx(filePath: string): Promise<ParsedStructuredHwpx> {
  const { zip, sections, validation } = await loadZipAndSections(filePath);
  const manifestItems = await parseManifestItems(zip, 'Contents/content.hpf');
  const assetLookup = buildImageAssetLookup(validation.assetEntries, manifestItems);
  const headerPartPresent = zip.file('Contents/header.xml') !== null;
  const footerPartPresent = zip.file('Contents/footer.xml') !== null;
  const headerPart = await loadOptionalTextPart(zip, 'Contents/header.xml', 'header');
  const footerPart = await loadOptionalTextPart(zip, 'Contents/footer.xml', 'footer');
  const headerTexts = headerPart.texts;
  const footerTexts = footerPart.texts;
  const supplementalParts = [headerPart.context, footerPart.context].filter((entry): entry is SupplementalTextPartContext => entry !== null);
  const supplementalParagraphRefs = [...headerPart.paragraphRefs, ...footerPart.paragraphRefs];
  const bodySections: NormalizedDocumentBody['sections'] = [];
  const blockRefs: BlockRef[][] = [];
  const imageReferences: EmbeddedImageReference[] = [];
  const plainTextSegments: string[] = [];

  sections.forEach((section, sectionIndex) => {
    const blocks: DocumentBlock[] = [];
    const refs: BlockRef[] = [];
    let blockCounter = 0;

    for (const child of elementChildren(section.root)) {
      if (localNameOf(child) !== 'p') {
        continue;
      }

      const blockId = `${sectionIndex}:${blockCounter}`;
      const firstTable = findFirstTableInParagraph(child);
      if (firstTable && paragraphCanBeRepresentedAsTable(child, firstTable)) {
        const rows = elementChildren(firstTable).filter((element) => localNameOf(element) === 'tr');
        const { block, cellRefs, hasMergedCells, imageReferences: tableImages } = createTableBlock(
          blockId,
          firstTable,
          rows,
          assetLookup,
          sectionIndex,
          blockCounter
        );
        blocks.push(block);
        imageReferences.push(...tableImages);
        refs.push({
          type: 'table',
          anchorParagraphElement: child,
          tableElement: firstTable,
          editable: true,
          block,
          rows,
          cells: cellRefs,
          hasMergedCells,
        });

        const orderedCells = [...block.cells].sort((left, right) => {
          if (left.rowIndex !== right.rowIndex) return left.rowIndex - right.rowIndex;
          return left.columnIndex - right.columnIndex;
        });
        const tableText = orderedCells.map((cell) => cell.text).filter((text) => text.length > 0).join('\n');
        if (tableText.length > 0) {
          plainTextSegments.push(tableText);
        }
        blockCounter += 1;
        continue;
      }

      const block = createParagraphBlock(blockId, child);
      blocks.push(block);
      imageReferences.push(...extractImageReferencesFromParagraph(
        child,
        {
          kind: 'paragraph',
          targetId: blockId,
          containerId: blockId,
          sectionIndex,
          blockIndex: blockCounter,
        },
        assetLookup
      ));
      refs.push({
        type: 'paragraph',
        paragraphElement: child,
        editable: block.editable,
        block,
      });
      if (block.text.length > 0) {
        plainTextSegments.push(block.text);
      }
      blockCounter += 1;
    }

    bodySections.push({
      sectionId: section.fileName,
      blocks,
    });
    blockRefs.push(refs);
  });

  supplementalParagraphRefs.forEach((entry) => {
    imageReferences.push(...extractImageReferencesFromParagraph(
      entry.paragraphElement,
      {
        kind: entry.kind === 'header' ? 'header_paragraph' : 'footer_paragraph',
        targetId: entry.block.targetId,
        containerId: entry.kind,
        sectionIndex: -1,
        blockIndex: entry.paragraphIndex,
        paragraphIndex: entry.paragraphIndex,
      },
      assetLookup
    ));
  });

  const features = summarizeFeatures(sections, blockRefs, {
    supplementalParts,
    headerTexts,
    footerTexts,
    headerPartPresent,
    footerPartPresent,
    headerEditableCount: headerPart.paragraphRefs.filter((entry) => entry.editable).length,
    footerEditableCount: footerPart.paragraphRefs.filter((entry) => entry.editable).length,
  });
  const warnings = buildWarnings(features, validation);

  return {
    zip,
    sections,
    supplementalParts,
    supplementalParagraphRefs,
    blockRefs,
    body: { sections: bodySections },
    supplementalText: supplementalParagraphRefs.map((entry) => entry.block),
    imageReferences,
    paragraphs: bodySections.flatMap((section) =>
      section.blocks.filter((block): block is ParagraphBlock => block.type === 'paragraph').map((block) => block.text)
    ).filter((text) => text.length > 0),
    rawText: plainTextSegments.join('\n\n'),
    headerTexts,
    footerTexts,
    features,
    warnings,
    packageValidation: validation,
  };
}

function requireSection(parsed: ParsedStructuredHwpx, sectionIndex: number): BlockRef[] {
  const section = parsed.blockRefs[sectionIndex];
  if (!section) {
    throw new Error(`sectionIndex ${sectionIndex} is out of range.`);
  }
  return section;
}

function requireBlock(parsed: ParsedStructuredHwpx, sectionIndex: number, blockIndex: number): BlockRef {
  const section = requireSection(parsed, sectionIndex);
  const block = section[blockIndex];
  if (!block) {
    throw new Error(`blockIndex ${blockIndex} is out of range.`);
  }
  return block;
}

function requireParagraphBlock(parsed: ParsedStructuredHwpx, sectionIndex: number, blockIndex: number): ParagraphBlockRef {
  const block = requireBlock(parsed, sectionIndex, blockIndex);
  if (block.type !== 'paragraph') {
    throw new Error(`Block ${sectionIndex}:${blockIndex} is not a paragraph.`);
  }
  if (!block.editable) {
    throw new Error(`Paragraph ${sectionIndex}:${blockIndex} contains preserved objects and cannot be edited as plain text.`);
  }
  return block;
}

function requireTableBlock(parsed: ParsedStructuredHwpx, sectionIndex: number, blockIndex: number): TableBlockRef {
  const block = requireBlock(parsed, sectionIndex, blockIndex);
  if (block.type !== 'table') {
    throw new Error(`Block ${sectionIndex}:${blockIndex} is not a table.`);
  }
  return block;
}

function requireTableCell(tableRef: TableBlockRef, rowIndex: number, columnIndex: number): TableCellRef {
  const cell = tableRef.cells.find((entry) => entry.rowIndex === rowIndex && entry.columnIndex === columnIndex);
  if (!cell) {
    throw new Error(`Cell r${rowIndex} c${columnIndex} was not found in the target table.`);
  }
  return cell;
}

async function loadManifestDocument(zip: JSZip): Promise<Document> {
  const file = zip.file(MAIN_CONTENT_PATH);
  if (!file) {
    throw new Error(`Missing required manifest part: ${MAIN_CONTENT_PATH}`);
  }
  const xml = await file.async('string');
  const validated = XMLValidator.validate(xml);
  if (validated !== true) {
    throw new Error(`Invalid XML in manifest part: ${MAIN_CONTENT_PATH}`);
  }
  const parser = new DOMParser();
  const document = parser.parseFromString(xml, 'application/xml');
  if (!document.documentElement) {
    throw new Error(`Missing root element in manifest part: ${MAIN_CONTENT_PATH}`);
  }
  return document;
}

function findManifestElement(document: Document): Element {
  const manifest = descendantElements(document.documentElement, 'manifest')[0];
  if (!manifest) {
    throw new Error('Manifest part does not contain an <opf:manifest> element.');
  }
  return manifest;
}

function findManifestItemElement(document: Document, itemId: string): Element | null {
  return descendantElements(document.documentElement, 'item')
    .find((element) => element.getAttribute('id') === itemId) ?? null;
}

function listManifestItemIds(document: Document): Set<string> {
  return new Set(
    descendantElements(document.documentElement, 'item')
      .map((element) => element.getAttribute('id')?.trim() ?? '')
      .filter((value) => value.length > 0)
  );
}

function inferImageMediaType(imagePath: string): string {
  const extension = extname(imagePath).toLowerCase();
  const mediaType = IMAGE_MEDIA_TYPES.get(extension);
  if (!mediaType) {
    throw new Error(`Unsupported replacement image type: ${extension || '<none>'}`);
  }
  return mediaType;
}

function buildReplacementAssetPath(existingAssetPath: string | null, binaryItemId: string, imagePath: string): string {
  const nextExtension = extname(imagePath).toLowerCase();
  if (existingAssetPath) {
    const dir = pathPosix.dirname(existingAssetPath);
    const baseName = pathPosix.basename(existingAssetPath, pathPosix.extname(existingAssetPath));
    return normalizePackagePath(pathPosix.join(dir, `${baseName}${nextExtension}`));
  }
  return normalizePackagePath(`BinData/${binaryItemId}${nextExtension}`);
}

function buildUniqueManifestItemId(document: Document, preferredBase: string): string {
  const existing = listManifestItemIds(document);
  const sanitizedBase = preferredBase.replace(/[^A-Za-z0-9_.-]+/g, '_') || 'image';
  if (!existing.has(sanitizedBase)) {
    return sanitizedBase;
  }
  let counter = 2;
  while (existing.has(`${sanitizedBase}_${counter}`)) {
    counter += 1;
  }
  return `${sanitizedBase}_${counter}`;
}

function buildUniqueDerivedAssetPath(
  zip: JSZip,
  document: Document,
  existingAssetPath: string | null,
  manifestItemId: string,
  imagePath: string
): string {
  const extension = extname(imagePath).toLowerCase();
  const currentPath = existingAssetPath ? normalizePackagePath(existingAssetPath) : null;
  const baseDir = currentPath ? pathPosix.dirname(currentPath) : 'BinData';
  const currentBase = currentPath ? pathPosix.basename(currentPath, pathPosix.extname(currentPath)) : manifestItemId;
  const existingPaths = new Set(
    descendantElements(document.documentElement, 'item')
      .map((element) => resolveManifestHref(zip, MAIN_CONTENT_PATH, element.getAttribute('href')?.trim() ?? ''))
      .filter((value): value is string => typeof value === 'string' && value.length > 0)
  );

  let index = 1;
  while (true) {
    const suffix = index === 1 ? manifestItemId : `${manifestItemId}_${index}`;
    const candidate = normalizePackagePath(pathPosix.join(baseDir, `${currentBase}_${suffix}${extension}`));
    if (!existingPaths.has(candidate) && zip.file(candidate) === null) {
      return candidate;
    }
    index += 1;
  }
}

function collectPictureAttributeValues(parsed: ParsedStructuredHwpx, attributeName: 'id' | 'instid'): Set<string> {
  const values = new Set<string>();
  const collect = (root: Element): void => {
    descendantElements(root, 'pic').forEach((pictureElement) => {
      const value = pictureElement.getAttribute(attributeName)?.trim() ?? '';
      if (value.length > 0) {
        values.add(value);
      }
    });
  };

  parsed.sections.forEach((section) => collect(section.root));
  parsed.supplementalParts.forEach((part) => collect(part.root));
  return values;
}

function buildUniquePictureAttributeValue(parsed: ParsedStructuredHwpx, attributeName: 'id' | 'instid'): string {
  const existing = collectPictureAttributeValues(parsed, attributeName);
  const numericValues = [...existing]
    .map((value) => Number.parseInt(value, 10))
    .filter((value) => Number.isFinite(value));

  let candidate = numericValues.length > 0 ? Math.max(...numericValues) + 1 : 1;
  while (existing.has(String(candidate))) {
    candidate += 1;
  }
  return String(candidate);
}

function appendChildBeforeLineSegArray(paragraphElement: Element, child: Element): void {
  const lineSegArray = elementChildren(paragraphElement).find((entry) => localNameOf(entry) === 'linesegarray') ?? null;
  if (lineSegArray) {
    paragraphElement.insertBefore(child, lineSegArray);
    return;
  }
  paragraphElement.appendChild(child);
}

function ensureParagraphInCell(cell: TableCellRef): Element {
  const paragraph = cell.paragraphs[0];
  if (paragraph) {
    return paragraph;
  }

  const nextParagraph = createEmptyParagraph(cell.subListElement.ownerDocument);
  setParagraphPlainText(nextParagraph, '');
  cell.subListElement.appendChild(nextParagraph);
  cell.paragraphs.push(nextParagraph);
  return nextParagraph;
}

function resolveDestinationParagraphByTargetId(parsed: ParsedStructuredHwpx, targetId: string): Element {
  for (const section of parsed.blockRefs) {
    for (const block of section) {
      if (block.type === 'paragraph' && block.block.blockId === targetId) {
        return block.paragraphElement;
      }
      if (block.type === 'table') {
        const cell = block.cells.find((entry) => entry.block.cellId === targetId);
        if (cell) {
          return ensureParagraphInCell(cell);
        }
      }
    }
  }

  const supplemental = parsed.supplementalParagraphRefs.find((entry) => entry.block.targetId === targetId);
  if (supplemental) {
    return supplemental.paragraphElement;
  }

  throw new Error(`Destination target "${targetId}" was not found in the current .hwpx package.`);
}

function findContainingRun(paragraphElement: Element, descendant: Element): Element {
  for (const run of directRunChildren(paragraphElement)) {
    if (run === descendant) {
      return run;
    }
    if (descendantElements(run).includes(descendant)) {
      return run;
    }
  }
  throw new Error('The requested picture element is not contained within a direct paragraph run.');
}

function findPrototypePayloadWithinRun(runElement: Element, pictureElement: Element): Element {
  let current: Node | null = pictureElement;
  while (current && current.parentNode && current.parentNode !== runElement) {
    current = current.parentNode;
  }
  if (!current || current.nodeType !== current.ELEMENT_NODE) {
    throw new Error('Failed to identify a prototype payload element for image insertion.');
  }
  return current as Element;
}

function applyPlacementPatchToPictureElement(pictureElement: Element, patch: NonNullable<Extract<TextMutationOperation, { op: 'update_image_placement' }>['patch']>): void {
  const size = findDirectChild(pictureElement, 'sz');
  const imgDim = findDirectChild(pictureElement, 'imgDim');
  const previousWidth = parseNullableInteger(imgDim?.getAttribute('dimwidth') ?? size?.getAttribute('width') ?? null);
  const previousHeight = parseNullableInteger(imgDim?.getAttribute('dimheight') ?? size?.getAttribute('height') ?? null);
  const nextWidth = typeof patch.width === 'number' ? patch.width : previousWidth;
  const nextHeight = typeof patch.height === 'number' ? patch.height : previousHeight;

  if (typeof patch.textWrap === 'string') {
    pictureElement.setAttribute('textWrap', patch.textWrap);
  }
  if (typeof patch.textFlow === 'string') {
    pictureElement.setAttribute('textFlow', patch.textFlow);
  }
  if (typeof patch.zOrder === 'number') {
    pictureElement.setAttribute('zOrder', String(patch.zOrder));
  }

  if (typeof patch.width === 'number' || typeof patch.height === 'number') {
    const ensuredSize = ensureDirectChildElement(pictureElement, HP_NS, 'sz', 'hp');
    const ensuredImgDim = ensureDirectChildElement(pictureElement, HP_NS, 'imgDim', 'hp');
    if (nextWidth !== null) {
      ensuredSize.setAttribute('width', String(nextWidth));
      ensuredSize.setAttribute('widthRelTo', ensuredSize.getAttribute('widthRelTo') || 'ABSOLUTE');
      ensuredImgDim.setAttribute('dimwidth', String(nextWidth));
    }
    if (nextHeight !== null) {
      ensuredSize.setAttribute('height', String(nextHeight));
      ensuredSize.setAttribute('heightRelTo', ensuredSize.getAttribute('heightRelTo') || 'ABSOLUTE');
      ensuredImgDim.setAttribute('dimheight', String(nextHeight));
    }
    maybeUpdatePictureClipForResizedImage(pictureElement, previousWidth, previousHeight, nextWidth, nextHeight, patch.clip);
    syncPictureRectangle(pictureElement, nextWidth, nextHeight);
  }

  if (patch.margins) {
    const inMargin = ensureDirectChildElement(pictureElement, HP_NS, 'inMargin', 'hp');
    setOptionalNumericAttribute(inMargin, 'left', patch.margins.left);
    setOptionalNumericAttribute(inMargin, 'right', patch.margins.right);
    setOptionalNumericAttribute(inMargin, 'top', patch.margins.top);
    setOptionalNumericAttribute(inMargin, 'bottom', patch.margins.bottom);
  }

  if (patch.clip) {
    const imgClip = ensureDirectChildElement(pictureElement, HP_NS, 'imgClip', 'hp');
    setOptionalNumericAttribute(imgClip, 'left', patch.clip.left);
    setOptionalNumericAttribute(imgClip, 'right', patch.clip.right);
    setOptionalNumericAttribute(imgClip, 'top', patch.clip.top);
    setOptionalNumericAttribute(imgClip, 'bottom', patch.clip.bottom);
  }
}

function buildSkeletonImageRun(
  paragraphElement: Element,
  parsed: ParsedStructuredHwpx,
  binaryItemId: string,
  size: { width: number; height: number }
): { runElement: Element; pictureElement: Element } {
  const document = paragraphElement.ownerDocument;
  const runElement = document.createElementNS(HP_NS, 'hp:run');
  const templateRun = directRunChildren(paragraphElement)[0] ?? null;
  runElement.setAttribute('charPrIDRef', templateRun?.getAttribute('charPrIDRef') || '0');

  const pictureElement = document.createElementNS(HP_NS, 'hp:pic');
  pictureElement.setAttribute('id', buildUniquePictureAttributeValue(parsed, 'id'));
  pictureElement.setAttribute('zOrder', '0');
  pictureElement.setAttribute('numberingType', 'PICTURE');
  pictureElement.setAttribute('textWrap', 'TOP_AND_BOTTOM');
  pictureElement.setAttribute('textFlow', 'BOTH_SIDES');
  pictureElement.setAttribute('lock', '0');
  pictureElement.setAttribute('dropcapstyle', 'None');
  pictureElement.setAttribute('href', '');
  pictureElement.setAttribute('groupLevel', '0');
  pictureElement.setAttribute('instid', buildUniquePictureAttributeValue(parsed, 'instid'));
  pictureElement.setAttribute('reverse', '0');

  const sizeElement = document.createElementNS(HP_NS, 'hp:sz');
  sizeElement.setAttribute('width', String(size.width));
  sizeElement.setAttribute('widthRelTo', 'ABSOLUTE');
  sizeElement.setAttribute('height', String(size.height));
  sizeElement.setAttribute('heightRelTo', 'ABSOLUTE');
  sizeElement.setAttribute('protect', '0');
  pictureElement.appendChild(sizeElement);

  const imgRect = document.createElementNS(HP_NS, 'hp:imgRect');
  const pt0 = document.createElementNS(HC_NS, 'hc:pt0');
  pt0.setAttribute('x', '0');
  pt0.setAttribute('y', '0');
  const pt1 = document.createElementNS(HC_NS, 'hc:pt1');
  pt1.setAttribute('x', String(size.width));
  pt1.setAttribute('y', '0');
  const pt2 = document.createElementNS(HC_NS, 'hc:pt2');
  pt2.setAttribute('x', String(size.width));
  pt2.setAttribute('y', String(size.height));
  const pt3 = document.createElementNS(HC_NS, 'hc:pt3');
  pt3.setAttribute('x', '0');
  pt3.setAttribute('y', String(size.height));
  imgRect.appendChild(pt0);
  imgRect.appendChild(pt1);
  imgRect.appendChild(pt2);
  imgRect.appendChild(pt3);
  pictureElement.appendChild(imgRect);

  const imgClip = document.createElementNS(HP_NS, 'hp:imgClip');
  imgClip.setAttribute('left', '0');
  imgClip.setAttribute('right', String(size.width));
  imgClip.setAttribute('top', '0');
  imgClip.setAttribute('bottom', String(size.height));
  pictureElement.appendChild(imgClip);

  const inMargin = document.createElementNS(HP_NS, 'hp:inMargin');
  inMargin.setAttribute('left', '0');
  inMargin.setAttribute('right', '0');
  inMargin.setAttribute('top', '0');
  inMargin.setAttribute('bottom', '0');
  pictureElement.appendChild(inMargin);

  const imgDim = document.createElementNS(HP_NS, 'hp:imgDim');
  imgDim.setAttribute('dimwidth', String(size.width));
  imgDim.setAttribute('dimheight', String(size.height));
  pictureElement.appendChild(imgDim);

  const img = document.createElementNS(HC_NS, 'hc:img');
  img.setAttribute('binaryItemIDRef', binaryItemId);
  img.setAttribute('bright', '0');
  img.setAttribute('contrast', '0');
  img.setAttribute('effect', 'REAL_PIC');
  img.setAttribute('alpha', '0');
  pictureElement.appendChild(img);

  runElement.appendChild(pictureElement);
  return {
    runElement,
    pictureElement,
  };
}

async function applyImageReplacementMutations(parsed: ParsedStructuredHwpx, mutations: TextMutationOperation[]): Promise<void> {
  const replacements = mutations.filter((mutation): mutation is Extract<TextMutationOperation, { op: 'replace_image_asset' }> => mutation.op === 'replace_image_asset');
  if (replacements.length === 0) {
    return;
  }

  const manifestDocument = await loadManifestDocument(parsed.zip);
  const manifestElement = findManifestElement(manifestDocument);

  for (const mutation of replacements) {
    if ((!mutation.targetId && !mutation.binaryItemId) || (mutation.targetId && mutation.binaryItemId)) {
      throw new Error('replace_image_asset requires exactly one of targetId or binaryItemId.');
    }

    const matchedReferences = mutation.binaryItemId
      ? parsed.imageReferences.filter((reference) => reference.binaryItemId === mutation.binaryItemId)
      : parsed.imageReferences.filter((reference) => reference.targetId === mutation.targetId);

    if (matchedReferences.length === 0) {
      throw new Error(
        mutation.binaryItemId
          ? `No image reference matched binaryItemId "${mutation.binaryItemId}".`
          : `No image reference matched targetId "${mutation.targetId}".`
      );
    }

    if (mutation.targetId && matchedReferences.length > 1) {
      throw new Error(`Target "${mutation.targetId}" contains ${matchedReferences.length} image references. Use binaryItemId for deterministic replacement.`);
    }

    const reference = matchedReferences[0];
    const replacementBytes = await readFile(mutation.imagePath);
    const nextMediaType = inferImageMediaType(mutation.imagePath);
    const binaryItemId = reference.binaryItemId;
    const nextAssetPath = buildReplacementAssetPath(reference.assetPath, binaryItemId, mutation.imagePath);
    const manifestItem = findManifestItemElement(manifestDocument, binaryItemId);

    if (manifestItem) {
      const previousHref = manifestItem.getAttribute('href')?.trim() ?? reference.assetPath ?? nextAssetPath;
      const previousResolvedPath = resolveManifestHref(parsed.zip, MAIN_CONTENT_PATH, previousHref);
      manifestItem.setAttribute('href', nextAssetPath);
      manifestItem.setAttribute('media-type', nextMediaType);
      if (previousResolvedPath && previousResolvedPath !== nextAssetPath) {
        parsed.zip.remove(previousResolvedPath);
      }
    } else {
      const created = manifestDocument.createElementNS(OPF_NS, 'opf:item');
      created.setAttribute('id', binaryItemId);
      created.setAttribute('href', nextAssetPath);
      created.setAttribute('media-type', nextMediaType);
      manifestElement.appendChild(created);
    }

    parsed.zip.file(nextAssetPath, replacementBytes);
  }

  parsed.zip.file(MAIN_CONTENT_PATH, serializeXmlDocument(manifestDocument, MAIN_CONTENT_PATH));
}

async function applyImageInsertionMutations(parsed: ParsedStructuredHwpx, mutations: TextMutationOperation[]): Promise<void> {
  const insertions = mutations.filter((mutation): mutation is Extract<TextMutationOperation, { op: 'insert_image_from_prototype' }> => mutation.op === 'insert_image_from_prototype');
  if (insertions.length === 0) {
    return;
  }

  const manifestDocument = await loadManifestDocument(parsed.zip);
  const manifestElement = findManifestElement(manifestDocument);

  for (const mutation of insertions) {
    const destinationParagraph = resolveDestinationParagraphByTargetId(parsed, mutation.destinationTargetId);
    const replacementBytes = await readFile(mutation.imagePath);
    const nextMediaType = inferImageMediaType(mutation.imagePath);
    const hasPrototypeLocator = Boolean(
      mutation.prototypeTargetId
      || mutation.prototypeBinaryItemId
      || mutation.prototypeObjectId
      || mutation.prototypeInstanceId
    );
    const matchedReferences = hasPrototypeLocator
      ? resolveImageReferencesByInstanceLocator(parsed, {
          targetId: mutation.prototypeTargetId,
          binaryItemId: mutation.prototypeBinaryItemId,
          objectId: mutation.prototypeObjectId,
          instanceId: mutation.prototypeInstanceId,
        })
      : [];

    if (hasPrototypeLocator && matchedReferences.length === 0) {
      throw new Error('No prototype image reference matched the requested locator.');
    }
    if (matchedReferences.length > 1) {
      throw new Error('The requested prototype locator matched multiple image references. Use a deterministic single-image locator.');
    }

    const prototypeReference = matchedReferences[0] ?? null;
    const baseId = prototypeReference?.binaryItemId
      ? `${prototypeReference.binaryItemId}_inserted`
      : 'inserted_image';
    const nextBinaryItemId = buildUniqueManifestItemId(manifestDocument, baseId);
    const nextAssetPath = buildUniqueDerivedAssetPath(parsed.zip, manifestDocument, prototypeReference?.assetPath ?? null, nextBinaryItemId, mutation.imagePath);

    const created = manifestDocument.createElementNS(OPF_NS, 'opf:item');
    created.setAttribute('id', nextBinaryItemId);
    created.setAttribute('href', nextAssetPath);
    created.setAttribute('media-type', nextMediaType);
    manifestElement.appendChild(created);
    parsed.zip.file(nextAssetPath, replacementBytes);

    let nextRun: Element;
    let nextPictureElement: Element;
    if (prototypeReference) {
      const prototypePictureElements = pictureElementsForReference(parsed, prototypeReference);
      if (prototypePictureElements.length !== 1) {
        throw new Error('The requested prototype image did not resolve to exactly one picture element.');
      }

      const prototypePictureElement = prototypePictureElements[0];
      const prototypeParagraph = paragraphElementForImageReference(parsed, prototypeReference);
      const prototypeRun = findContainingRun(prototypeParagraph, prototypePictureElement);
      const prototypePayload = findPrototypePayloadWithinRun(prototypeRun, prototypePictureElement);
      nextRun = prototypeRun.cloneNode(false) as Element;
      const nextPayload = prototypePayload.cloneNode(true) as Element;
      const resolvedNextPictureElement = nextPayload.nodeType === nextPayload.ELEMENT_NODE && localNameOf(nextPayload) === 'pic'
        ? nextPayload
        : findDirectDescendantByName(nextPayload, 'pic');
      if (!resolvedNextPictureElement) {
        throw new Error('The prototype payload does not contain a picture element.');
      }
      nextPictureElement = resolvedNextPictureElement;

      const nextImageElement = findDirectDescendantByName(nextPictureElement, 'img');
      if (!nextImageElement) {
        throw new Error('The prototype picture element does not contain an image binding.');
      }

      nextPictureElement.setAttribute('id', buildUniquePictureAttributeValue(parsed, 'id'));
      nextPictureElement.setAttribute('instid', buildUniquePictureAttributeValue(parsed, 'instid'));
      nextImageElement.setAttribute('binaryItemIDRef', nextBinaryItemId);
      nextRun.appendChild(nextPayload);
    } else {
      const inferredSize = inferInsertedPictureSize(mutation.imagePath, replacementBytes, mutation.patch);
      const createdSkeleton = buildSkeletonImageRun(destinationParagraph, parsed, nextBinaryItemId, inferredSize);
      nextRun = createdSkeleton.runElement;
      nextPictureElement = createdSkeleton.pictureElement;
    }

    if (mutation.patch) {
      applyPlacementPatchToPictureElement(nextPictureElement, mutation.patch);
    }
    appendChildBeforeLineSegArray(destinationParagraph, nextRun);
  }

  parsed.zip.file(MAIN_CONTENT_PATH, serializeXmlDocument(manifestDocument, MAIN_CONTENT_PATH));
}

function resolveImageReferencesByInstanceLocator(
  parsed: ParsedStructuredHwpx,
  locator: { targetId?: string; binaryItemId?: string; objectId?: string; instanceId?: string }
): EmbeddedImageReference[] {
  const locatorCount = [locator.targetId, locator.binaryItemId, locator.objectId, locator.instanceId]
    .filter((value) => typeof value === 'string' && value.length > 0)
    .length;
  if (locatorCount !== 1) {
    throw new Error('Expected exactly one image locator.');
  }

  if (locator.targetId) {
    return parsed.imageReferences.filter((reference) => reference.targetId === locator.targetId);
  }
  if (locator.binaryItemId) {
    return parsed.imageReferences.filter((reference) => reference.binaryItemId === locator.binaryItemId);
  }
  if (locator.objectId) {
    return parsed.imageReferences.filter((reference) => reference.placement?.objectId === locator.objectId);
  }
  return parsed.imageReferences.filter((reference) => reference.placement?.instanceId === locator.instanceId);
}

async function applyImageInstanceReplacementMutations(parsed: ParsedStructuredHwpx, mutations: TextMutationOperation[]): Promise<void> {
  const replacements = mutations.filter((mutation): mutation is Extract<TextMutationOperation, { op: 'replace_image_instance_asset' }> => mutation.op === 'replace_image_instance_asset');
  if (replacements.length === 0) {
    return;
  }

  const manifestDocument = await loadManifestDocument(parsed.zip);
  const manifestElement = findManifestElement(manifestDocument);

  for (const mutation of replacements) {
    const matchedReferences = resolveImageReferencesByInstanceLocator(parsed, mutation);
    if (matchedReferences.length === 0) {
      throw new Error('No image reference matched the requested locator.');
    }
    if (matchedReferences.length > 1) {
      throw new Error('The requested locator matched multiple image references. Use a deterministic single-image locator.');
    }

    const reference = matchedReferences[0];
    const pictureElements = pictureElementsForReference(parsed, reference);
    if (pictureElements.length !== 1) {
      throw new Error('The requested image instance did not resolve to exactly one picture element.');
    }

    const pictureElement = pictureElements[0];
    const imageElement = findDirectDescendantByName(pictureElement, 'img');
    if (!imageElement) {
      throw new Error('The target picture element does not contain an image binding.');
    }

    const replacementBytes = await readFile(mutation.imagePath);
    const nextMediaType = inferImageMediaType(mutation.imagePath);
    const baseId = reference.placement?.objectId
      ? `${reference.binaryItemId}_${reference.placement.objectId}`
      : `${reference.binaryItemId}_instance`;
    const nextBinaryItemId = buildUniqueManifestItemId(manifestDocument, baseId);
    const nextAssetPath = buildUniqueDerivedAssetPath(parsed.zip, manifestDocument, reference.assetPath, nextBinaryItemId, mutation.imagePath);

    const created = manifestDocument.createElementNS(OPF_NS, 'opf:item');
    created.setAttribute('id', nextBinaryItemId);
    created.setAttribute('href', nextAssetPath);
    created.setAttribute('media-type', nextMediaType);
    manifestElement.appendChild(created);
    parsed.zip.file(nextAssetPath, replacementBytes);
    imageElement.setAttribute('binaryItemIDRef', nextBinaryItemId);
  }

  parsed.zip.file(MAIN_CONTENT_PATH, serializeXmlDocument(manifestDocument, MAIN_CONTENT_PATH));
}

function paragraphElementForImageReference(parsed: ParsedStructuredHwpx, reference: EmbeddedImageReference): Element {
  switch (reference.kind) {
    case 'paragraph': {
      const block = requireBlock(parsed, reference.sectionIndex, reference.blockIndex);
      if (block.type !== 'paragraph') {
        throw new Error(`Block ${reference.sectionIndex}:${reference.blockIndex} is not a paragraph.`);
      }
      return block.paragraphElement;
    }
    case 'table_cell': {
      const table = requireTableBlock(parsed, reference.sectionIndex, reference.blockIndex);
      const cell = requireTableCell(table, reference.rowIndex ?? -1, reference.columnIndex ?? -1);
      const paragraph = cell.paragraphs[reference.paragraphIndex ?? 0];
      if (!paragraph) {
        throw new Error(`Paragraph index ${reference.paragraphIndex ?? 0} was not found in target ${reference.targetId}.`);
      }
      return paragraph;
    }
    case 'header_paragraph': {
      const paragraph = parsed.supplementalParagraphRefs.find((entry) => entry.kind === 'header' && entry.paragraphIndex === (reference.paragraphIndex ?? 0));
      if (!paragraph) {
        throw new Error(`header paragraph ${reference.paragraphIndex ?? 0} was not found in the target package.`);
      }
      return paragraph.paragraphElement;
    }
    case 'footer_paragraph': {
      const paragraph = parsed.supplementalParagraphRefs.find((entry) => entry.kind === 'footer' && entry.paragraphIndex === (reference.paragraphIndex ?? 0));
      if (!paragraph) {
        throw new Error(`footer paragraph ${reference.paragraphIndex ?? 0} was not found in the target package.`);
      }
      return paragraph.paragraphElement;
    }
  }
}

function pictureElementsForReference(parsed: ParsedStructuredHwpx, reference: EmbeddedImageReference): Element[] {
  const paragraphElement = paragraphElementForImageReference(parsed, reference);
  return descendantElements(paragraphElement, 'pic').filter((pictureElement) => {
    const imageElement = findDirectDescendantByName(pictureElement, 'img');
    if (imageElement?.getAttribute('binaryItemIDRef')?.trim() !== reference.binaryItemId) {
      return false;
    }
    if (reference.placement?.objectId && pictureElement.getAttribute('id')?.trim() !== reference.placement.objectId) {
      return false;
    }
    if (reference.placement?.instanceId && pictureElement.getAttribute('instid')?.trim() !== reference.placement.instanceId) {
      return false;
    }
    return true;
  });
}

function pruneEmptyRunElement(runElement: Element): void {
  const hasElementChildren = elementChildren(runElement).length > 0;
  const hasTextChildren = (() => {
    for (let child = runElement.firstChild; child; child = child.nextSibling) {
      if (child.nodeType === child.TEXT_NODE && (child.nodeValue ?? '').trim().length > 0) {
        return true;
      }
    }
    return false;
  })();

  if (!hasElementChildren && !hasTextChildren) {
    runElement.parentNode?.removeChild(runElement);
  }
}

function countBinaryItemReferencesInPackage(parsed: ParsedStructuredHwpx, binaryItemId: string): number {
  let count = 0;
  const collect = (root: Element): void => {
    descendantElements(root, 'img').forEach((imageElement) => {
      if (imageElement.getAttribute('binaryItemIDRef')?.trim() === binaryItemId) {
        count += 1;
      }
    });
  };

  parsed.sections.forEach((section) => collect(section.root));
  parsed.supplementalParts.forEach((part) => collect(part.root));
  return count;
}

async function applyImageDeletionMutations(parsed: ParsedStructuredHwpx, mutations: TextMutationOperation[]): Promise<void> {
  const deletions = mutations.filter((mutation): mutation is Extract<TextMutationOperation, { op: 'delete_image_instance' }> => mutation.op === 'delete_image_instance');
  if (deletions.length === 0) {
    return;
  }

  const manifestDocument = await loadManifestDocument(parsed.zip);

  for (const mutation of deletions) {
    const matchedReferences = resolveImageReferencesByInstanceLocator(parsed, mutation);
    if (matchedReferences.length === 0) {
      throw new Error('No image reference matched the requested locator.');
    }
    if (matchedReferences.length > 1) {
      throw new Error('The requested locator matched multiple image references. Use a deterministic single-image locator.');
    }

    const reference = matchedReferences[0];
    const pictureElements = pictureElementsForReference(parsed, reference);
    if (pictureElements.length !== 1) {
      throw new Error('The requested image instance did not resolve to exactly one picture element.');
    }

    const pictureElement = pictureElements[0];
    const paragraphElement = paragraphElementForImageReference(parsed, reference);
    const runElement = findContainingRun(paragraphElement, pictureElement);
    const payloadElement = findPrototypePayloadWithinRun(runElement, pictureElement);

    runElement.removeChild(payloadElement);
    pruneEmptyRunElement(runElement);

    const remainingReferences = countBinaryItemReferencesInPackage(parsed, reference.binaryItemId);
    if (remainingReferences === 0) {
      const manifestItem = findManifestItemElement(manifestDocument, reference.binaryItemId);
      if (manifestItem) {
        const href = manifestItem.getAttribute('href')?.trim() ?? reference.assetPath ?? '';
        const resolvedPath = resolveManifestHref(parsed.zip, MAIN_CONTENT_PATH, href);
        manifestItem.parentNode?.removeChild(manifestItem);
        if (resolvedPath) {
          parsed.zip.remove(resolvedPath);
        }
      }
    }
  }

  parsed.zip.file(MAIN_CONTENT_PATH, serializeXmlDocument(manifestDocument, MAIN_CONTENT_PATH));
}

function applyImagePlacementMutations(parsed: ParsedStructuredHwpx, mutations: TextMutationOperation[]): void {
  const updates = mutations.filter((mutation): mutation is Extract<TextMutationOperation, { op: 'update_image_placement' }> => mutation.op === 'update_image_placement');
  if (updates.length === 0) {
    return;
  }

  for (const mutation of updates) {
    const matchedReferences = resolveImageReferencesByInstanceLocator(parsed, mutation);

    if (matchedReferences.length === 0) {
      const locatorLabel = mutation.targetId
        ? `targetId "${mutation.targetId}"`
        : mutation.binaryItemId
          ? `binaryItemId "${mutation.binaryItemId}"`
          : mutation.objectId
            ? `objectId "${mutation.objectId}"`
            : `instanceId "${mutation.instanceId}"`;
      throw new Error(`No image reference matched ${locatorLabel}.`);
    }
    if (matchedReferences.length > 1) {
      const locatorLabel = mutation.targetId
        ? `targetId "${mutation.targetId}"`
        : mutation.binaryItemId
          ? `binaryItemId "${mutation.binaryItemId}"`
          : mutation.objectId
            ? `objectId "${mutation.objectId}"`
            : `instanceId "${mutation.instanceId}"`;
      throw new Error(`${locatorLabel} matched ${matchedReferences.length} image references. Use a deterministic single-image locator.`);
    }

    const reference = matchedReferences[0];
    const pictureElements = pictureElementsForReference(parsed, reference);
    if (pictureElements.length !== 1) {
      throw new Error(`Image reference "${reference.binaryItemId}" did not resolve to exactly one picture element.`);
    }

    const pictureElement = pictureElements[0];
    applyPlacementPatchToPictureElement(pictureElement, mutation.patch);
  }
}

function requireSupplementalParagraph(
  parsed: ParsedStructuredHwpx,
  kind: SupplementalTextKind,
  paragraphIndex: number
): SupplementalParagraphRef {
  const paragraph = parsed.supplementalParagraphRefs.find((entry) => entry.kind === kind && entry.paragraphIndex === paragraphIndex);
  if (!paragraph) {
    throw new Error(`${kind} paragraph ${paragraphIndex} was not found in the target package.`);
  }
  if (!paragraph.editable) {
    throw new Error(`${kind} paragraph ${paragraphIndex} contains preserved objects and cannot be edited as plain text.`);
  }
  return paragraph;
}

function applyMutations(parsed: ParsedStructuredHwpx, mutations: TextMutationOperation[]): void {
  for (const mutation of mutations) {
    switch (mutation.op) {
      case 'replace_text_in_paragraph': {
        const paragraph = requireParagraphBlock(parsed, mutation.sectionIndex, mutation.blockIndex);
        setParagraphPlainText(paragraph.paragraphElement, mutation.text.trim());
        break;
      }

      case 'replace_text_in_supplemental_paragraph': {
        const paragraph = requireSupplementalParagraph(parsed, mutation.kind, mutation.paragraphIndex);
        setParagraphPlainText(paragraph.paragraphElement, mutation.text.trim());
        break;
      }

      case 'replace_image_asset':
      case 'replace_image_instance_asset':
      case 'delete_image_instance':
      case 'insert_image_from_prototype':
      case 'update_image_placement':
        break;

      case 'insert_paragraph_after': {
        const block = requireBlock(parsed, mutation.sectionIndex, mutation.blockIndex);
        const anchor = getAnchorParagraph(block);
        const nextParagraph = createParagraphFromTemplate(anchor, mutation.text.trim());
        anchor.parentNode?.insertBefore(nextParagraph, anchor.nextSibling);
        break;
      }

      case 'delete_paragraph': {
        const paragraph = requireParagraphBlock(parsed, mutation.sectionIndex, mutation.blockIndex);
        const sectionRoot = paragraph.paragraphElement.parentNode as Element | null;
        sectionRoot?.removeChild(paragraph.paragraphElement);
        if (sectionRoot) {
          ensureSectionHasParagraph(sectionRoot);
        }
        break;
      }

      case 'replace_table_cell_text': {
        const table = requireTableBlock(parsed, mutation.sectionIndex, mutation.blockIndex);
        const cell = requireTableCell(table, mutation.rowIndex, mutation.columnIndex);
        setCellText(cell, mutation.text);
        break;
      }

      case 'insert_table_row': {
        const table = requireTableBlock(parsed, mutation.sectionIndex, mutation.blockIndex);
        insertTableRow(table, mutation.rowIndex, mutation.boundaryPolicy ?? 'reject');
        break;
      }

      case 'delete_table_row': {
        const table = requireTableBlock(parsed, mutation.sectionIndex, mutation.blockIndex);
        deleteTableRow(table, mutation.rowIndex, mutation.boundaryPolicy ?? 'reject');
        break;
      }

      case 'clone_table_region': {
        const table = requireTableBlock(parsed, mutation.sectionIndex, mutation.blockIndex);
        cloneTableRegion(
          table,
          mutation.templateStartRowIndex,
          mutation.templateEndRowIndex,
          mutation.insertAfterRowIndex,
          mutation.boundaryPolicy ?? 'reject'
        );
        break;
      }
    }
  }
}

export async function saveStructuredHwpxWithMutations(filePath: string, mutations: TextMutationOperation[]): Promise<ParsedStructuredHwpx> {
  const parsed = await parseStructuredHwpx(filePath);
  await applyImageReplacementMutations(parsed, mutations);
  await applyImageInstanceReplacementMutations(parsed, mutations);
  await applyImageDeletionMutations(parsed, mutations);
  await applyImageInsertionMutations(parsed, mutations);
  applyImagePlacementMutations(parsed, mutations);
  applyMutations(parsed, mutations);

  for (const section of parsed.sections) {
    parsed.zip.file(section.fileName, serializeSection(section));
  }
  for (const part of parsed.supplementalParts) {
    parsed.zip.file(part.fileName, serializeSupplementalPart(part));
  }

  await writeValidatedHwpxPackage(filePath, parsed.zip);
  return parseStructuredHwpx(filePath);
}

export async function saveStructuredHwpxWithPlainText(filePath: string, text: string): Promise<ParsedStructuredHwpx> {
  const parsed = await parseStructuredHwpx(filePath);
  if (!parsed.features.allowPlainTextSave) {
    throw new Error('This document cannot be saved from a plain-text projection. Use structured mutations instead.');
  }

  const paragraphBlocks = parsed.blockRefs.flat().filter((block): block is ParagraphBlockRef => block.type === 'paragraph');
  const paragraphs = splitParagraphText(text);
  const values = paragraphs.length > 0 ? paragraphs : [''];

  paragraphBlocks.forEach((block, index) => {
    if (index < values.length) {
      setParagraphPlainText(block.paragraphElement, values[index] ?? '');
      return;
    }
    const sectionRoot = block.paragraphElement.parentNode as Element | null;
    if (sectionRoot) {
      sectionRoot.removeChild(block.paragraphElement);
      ensureSectionHasParagraph(sectionRoot);
    }
  });

  const currentParagraphs = paragraphBlocks.filter((block) => block.paragraphElement.parentNode);
  if (values.length > currentParagraphs.length) {
    const anchor = currentParagraphs[currentParagraphs.length - 1]?.paragraphElement
      ?? parsed.sections[0]?.root.getElementsByTagNameNS(HP_NS, 'p')[0]
      ?? null;

    if (!anchor) {
      throw new Error('Failed to find a paragraph anchor for plain-text save.');
    }

    let lastParagraph = anchor;
    values.slice(currentParagraphs.length).forEach((value) => {
      const nextParagraph = createParagraphFromTemplate(lastParagraph, value);
      lastParagraph.parentNode?.insertBefore(nextParagraph, lastParagraph.nextSibling);
      lastParagraph = nextParagraph;
    });
  }

  for (const section of parsed.sections) {
    parsed.zip.file(section.fileName, serializeSection(section));
  }
  for (const part of parsed.supplementalParts) {
    parsed.zip.file(part.fileName, serializeSupplementalPart(part));
  }

  await writeValidatedHwpxPackage(filePath, parsed.zip);
  return parseStructuredHwpx(filePath);
}

export async function copyStructuredHwpx(sourcePath: string, outputPath: string): Promise<ParsedStructuredHwpx> {
  await copyFile(sourcePath, outputPath);
  return parseStructuredHwpx(outputPath);
}
