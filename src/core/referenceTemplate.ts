import { unlink } from 'node:fs/promises';
import { basename } from 'node:path';
import type {
  ApiWarning,
  NormalizedDocumentBody,
  ParagraphBlock,
  ReferencePlaceholderSummary,
  ReferenceRepeatableRegionSummary,
  ReferenceStyleSummary,
  ReferenceTokenOccurrence,
  ReferenceTokenSummary,
  ReferenceTableSummary,
  SupplementalTextBlock,
  TableRepeatRegionAnalysis,
  ReferenceTemplateAnalysis,
  SourceFormat,
  TableBlock,
} from '../types.js';
import { writeTempFile } from '../utils/files.js';
import { convertHwpToHwpxWithJvm } from './jvmCore.js';
import { buildImageTargetSummaries } from './imageTargets.js';
import { buildImageResourceSummaries } from './imageResources.js';
import { parseStructuredHwpx } from './structuredHwpx.js';
import { analyzeTableRepeatRegion } from './tableRepeatAnalysis.js';

const PLACEHOLDER_PATTERNS = [
  /^\[[^\]]+\]$/,
  /^\{\{[^}]+\}\}$/,
  /^<[^>]+>$/,
];

const TOKEN_PATTERNS = [
  /\[[^\]]+\]/g,
  /\{\{[^}]+\}\}/g,
  /<[^>\s][^>]*>/g,
];

function looksLikePlaceholder(text: string): boolean {
  const normalized = text.trim();
  if (normalized.length === 0) return false;
  return PLACEHOLDER_PATTERNS.some((pattern) => pattern.test(normalized));
}

function extractTokens(text: string): string[] {
  const tokens = new Set<string>();
  for (const pattern of TOKEN_PATTERNS) {
    const matches = text.match(pattern) ?? [];
    matches.forEach((match) => tokens.add(match));
  }
  return [...tokens];
}

function buildStyleSummaries(paragraphs: ParagraphBlock[]): ReferenceStyleSummary[] {
  const counters = new Map<string, ReferenceStyleSummary>();

  for (const paragraph of paragraphs) {
    const key = paragraph.styleRef ?? 'null';
    const current = counters.get(key) ?? {
      styleRef: paragraph.styleRef,
      paragraphCount: 0,
      sampleText: null,
    };
    current.paragraphCount += 1;
    if (!current.sampleText && paragraph.text.trim().length > 0) {
      current.sampleText = paragraph.text.trim();
    }
    counters.set(key, current);
  }

  return [...counters.values()].sort((left, right) => right.paragraphCount - left.paragraphCount);
}

function buildTableSummaries(tables: TableBlock[]): ReferenceTableSummary[] {
  return tables.map((table) => {
    const firstRowPreview = table.cells
      .filter((cell) => cell.rowIndex === 0)
      .sort((left, right) => left.columnIndex - right.columnIndex)
      .map((cell) => cell.text)
      .slice(0, 8);

    return {
      blockId: table.blockId,
      rowCount: table.rowCount,
      columnCount: table.columnCount,
      mergedCellCount: table.cells.filter((cell) => cell.rowSpan > 1 || cell.colSpan > 1).length,
      firstRowPreview,
    };
  });
}

function buildRepeatableRegions(tables: TableBlock[]): ReferenceRepeatableRegionSummary[] {
  const regions: ReferenceRepeatableRegionSummary[] = [];

  tables.forEach((table) => {
    const boundarySafe = Array.from({ length: table.rowCount + 1 }, (_, boundaryIndex) =>
      !table.cells.some(
        (cell) =>
          cell.rowIndex < boundaryIndex
          && cell.rowIndex + cell.rowSpan > boundaryIndex
      )
    );

    let runStart: number | null = null;
    for (let boundaryIndex = 0; boundaryIndex < boundarySafe.length; boundaryIndex += 1) {
      if (boundarySafe[boundaryIndex]) {
        if (runStart === null) {
          runStart = boundaryIndex;
        }
        continue;
      }

      if (runStart !== null && boundaryIndex - runStart > 1) {
        regions.push({
          tableBlockId: table.blockId,
          startRowIndex: runStart,
          endRowIndex: boundaryIndex - 1,
          rowCount: boundaryIndex - runStart,
          columnCount: table.columnCount,
          hasMergedCells: table.cells.some((cell) => cell.rowSpan > 1 || cell.colSpan > 1),
          boundaryCount: boundaryIndex - runStart + 1,
          supportedBoundaryPolicies: ['reject', 'split_boundary_merges'],
          recommendedBoundaryPolicy: 'reject',
        });
      }
      runStart = null;
    }

    if (runStart !== null && boundarySafe.length - runStart > 1) {
      regions.push({
        tableBlockId: table.blockId,
        startRowIndex: runStart,
        endRowIndex: boundarySafe.length - 2,
        rowCount: boundarySafe.length - runStart - 1,
        columnCount: table.columnCount,
        hasMergedCells: table.cells.some((cell) => cell.rowSpan > 1 || cell.colSpan > 1),
        boundaryCount: boundarySafe.length - runStart,
        supportedBoundaryPolicies: ['reject', 'split_boundary_merges'],
        recommendedBoundaryPolicy: 'reject',
      });
    }
  });

  return regions;
}

function buildRepeatRegionRecommendations(
  body: NormalizedDocumentBody,
  tables: TableBlock[],
  strictRegions: ReferenceRepeatableRegionSummary[]
): TableRepeatRegionAnalysis[] {
  const recommendations = new Map<string, TableRepeatRegionAnalysis>();

  const addRecommendation = (tableBlockId: string, startRowIndex: number, endRowIndex: number) => {
    const analyzed = analyzeTableRepeatRegion(body, {
      tableBlockId,
      templateRowIndex: startRowIndex,
      templateEndRowIndex: endRowIndex,
    });
    if (!analyzed) return;
    if (endRowIndex < startRowIndex) return;
    const key = `${tableBlockId}:${startRowIndex}:${endRowIndex}`;
    recommendations.set(key, analyzed.analysis);
  };

  strictRegions.forEach((region) => {
    addRecommendation(region.tableBlockId, region.startRowIndex, region.endRowIndex);
  });

  tables.forEach((table) => {
    table.cells
      .filter((cell) => cell.rowSpan > 1)
      .forEach((cell) => {
        const coveredStart = cell.rowIndex + 1;
        const coveredEnd = cell.rowIndex + cell.rowSpan - 1;
        if (coveredStart <= coveredEnd) {
          addRecommendation(table.blockId, coveredStart, coveredEnd);
        }

        const innerEnd = cell.rowIndex + cell.rowSpan - 2;
        if (coveredStart <= innerEnd) {
          addRecommendation(table.blockId, coveredStart, innerEnd);
        }
      });
  });

  return [...recommendations.values()]
    .sort((left, right) => {
      if (left.tableBlockId !== right.tableBlockId) {
        return left.tableBlockId.localeCompare(right.tableBlockId);
      }
      if (left.templateRowIndex !== right.templateRowIndex) {
        return left.templateRowIndex - right.templateRowIndex;
      }
      return left.templateEndRowIndex - right.templateEndRowIndex;
    })
    .slice(0, 96);
}

function buildPlaceholderSummaries(
  paragraphs: ParagraphBlock[],
  tables: TableBlock[],
  supplementalText: SupplementalTextBlock[]
): ReferencePlaceholderSummary[] {
  const placeholders: ReferencePlaceholderSummary[] = [];

  for (const paragraph of paragraphs) {
    if (looksLikePlaceholder(paragraph.text)) {
      placeholders.push({
        kind: 'paragraph',
        blockId: paragraph.blockId,
        text: paragraph.text,
      });
    }
  }

  for (const table of tables) {
    for (const cell of table.cells) {
      if (looksLikePlaceholder(cell.text)) {
        placeholders.push({
          kind: 'table_cell',
          blockId: cell.cellId,
          text: cell.text,
        });
      }
    }
  }

  supplementalText.forEach((block) => {
    if (!block.editable || !looksLikePlaceholder(block.text)) {
      return;
    }
    placeholders.push({
      kind: block.kind === 'header' ? 'header_paragraph' : 'footer_paragraph',
      blockId: block.targetId,
      text: block.text,
    });
  });

  return placeholders.slice(0, 64);
}

function buildTokenRegistry(
  paragraphs: ParagraphBlock[],
  tables: TableBlock[],
  supplementalText: SupplementalTextBlock[]
): ReferenceTokenSummary[] {
  const registry = new Map<string, ReferenceTokenOccurrence[]>();

  paragraphs.forEach((paragraph) => {
    extractTokens(paragraph.text).forEach((token) => {
      const bucket = registry.get(token) ?? [];
      bucket.push({
        kind: 'paragraph',
        targetId: paragraph.blockId,
        containerId: paragraph.blockId,
        sectionIndex: Number(paragraph.blockId.split(':')[0] ?? 0),
        blockIndex: Number(paragraph.blockId.split(':')[1] ?? 0),
        textPreview: paragraph.text,
      });
      registry.set(token, bucket);
    });
  });

  tables.forEach((table) => {
    const blockParts = table.blockId.split(':');
    const sectionIndex = Number(blockParts[0] ?? 0);
    const blockIndex = Number(blockParts[1] ?? 0);

    table.cells.forEach((cell) => {
      extractTokens(cell.text).forEach((token) => {
        const bucket = registry.get(token) ?? [];
        bucket.push({
          kind: 'table_cell',
          targetId: cell.cellId,
          containerId: table.blockId,
          sectionIndex,
          blockIndex,
          rowIndex: cell.rowIndex,
          columnIndex: cell.columnIndex,
          textPreview: cell.text,
        });
        registry.set(token, bucket);
      });
    });
  });

  supplementalText.forEach((block) => {
    if (!block.editable) {
      return;
    }
    extractTokens(block.text).forEach((token) => {
      const bucket = registry.get(token) ?? [];
      bucket.push({
        kind: block.kind === 'header' ? 'header_paragraph' : 'footer_paragraph',
        targetId: block.targetId,
        containerId: block.kind,
        sectionIndex: -1,
        blockIndex: block.paragraphIndex,
        textPreview: block.text,
      });
      registry.set(token, bucket);
    });
  });

  return [...registry.entries()]
    .map(([token, occurrences]) => ({
      token,
      occurrenceCount: occurrences.length,
      ambiguous: occurrences.length > 1,
      occurrences,
    }))
    .sort((left, right) => left.token.localeCompare(right.token));
}

export async function analyzeReferenceTemplate(sourcePath: string, sourceFormat: SourceFormat): Promise<ReferenceTemplateAnalysis> {
  let canonicalPath = sourcePath;
  let importedFromHwp = false;
  const warnings: ApiWarning[] = [];

  if (sourceFormat === 'hwp') {
    canonicalPath = await writeTempFile(Buffer.alloc(0), '.hwpx');
    await convertHwpToHwpxWithJvm(sourcePath, canonicalPath);
    importedFromHwp = true;
    warnings.push({
      code: 'REFERENCE_IMPORTED_TO_HWPX',
      message: 'Reference analysis was performed on an imported HWPX model derived from the original HWP source.',
    });
  }

  try {
    const parsed = await parseStructuredHwpx(canonicalPath);
    const paragraphs = parsed.body.sections.flatMap((section) =>
      section.blocks.filter((block): block is ParagraphBlock => block.type === 'paragraph')
    );
    const tables = parsed.body.sections.flatMap((section) =>
      section.blocks.filter((block): block is TableBlock => block.type === 'table')
    );
    const supplementalText = parsed.supplementalText;
    const detectedTokens = new Set<string>();

    paragraphs.forEach((paragraph) => {
      extractTokens(paragraph.text).forEach((token) => detectedTokens.add(token));
    });
    tables.forEach((table) => {
      table.cells.forEach((cell) => {
        extractTokens(cell.text).forEach((token) => detectedTokens.add(token));
      });
    });
    supplementalText.forEach((block) => {
      if (!block.editable) {
        return;
      }
      extractTokens(block.text).forEach((token) => detectedTokens.add(token));
    });
    const tokenRegistry = buildTokenRegistry(paragraphs, tables, supplementalText);

    const repeatableRegions = buildRepeatableRegions(tables);

    return {
      sourcePath,
      sourceFormat,
      canonicalFormat: 'hwpx',
      importedFromHwp,
      title: basename(sourcePath),
      sectionCount: parsed.body.sections.length,
      paragraphCount: paragraphs.length,
      tableCount: tables.length,
      assetCount: parsed.packageValidation.assetEntries.length,
      unsupportedFeatures: parsed.features.unsupported,
      packageValidation: parsed.packageValidation,
      headerPreview: parsed.headerTexts.slice(0, 8),
      footerPreview: parsed.footerTexts.slice(0, 8),
      styles: buildStyleSummaries(paragraphs),
      tables: buildTableSummaries(tables),
      repeatableRegions,
      repeatRegionRecommendations: buildRepeatRegionRecommendations(parsed.body, tables, repeatableRegions),
      placeholders: buildPlaceholderSummaries(paragraphs, tables, supplementalText),
      imageReferences: parsed.imageReferences.slice(0, 96),
      imageTargets: buildImageTargetSummaries(parsed.imageReferences).slice(0, 96),
      imageResources: buildImageResourceSummaries(parsed.imageReferences).slice(0, 96),
      detectedTokens: [...detectedTokens].sort(),
      tokenRegistry,
      textPreview: paragraphs
        .map((paragraph) => paragraph.text.trim())
        .filter((text) => text.length > 0)
        .slice(0, 12),
      warnings: [...warnings, ...parsed.warnings],
    };
  } finally {
    if (importedFromHwp && canonicalPath !== sourcePath) {
      await unlink(canonicalPath).catch(() => undefined);
    }
  }
}
