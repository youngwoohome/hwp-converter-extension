import type {
  NormalizedDocumentBody,
  TableBlock,
  TableCellBlock,
  TableRepeatInstruction,
  TableRepeatRegionAnalysis,
} from '../types.js';

type LocatedTable = {
  sectionIndex: number;
  blockIndex: number;
  table: TableBlock;
};

function cellKey(cell: TableCellBlock): string {
  return `${cell.rowIndex}:${cell.columnIndex}`;
}

function crossingKey(cell: { rowIndex: number; columnIndex: number }): string {
  return `${cell.rowIndex}:${cell.columnIndex}`;
}

export function findTableByBlockId(body: NormalizedDocumentBody, tableBlockId: string): LocatedTable | null {
  for (let sectionIndex = 0; sectionIndex < body.sections.length; sectionIndex += 1) {
    const section = body.sections[sectionIndex];
    for (let blockIndex = 0; blockIndex < section.blocks.length; blockIndex += 1) {
      const block = section.blocks[blockIndex];
      if (block.type === 'table' && block.blockId === tableBlockId) {
        return {
          sectionIndex,
          blockIndex,
          table: block,
        };
      }
    }
  }

  return null;
}

function summarizeCrossings(cells: TableCellBlock[], regionStart: number, regionEndExclusive: number) {
  return cells.map((cell) => ({
    rowIndex: cell.rowIndex,
    columnIndex: cell.columnIndex,
    rowSpan: cell.rowSpan,
    colSpan: cell.colSpan,
    coversWholeRegion: cell.rowIndex < regionStart && cell.rowIndex + cell.rowSpan > regionEndExclusive,
    textPreview: cell.text,
  }));
}

export function analyzeTableRepeatRegion(
  body: NormalizedDocumentBody,
  instruction: Pick<TableRepeatInstruction, 'tableBlockId' | 'templateRowIndex' | 'templateEndRowIndex'>
): { located: LocatedTable; analysis: TableRepeatRegionAnalysis } | null {
  const located = findTableByBlockId(body, instruction.tableBlockId);
  if (!located) {
    return null;
  }

  const templateEndRowIndex = instruction.templateEndRowIndex ?? instruction.templateRowIndex;
  const regionStart = instruction.templateRowIndex;
  const regionEndExclusive = templateEndRowIndex + 1;

  const topCrossings = located.table.cells.filter(
    (cell) => cell.rowIndex < regionStart && cell.rowIndex + cell.rowSpan > regionStart
  );
  const bottomCrossings = located.table.cells.filter(
    (cell) => cell.rowIndex < regionEndExclusive && cell.rowIndex + cell.rowSpan > regionEndExclusive
  );
  const strictSafe = topCrossings.length === 0 && bottomCrossings.length === 0;
  const supportedBoundaryPolicies = strictSafe
    ? ['reject', 'split_boundary_merges'] as const
    : ['split_boundary_merges'] as const;
  const internalMergedCellCount = located.table.cells.filter(
    (cell) =>
      cell.rowIndex >= regionStart
      && cell.rowIndex + cell.rowSpan <= regionEndExclusive
      && (cell.rowSpan > 1 || cell.colSpan > 1)
  ).length;

  return {
    located,
    analysis: {
      tableBlockId: instruction.tableBlockId,
      templateRowIndex: instruction.templateRowIndex,
      templateEndRowIndex,
      templateRowCount: templateEndRowIndex - instruction.templateRowIndex + 1,
      tableRowCount: located.table.rowCount,
      strictSafe,
      splitSafe: true,
      supportedBoundaryPolicies: [...supportedBoundaryPolicies],
      recommendedBoundaryPolicy: strictSafe ? 'reject' : 'split_boundary_merges',
      topBoundaryCrossings: summarizeCrossings(topCrossings, regionStart, regionEndExclusive),
      bottomBoundaryCrossings: summarizeCrossings(bottomCrossings, regionStart, regionEndExclusive),
      internalMergedCellCount,
    },
  };
}

export function hasCoveringBoundaryCrossings(analysis: TableRepeatRegionAnalysis): boolean {
  const bottomKeys = new Set(analysis.bottomBoundaryCrossings.map(crossingKey));
  return analysis.topBoundaryCrossings.some((cell) => bottomKeys.has(crossingKey(cell)));
}
