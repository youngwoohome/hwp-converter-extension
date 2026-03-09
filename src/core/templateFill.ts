import type {
  ApiWarning,
  NormalizedDocumentBody,
  SupplementalTextBlock,
  SupplementalTextKind,
  TableCellBlock,
  TableRepeatInstruction,
  TableRepeatMatch,
  TemplateFillInstruction,
  TemplateFillRequirements,
  TemplateFillMatch,
  TemplateFillResult,
  TextMutationOperation,
} from '../types.js';
import { analyzeTableRepeatRegion, findTableByBlockId } from './tableRepeatAnalysis.js';

type PlannedTextTarget = {
  key: string;
  kind?: SupplementalTextKind;
  paragraphIndex?: number;
  sectionIndex: number;
  blockIndex: number;
  rowIndex?: number;
  columnIndex?: number;
  originalText: string;
  text: string;
};

type InstructionApplyResult = {
  matchCount: number;
  warnings: ApiWarning[];
};

function textKey(sectionIndex: number, blockIndex: number): string {
  return `p:${sectionIndex}:${blockIndex}`;
}

function cellKey(sectionIndex: number, blockIndex: number, rowIndex: number, columnIndex: number): string {
  return `c:${sectionIndex}:${blockIndex}:${rowIndex}:${columnIndex}`;
}

function supplementalKey(kind: SupplementalTextKind, paragraphIndex: number): string {
  return `s:${kind}:${paragraphIndex}`;
}

function replaceAllToken(text: string, token: string, value: string): string {
  return text.split(token).join(value);
}

function ensureParagraphPlan(
  plans: Map<string, PlannedTextTarget>,
  sectionIndex: number,
  blockIndex: number,
  originalText: string
): PlannedTextTarget {
  const key = textKey(sectionIndex, blockIndex);
  const existing = plans.get(key);
  if (existing) {
    return existing;
  }

  const created: PlannedTextTarget = {
    key,
    sectionIndex,
    blockIndex,
    originalText,
    text: originalText,
  };
  plans.set(key, created);
  return created;
}

function ensureCellPlan(
  plans: Map<string, PlannedTextTarget>,
  sectionIndex: number,
  blockIndex: number,
  rowIndex: number,
  columnIndex: number,
  originalText: string
): PlannedTextTarget {
  const key = cellKey(sectionIndex, blockIndex, rowIndex, columnIndex);
  const existing = plans.get(key);
  if (existing) {
    return existing;
  }

  const created: PlannedTextTarget = {
    key,
    sectionIndex,
    blockIndex,
    rowIndex,
    columnIndex,
    originalText,
    text: originalText,
  };
  plans.set(key, created);
  return created;
}

function ensureSupplementalPlan(
  plans: Map<string, PlannedTextTarget>,
  kind: SupplementalTextKind,
  paragraphIndex: number,
  originalText: string
): PlannedTextTarget {
  const key = supplementalKey(kind, paragraphIndex);
  const existing = plans.get(key);
  if (existing) {
    return existing;
  }

  const created: PlannedTextTarget = {
    key,
    kind,
    paragraphIndex,
    sectionIndex: -1,
    blockIndex: -1,
    originalText,
    text: originalText,
  };
  plans.set(key, created);
  return created;
}

function sortRowCells(cells: TableCellBlock[], rowIndex: number): TableCellBlock[] {
  return cells
    .filter((cell) => cell.rowIndex === rowIndex)
    .sort((left, right) => left.columnIndex - right.columnIndex);
}

function resolveRepeatedCellValue(
  rowData: string[] | Record<string, string>,
  cell: TableCellBlock,
  templateText: string,
  orderedCellIndex: number
): string {
  if (Array.isArray(rowData)) {
    return rowData[orderedCellIndex] ?? '';
  }

  if (Object.prototype.hasOwnProperty.call(rowData, templateText)) {
    return rowData[templateText] ?? '';
  }
  if (Object.prototype.hasOwnProperty.call(rowData, String(cell.columnIndex))) {
    return rowData[String(cell.columnIndex)] ?? '';
  }
  return '';
}

function normalizeTableRepeatInstructions(input: unknown): TableRepeatInstruction[] {
  if (input === undefined || input === null) {
    return [];
  }
  if (!Array.isArray(input)) {
    throw new Error('tableRepeats must be an array.');
  }

  return input.map((entry, index) => {
    if (!entry || typeof entry !== 'object') {
      throw new Error(`tableRepeats[${index}] must be an object.`);
    }
    const candidate = entry as Record<string, unknown>;
    if (typeof candidate.tableBlockId !== 'string' || candidate.tableBlockId.trim().length === 0) {
      throw new Error(`tableRepeats[${index}].tableBlockId must be a string.`);
    }
    if (!Number.isInteger(candidate.templateRowIndex)) {
      throw new Error(`tableRepeats[${index}].templateRowIndex must be an integer.`);
    }
    if (
      candidate.templateEndRowIndex !== undefined
      && !Number.isInteger(candidate.templateEndRowIndex)
    ) {
      throw new Error(`tableRepeats[${index}].templateEndRowIndex must be an integer when provided.`);
    }
    if (!Array.isArray(candidate.rows)) {
      throw new Error(`tableRepeats[${index}].rows must be an array.`);
    }
    if (
      candidate.boundaryPolicy !== undefined
      && candidate.boundaryPolicy !== 'reject'
      && candidate.boundaryPolicy !== 'split_boundary_merges'
    ) {
      throw new Error(`tableRepeats[${index}].boundaryPolicy must be "reject" or "split_boundary_merges".`);
    }
    return {
      tableBlockId: candidate.tableBlockId,
      templateRowIndex: candidate.templateRowIndex as number,
      templateEndRowIndex: candidate.templateEndRowIndex as number | undefined,
      rows: candidate.rows as Array<string[] | Record<string, string>>,
      boundaryPolicy: (candidate.boundaryPolicy as TableRepeatInstruction['boundaryPolicy']) ?? 'reject',
    };
  });
}

function normalizeStringList(input: unknown, fieldName: string): string[] {
  if (input === undefined || input === null) {
    return [];
  }
  if (!Array.isArray(input)) {
    throw new Error(`${fieldName} must be an array.`);
  }

  return input.map((entry, index) => {
    if (typeof entry !== 'string' || entry.trim().length === 0) {
      throw new Error(`${fieldName}[${index}] must be a non-empty string.`);
    }
    return entry.trim();
  });
}

export function normalizeTemplateFillInstructions(
  input: { fills?: unknown; values?: unknown; tableRepeats?: unknown; requiredPlaceholders?: unknown; requiredTargets?: unknown }
): { fills: TemplateFillInstruction[]; tableRepeats: TableRepeatInstruction[]; requiredPlaceholders: string[]; requiredTargets: string[] } {
  const instructions: TemplateFillInstruction[] = [];

  if (Array.isArray(input.fills)) {
    input.fills.forEach((entry, index) => {
      if (!entry || typeof entry !== 'object') {
        throw new Error(`fills[${index}] must be an object.`);
      }
      const candidate = entry as Record<string, unknown>;
      if (typeof candidate.value !== 'string') {
        throw new Error(`fills[${index}].value must be a string.`);
      }
      if (typeof candidate.target !== 'string' && typeof candidate.placeholder !== 'string') {
        throw new Error(`fills[${index}] must include target or placeholder.`);
      }
      instructions.push({
        target: typeof candidate.target === 'string' ? candidate.target : undefined,
        placeholder: typeof candidate.placeholder === 'string' ? candidate.placeholder : undefined,
        value: candidate.value,
        allowMultipleMatches: candidate.allowMultipleMatches === true,
      });
    });
  }

  if (input.values && typeof input.values === 'object' && !Array.isArray(input.values)) {
    Object.entries(input.values as Record<string, unknown>).forEach(([placeholder, value]) => {
      if (typeof value !== 'string') {
        throw new Error(`values[${placeholder}] must be a string.`);
      }
      instructions.push({
        placeholder,
        value,
        allowMultipleMatches: false,
      });
    });
  }

  return {
    fills: instructions,
    tableRepeats: normalizeTableRepeatInstructions(input.tableRepeats),
    requiredPlaceholders: normalizeStringList(input.requiredPlaceholders, 'requiredPlaceholders'),
    requiredTargets: normalizeStringList(input.requiredTargets, 'requiredTargets'),
  };
}

function countExactTargetMatches(body: NormalizedDocumentBody, target: string): number {
  let matches = 0;

  body.sections.forEach((section) => {
    section.blocks.forEach((block) => {
      if (block.type === 'paragraph' && block.blockId === target) {
        matches += 1;
        return;
      }

      if (block.type === 'table') {
        block.cells.forEach((cell) => {
          if (cell.cellId === target) {
            matches += 1;
          }
        });
      }
    });
  });

  return matches;
}

function countExactTargetMatchesInSupplemental(supplementalText: SupplementalTextBlock[], target: string): number {
  return supplementalText.filter((block) => block.targetId === target && block.editable).length;
}

function countPlaceholderTargetMatches(body: NormalizedDocumentBody, supplementalText: SupplementalTextBlock[], placeholder: string): number {
  const matches = new Set<string>();

  body.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (block.type === 'paragraph' && block.text.includes(placeholder)) {
        matches.add(textKey(sectionIndex, blockIndex));
        return;
      }

      if (block.type === 'table') {
        block.cells.forEach((cell) => {
          if (cell.text.includes(placeholder)) {
            matches.add(cellKey(sectionIndex, blockIndex, cell.rowIndex, cell.columnIndex));
          }
        });
      }
    });
  });

  supplementalText.forEach((block) => {
    if (!block.editable || !block.text.includes(placeholder)) {
      return;
    }
    matches.add(supplementalKey(block.kind, block.paragraphIndex));
  });

  return matches.size;
}

export function validateTemplateFillRequirements(
  body: NormalizedDocumentBody,
  supplementalText: SupplementalTextBlock[],
  requirements: TemplateFillRequirements
): ApiWarning[] {
  const issues: ApiWarning[] = [];

  requirements.requiredPlaceholders.forEach((placeholder) => {
    const matchCount = countPlaceholderTargetMatches(body, supplementalText, placeholder);
    if (matchCount === 0) {
      issues.push({
        code: 'TEMPLATE_REQUIRED_PLACEHOLDER_MISSING',
        message: `Required placeholder "${placeholder}" was not found in the current document.`,
      });
      return;
    }
    if (matchCount > 1) {
      issues.push({
        code: 'TEMPLATE_REQUIRED_PLACEHOLDER_AMBIGUOUS',
        message: `Required placeholder "${placeholder}" matched ${matchCount} targets. Use explicit target ids or unique tokens before filling.`,
      });
    }
  });

  requirements.requiredTargets.forEach((target) => {
    const matchCount = countExactTargetMatches(body, target) + countExactTargetMatchesInSupplemental(supplementalText, target);
    if (matchCount === 0) {
      issues.push({
        code: 'TEMPLATE_REQUIRED_TARGET_MISSING',
        message: `Required target "${target}" was not found in the current document.`,
      });
      return;
    }
    if (matchCount > 1) {
      issues.push({
        code: 'TEMPLATE_REQUIRED_TARGET_DUPLICATE',
        message: `Required target "${target}" matched ${matchCount} document nodes and is not safe to fill automatically.`,
      });
    }
  });

  return issues;
}

function applyTargetInstruction(
  body: NormalizedDocumentBody,
  supplementalText: SupplementalTextBlock[],
  plans: Map<string, PlannedTextTarget>,
  instruction: TemplateFillInstruction
): InstructionApplyResult {
  const target = instruction.target ?? '';
  let matches = 0;

  body.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (block.type === 'paragraph' && block.blockId === target) {
        const plan = ensureParagraphPlan(plans, sectionIndex, blockIndex, block.text);
        plan.text = instruction.value;
        matches += 1;
        return;
      }

      if (block.type === 'table') {
        block.cells.forEach((cell) => {
          if (cell.cellId !== target) {
            return;
          }
          const plan = ensureCellPlan(plans, sectionIndex, blockIndex, cell.rowIndex, cell.columnIndex, cell.text);
          plan.text = instruction.value;
          matches += 1;
        });
      }
    });
  });

  supplementalText.forEach((block) => {
    if (!block.editable || block.targetId !== target) {
      return;
    }
    const plan = ensureSupplementalPlan(plans, block.kind, block.paragraphIndex, block.text);
    plan.text = instruction.value;
    matches += 1;
  });

  return {
    matchCount: matches,
    warnings: [],
  };
}

function applyPlaceholderInstruction(
  body: NormalizedDocumentBody,
  supplementalText: SupplementalTextBlock[],
  plans: Map<string, PlannedTextTarget>,
  instruction: TemplateFillInstruction
): InstructionApplyResult {
  const placeholder = (instruction.placeholder ?? '').trim();
  if (placeholder.length === 0) {
    return {
      matchCount: 0,
      warnings: [],
    };
  }

  const candidatePlans: PlannedTextTarget[] = [];

  body.sections.forEach((section, sectionIndex) => {
    section.blocks.forEach((block, blockIndex) => {
      if (block.type === 'paragraph') {
        const plan = ensureParagraphPlan(plans, sectionIndex, blockIndex, block.text);
        if (plan.text.includes(placeholder)) {
          candidatePlans.push(plan);
        }
        return;
      }

      if (block.type === 'table') {
        block.cells.forEach((cell) => {
          const plan = ensureCellPlan(plans, sectionIndex, blockIndex, cell.rowIndex, cell.columnIndex, cell.text);
          if (!plan.text.includes(placeholder)) {
            return;
          }
          candidatePlans.push(plan);
        });
      }
    });
  });

  supplementalText.forEach((block) => {
    if (!block.editable) {
      return;
    }
    const plan = ensureSupplementalPlan(plans, block.kind, block.paragraphIndex, block.text);
    if (plan.text.includes(placeholder)) {
      candidatePlans.push(plan);
    }
  });

  if (candidatePlans.length === 0) {
    return {
      matchCount: 0,
      warnings: [],
    };
  }

  const uniqueTargets = new Map<string, PlannedTextTarget>();
  candidatePlans.forEach((plan) => {
    uniqueTargets.set(plan.key, plan);
  });

  if (uniqueTargets.size > 1 && instruction.allowMultipleMatches !== true) {
    return {
      matchCount: 0,
      warnings: [
        {
          code: 'TEMPLATE_FILL_AMBIGUOUS_TOKEN',
          message: `Placeholder "${placeholder}" matched ${uniqueTargets.size} targets. Use fills[].allowMultipleMatches=true or explicit target ids to apply deterministically.`,
        },
      ],
    };
  }

  uniqueTargets.forEach((plan) => {
    plan.text = replaceAllToken(plan.text, placeholder, instruction.value);
  });

  return {
    matchCount: uniqueTargets.size,
    warnings: [],
  };
}

function convertPlansToMutations(plans: Map<string, PlannedTextTarget>): TextMutationOperation[] {
  const paragraphMutations: TextMutationOperation[] = [];
  const cellMutations: TextMutationOperation[] = [];

  for (const plan of plans.values()) {
    if (plan.text === plan.originalText) {
      continue;
    }

    if (typeof plan.rowIndex === 'number' && typeof plan.columnIndex === 'number') {
      cellMutations.push({
        op: 'replace_table_cell_text',
        sectionIndex: plan.sectionIndex,
        blockIndex: plan.blockIndex,
        rowIndex: plan.rowIndex,
        columnIndex: plan.columnIndex,
        text: plan.text,
      });
      continue;
    }

    if (plan.kind && typeof plan.paragraphIndex === 'number') {
      paragraphMutations.push({
        op: 'replace_text_in_supplemental_paragraph',
        kind: plan.kind,
        paragraphIndex: plan.paragraphIndex,
        text: plan.text,
      });
      continue;
    }

    paragraphMutations.push({
      op: 'replace_text_in_paragraph',
      sectionIndex: plan.sectionIndex,
      blockIndex: plan.blockIndex,
      text: plan.text,
    });
  }

  return [...paragraphMutations, ...cellMutations];
}

function planTableRepeats(
  body: NormalizedDocumentBody,
  tableRepeats: TableRepeatInstruction[]
): { mutations: TextMutationOperation[]; matches: TableRepeatMatch[]; warnings: ApiWarning[] } {
  const mutations: TextMutationOperation[] = [];
  const matches: TableRepeatMatch[] = [];
  const warnings: ApiWarning[] = [];

  tableRepeats.forEach((instruction) => {
    const analyzed = analyzeTableRepeatRegion(body, instruction);
    if (!analyzed) {
      warnings.push({
        code: 'TEMPLATE_REPEAT_NO_TABLE',
        message: `No table matched blockId "${instruction.tableBlockId}".`,
      });
      return;
    }
    const { located, analysis } = analyzed;

    const templateEndRowIndex = analysis.templateEndRowIndex;
    if (templateEndRowIndex < instruction.templateRowIndex) {
      warnings.push({
        code: 'TEMPLATE_REPEAT_INVALID_REGION',
        message: `templateEndRowIndex must be greater than or equal to templateRowIndex for table "${instruction.tableBlockId}".`,
      });
      return;
    }

    const templateRowCount = templateEndRowIndex - instruction.templateRowIndex + 1;
    const templateRows = Array.from({ length: templateRowCount }, (_, offset) => instruction.templateRowIndex + offset);
    const templateCellsByRow = templateRows.map((rowIndex) => sortRowCells(located.table.cells, rowIndex));
    if (templateCellsByRow.some((row) => row.length === 0)) {
      warnings.push({
        code: 'TEMPLATE_REPEAT_NO_ROW',
        message: `A row in template region ${instruction.templateRowIndex}-${templateEndRowIndex} does not exist in table "${instruction.tableBlockId}".`,
      });
      return;
    }

    if (!analysis.strictSafe && analysis.topBoundaryCrossings.length > 0) {
      if (instruction.boundaryPolicy === 'split_boundary_merges') {
        warnings.push({
          code: 'TEMPLATE_REPEAT_BOUNDARY_SPLIT_APPLIED',
          message: `Template region ${instruction.templateRowIndex}-${templateEndRowIndex} in table "${instruction.tableBlockId}" crosses a top merge boundary. The engine will normalize crossing merged cells before repeating the region.`,
        });
      } else {
        warnings.push({
          code: 'TEMPLATE_REPEAT_MERGED_REGION_UNSUPPORTED',
          message: `Template region ${instruction.templateRowIndex}-${templateEndRowIndex} in table "${instruction.tableBlockId}" is covered by a merged cell anchored above the selected repeat region.`,
        });
        return;
      }
    }

    if (!analysis.strictSafe && analysis.bottomBoundaryCrossings.length > 0) {
      if (instruction.boundaryPolicy === 'split_boundary_merges') {
        warnings.push({
          code: 'TEMPLATE_REPEAT_BOUNDARY_SPLIT_APPLIED',
          message: `Template region ${instruction.templateRowIndex}-${templateEndRowIndex} in table "${instruction.tableBlockId}" crosses a bottom merge boundary. The engine will normalize crossing merged cells before repeating the region.`,
        });
      } else {
        warnings.push({
          code: 'TEMPLATE_REPEAT_MERGED_REGION_UNSUPPORTED',
          message: `Template region ${instruction.templateRowIndex}-${templateEndRowIndex} in table "${instruction.tableBlockId}" crosses a vertical merge boundary and cannot be repeated safely yet.`,
        });
        return;
      }
    }

    if (instruction.rows.length === 0) {
      warnings.push({
        code: 'TEMPLATE_REPEAT_EMPTY',
        message: `No rows were provided for table "${instruction.tableBlockId}".`,
      });
      return;
    }

    if (instruction.rows.length % templateRowCount !== 0) {
      warnings.push({
        code: 'TEMPLATE_REPEAT_ROW_COUNT_MISMATCH',
        message: `tableRepeats rows for table "${instruction.tableBlockId}" must be a multiple of template region row count ${templateRowCount}.`,
      });
      return;
    }

    const repeatedRegionCount = Math.floor(instruction.rows.length / templateRowCount);
    const extraRegions = Math.max(repeatedRegionCount - 1, 0);
    for (let regionOffset = 0; regionOffset < extraRegions; regionOffset += 1) {
      mutations.push({
        op: 'clone_table_region',
        sectionIndex: located.sectionIndex,
        blockIndex: located.blockIndex,
        templateStartRowIndex: instruction.templateRowIndex,
        templateEndRowIndex,
        insertAfterRowIndex: templateEndRowIndex + regionOffset * templateRowCount,
        boundaryPolicy: instruction.boundaryPolicy ?? 'reject',
      });
    }

    instruction.rows.forEach((rowData, rowOffset) => {
      const targetRowIndex = instruction.templateRowIndex + rowOffset;
      const templateRowOffset = rowOffset % templateRowCount;
      const templateCells = templateCellsByRow[templateRowOffset] ?? [];
      templateCells.forEach((cell, cellIndex) => {
        const nextText = resolveRepeatedCellValue(rowData, cell, cell.text, cellIndex);
        mutations.push({
          op: 'replace_table_cell_text',
          sectionIndex: located.sectionIndex,
          blockIndex: located.blockIndex,
          rowIndex: targetRowIndex,
          columnIndex: cell.columnIndex,
          text: nextText,
        });
      });
    });

    matches.push({
      tableBlockId: instruction.tableBlockId,
      templateRowIndex: instruction.templateRowIndex,
      templateEndRowIndex,
      templateRowCount: analysis.templateRowCount,
      repeatedRegionCount,
      appliedRowCount: instruction.rows.length,
      insertedRowCount: extraRegions * templateRowCount,
    });
  });

  return { mutations, matches, warnings };
}

export function planTemplateFill(
  body: NormalizedDocumentBody,
  supplementalText: SupplementalTextBlock[],
  input: { fills: TemplateFillInstruction[]; tableRepeats: TableRepeatInstruction[] }
): { mutations: TextMutationOperation[]; result: TemplateFillResult } {
  const plans = new Map<string, PlannedTextTarget>();
  const applied: TemplateFillMatch[] = [];
  const warnings: ApiWarning[] = [];

  input.fills.forEach((instruction, instructionIndex) => {
    const applyResult = instruction.target
      ? applyTargetInstruction(body, supplementalText, plans, instruction)
      : applyPlaceholderInstruction(body, supplementalText, plans, instruction);

    warnings.push(...applyResult.warnings);

    if (applyResult.matchCount === 0) {
      if (applyResult.warnings.length > 0) {
        return;
      }
      warnings.push({
        code: 'TEMPLATE_FILL_NO_MATCH',
        message: instruction.target
          ? `No block or table cell matched target "${instruction.target}".`
          : `No placeholder matched "${instruction.placeholder}".`,
      });
      return;
    }

    applied.push({
      instructionIndex,
      target: instruction.target ?? instruction.placeholder ?? '',
      matchCount: applyResult.matchCount,
    });
  });

  const directFillMutations = convertPlansToMutations(plans);
  const repeated = planTableRepeats(body, input.tableRepeats);
  warnings.push(...repeated.warnings);

  return {
    mutations: [...repeated.mutations, ...directFillMutations],
    result: {
      applied,
      repeatedTables: repeated.matches,
      warnings,
      mutationCount: repeated.mutations.length + directFillMutations.length,
    },
  };
}
