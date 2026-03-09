import { basename, dirname, extname, join } from 'node:path';
import type { ApiWarning, ConvertContext, ConvertedArtifact, EmbeddedImageReference, MarkdownMode, TableBlock, TableCellBlock } from '../types.js';
import { ensureParentDirectory, writeFileAtomically } from '../utils/files.js';
import { resolveOutputPath } from '../converters/common.js';
import { parseStructuredHwpx } from './structuredHwpx.js';

type MarkdownExportOptions = {
  sourceFormat: 'hwp' | 'hwpx';
  mode: MarkdownMode;
  includeDiagnostics: boolean;
};

type SimplificationRecord = {
  blockId: string;
  reason: string;
};

function escapeMarkdown(text: string): string {
  return text
    .replaceAll('\\', '\\\\')
    .replaceAll('|', '\\|')
    .replaceAll('\n', '<br />');
}

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function sortCells(cells: TableCellBlock[]): TableCellBlock[] {
  return [...cells].sort((left, right) => {
    if (left.rowIndex !== right.rowIndex) return left.rowIndex - right.rowIndex;
    return left.columnIndex - right.columnIndex;
  });
}

function looksLikeImage(fileName: string): boolean {
  return /\.(png|jpe?g|gif|bmp|svg|webp|tiff?)$/i.test(fileName);
}

async function extractPackageAssets(
  parsed: Awaited<ReturnType<typeof parseStructuredHwpx>>,
  markdownPath: string
): Promise<{ assetDir: string | null; assetFiles: string[]; assetPathMap: Map<string, string> }> {
  if (parsed.packageValidation.assetEntries.length === 0) {
    return { assetDir: null, assetFiles: [], assetPathMap: new Map() };
  }

  const assetDir = join(dirname(markdownPath), `${basename(markdownPath, extname(markdownPath))}.assets`);
  await ensureParentDirectory(join(assetDir, '.keep'));

  const seen = new Set<string>();
  const assetFiles: string[] = [];
  const assetPathMap = new Map<string, string>();

  for (const entry of parsed.packageValidation.assetEntries) {
    const file = parsed.zip.file(entry);
    if (!file) continue;

    const originalName = basename(entry);
    let candidateName = originalName;
    let counter = 1;
    while (seen.has(candidateName)) {
      const ext = extname(originalName);
      const stem = basename(originalName, ext);
      candidateName = `${stem}-${counter}${ext}`;
      counter += 1;
    }
    seen.add(candidateName);

    const targetPath = join(assetDir, candidateName);
    const buffer = await file.async('nodebuffer');
    await writeFileAtomically(targetPath, buffer);
    assetFiles.push(targetPath);
    assetPathMap.set(entry, targetPath);
  }

  return { assetDir, assetFiles, assetPathMap };
}

function relativeAssetPath(assetDir: string | null, assetPath: string): string {
  const relativeAssetDir = assetDir ? basename(assetDir) : 'assets';
  return `${relativeAssetDir}/${basename(assetPath)}`;
}

function imageReferencesForTarget(imageReferences: EmbeddedImageReference[], targetId: string): EmbeddedImageReference[] {
  return imageReferences.filter((image) => image.targetId === targetId);
}

function renderMarkdownImageReference(
  image: EmbeddedImageReference,
  assetPathMap: Map<string, string>,
  assetDir: string | null
): string | null {
  if (!image.assetPath) {
    return null;
  }
  const outputPath = assetPathMap.get(image.assetPath);
  if (!outputPath) {
    return null;
  }
  const label = image.assetFileName || image.binaryItemId;
  return `![${escapeMarkdown(label)}](${relativeAssetPath(assetDir, outputPath)})`;
}

function renderHtmlImageReference(
  image: EmbeddedImageReference,
  assetPathMap: Map<string, string>,
  assetDir: string | null
): string | null {
  if (!image.assetPath) {
    return null;
  }
  const outputPath = assetPathMap.get(image.assetPath);
  if (!outputPath) {
    return null;
  }
  const relativePath = relativeAssetPath(assetDir, outputPath);
  const alt = escapeHtml(image.assetFileName || image.binaryItemId);
  return `<img src="${escapeHtml(relativePath)}" alt="${alt}" />`;
}

function renderCellMarkdown(
  cell: TableCellBlock,
  imageReferences: EmbeddedImageReference[],
  assetPathMap: Map<string, string>,
  assetDir: string | null
): string {
  const base = escapeMarkdown(cell.text);
  const images = imageReferencesForTarget(imageReferences, cell.cellId)
    .map((image) => renderMarkdownImageReference(image, assetPathMap, assetDir))
    .filter((entry): entry is string => entry !== null);
  return [base, ...images].filter((entry) => entry.length > 0).join('<br />');
}

function renderCellHtml(
  cell: TableCellBlock,
  imageReferences: EmbeddedImageReference[],
  assetPathMap: Map<string, string>,
  assetDir: string | null
): string {
  const base = escapeHtml(cell.text).replaceAll('\n', '<br />');
  const images = imageReferencesForTarget(imageReferences, cell.cellId)
    .map((image) => renderHtmlImageReference(image, assetPathMap, assetDir))
    .filter((entry): entry is string => entry !== null);
  return [base, ...images].filter((entry) => entry.length > 0).join('<br />');
}

function buildSimpleTableGrid(
  table: TableBlock,
  imageReferences: EmbeddedImageReference[],
  assetPathMap: Map<string, string>,
  assetDir: string | null
): string[][] | null {
  if (table.cells.some((cell) => cell.rowSpan !== 1 || cell.colSpan !== 1)) {
    return null;
  }

  const columnCount = Math.max(table.columnCount, ...table.cells.map((cell) => cell.columnIndex + 1));
  const rowCount = Math.max(table.rowCount, ...table.cells.map((cell) => cell.rowIndex + 1));
  const grid = Array.from({ length: rowCount }, () => Array.from({ length: columnCount }, () => ''));

  for (const cell of table.cells) {
    if (!grid[cell.rowIndex]?.[cell.columnIndex] && grid[cell.rowIndex]) {
      grid[cell.rowIndex][cell.columnIndex] = renderCellMarkdown(cell, imageReferences, assetPathMap, assetDir);
    }
  }

  return grid;
}

function renderGfmTable(
  table: TableBlock,
  imageReferences: EmbeddedImageReference[],
  assetPathMap: Map<string, string>,
  assetDir: string | null
): string | null {
  const grid = buildSimpleTableGrid(table, imageReferences, assetPathMap, assetDir);
  if (!grid || grid.length === 0) {
    return null;
  }

  const columnCount = grid[0]?.length ?? 0;
  if (columnCount === 0) {
    return null;
  }

  const blankHeader = `| ${Array.from({ length: columnCount }, () => '').join(' | ')} |`;
  const separator = `| ${Array.from({ length: columnCount }, () => '---').join(' | ')} |`;
  const rows = grid.map((row) => `| ${row.join(' | ')} |`);
  return [blankHeader, separator, ...rows].join('\n');
}

function renderHtmlTable(
  table: TableBlock,
  imageReferences: EmbeddedImageReference[],
  assetPathMap: Map<string, string>,
  assetDir: string | null
): string {
  const rows = new Map<number, TableCellBlock[]>();
  for (const cell of sortCells(table.cells)) {
    const row = rows.get(cell.rowIndex) ?? [];
    row.push(cell);
    rows.set(cell.rowIndex, row);
  }

  const body = [...rows.entries()]
    .sort((left, right) => left[0] - right[0])
    .map(([, cells]) => {
      const cellsHtml = cells.map((cell) => {
        const attrs = [
          cell.rowSpan > 1 ? ` rowspan="${cell.rowSpan}"` : '',
          cell.colSpan > 1 ? ` colspan="${cell.colSpan}"` : '',
        ].join('');
        return `<td${attrs}>${renderCellHtml(cell, imageReferences, assetPathMap, assetDir)}</td>`;
      }).join('');
      return `<tr>${cellsHtml}</tr>`;
    })
    .join('');

  return `<table>\n<tbody>\n${body}\n</tbody>\n</table>`;
}

function buildMarkdownDocument(
  parsed: Awaited<ReturnType<typeof parseStructuredHwpx>>,
  options: MarkdownExportOptions,
  extractedAssets: string[],
  assetDir: string | null,
  assetPathMap: Map<string, string>
): { markdown: string; diagnostics: Record<string, unknown>; warnings: ApiWarning[] } {
  const chunks: string[] = [];
  const warnings: ApiWarning[] = [...parsed.warnings];
  const simplifiedBlocks: SimplificationRecord[] = [];
  const inlineRenderedAssetPaths = new Set<string>();

  if (parsed.headerTexts.length > 0) {
    warnings.push({
      code: 'MARKDOWN_HEADER_APPENDIX',
      message: 'Header content was preserved as a Markdown header appendix because exact page placement is not reconstructed yet.',
    });
    chunks.push('## Header');
    chunks.push(parsed.headerTexts.join('\n\n'));
    const headerImages = parsed.imageReferences
      .filter((image) => image.kind === 'header_paragraph')
      .map((image) => {
        if (image.assetPath) {
          inlineRenderedAssetPaths.add(image.assetPath);
        }
        return renderMarkdownImageReference(image, assetPathMap, assetDir);
      })
      .filter((entry): entry is string => entry !== null);
    if (headerImages.length > 0) {
      warnings.push({
        code: 'MARKDOWN_INLINE_IMAGE_APPROXIMATED',
        message: 'Image placement was approximated near the original paragraph flow in Markdown export.',
      });
      chunks.push(headerImages.join('\n\n'));
    }
  }

  for (const section of parsed.body.sections) {
    for (const block of section.blocks) {
      if (block.type === 'paragraph') {
        if (block.text.trim().length > 0) {
          chunks.push(block.text.trim());
        }
        const paragraphImages = imageReferencesForTarget(parsed.imageReferences, block.blockId)
          .map((image) => {
            if (image.assetPath) {
              inlineRenderedAssetPaths.add(image.assetPath);
            }
            return renderMarkdownImageReference(image, assetPathMap, assetDir);
          })
          .filter((entry): entry is string => entry !== null);
        if (paragraphImages.length > 0) {
          if (!warnings.some((warning) => warning.code === 'MARKDOWN_INLINE_IMAGE_APPROXIMATED')) {
            warnings.push({
              code: 'MARKDOWN_INLINE_IMAGE_APPROXIMATED',
              message: 'Image placement was approximated near the original paragraph flow in Markdown export.',
            });
          }
          chunks.push(paragraphImages.join('\n\n'));
        }
        if (block.containsObjects) {
          simplifiedBlocks.push({
            blockId: block.blockId,
            reason: 'Paragraph contains preserved objects that cannot be positioned in Markdown output.',
          });
        }
        continue;
      }

      const tableImages = parsed.imageReferences.filter((image) => image.containerId === block.blockId);
      if (tableImages.length > 0 && !warnings.some((warning) => warning.code === 'MARKDOWN_INLINE_IMAGE_APPROXIMATED')) {
        warnings.push({
          code: 'MARKDOWN_INLINE_IMAGE_APPROXIMATED',
          message: 'Image placement was approximated near the original paragraph flow in Markdown export.',
        });
      }
      tableImages.forEach((image) => {
        if (image.assetPath) {
          inlineRenderedAssetPaths.add(image.assetPath);
        }
      });

      const gfmTable = renderGfmTable(block, parsed.imageReferences, assetPathMap, assetDir);
      if (options.mode === 'clean' && gfmTable) {
        chunks.push(gfmTable);
      } else {
        chunks.push(renderHtmlTable(block, parsed.imageReferences, assetPathMap, assetDir));
        if (!gfmTable) {
          simplifiedBlocks.push({
            blockId: block.blockId,
            reason: 'Rendered as HTML table because merged cells or irregular geometry cannot be represented as GFM.',
          });
        }
      }
    }
  }

  if (parsed.footerTexts.length > 0) {
    warnings.push({
      code: 'MARKDOWN_FOOTER_APPENDIX',
      message: 'Footer content was preserved as a Markdown footer appendix because exact page placement is not reconstructed yet.',
    });
    chunks.push('## Footer');
    chunks.push(parsed.footerTexts.join('\n\n'));
    const footerImages = parsed.imageReferences
      .filter((image) => image.kind === 'footer_paragraph')
      .map((image) => {
        if (image.assetPath) {
          inlineRenderedAssetPaths.add(image.assetPath);
        }
        return renderMarkdownImageReference(image, assetPathMap, assetDir);
      })
      .filter((entry): entry is string => entry !== null);
    if (footerImages.length > 0) {
      if (!warnings.some((warning) => warning.code === 'MARKDOWN_INLINE_IMAGE_APPROXIMATED')) {
        warnings.push({
          code: 'MARKDOWN_INLINE_IMAGE_APPROXIMATED',
          message: 'Image placement was approximated near the original paragraph flow in Markdown export.',
        });
      }
      chunks.push(footerImages.join('\n\n'));
    }
  }

  if (extractedAssets.length > 0) {
    warnings.push({
      code: 'MARKDOWN_ASSET_APPENDIX',
      message: 'Embedded assets were exported beside the Markdown file and appended as an asset appendix because exact inline placement is not yet reconstructed.',
    });

    const relativeAssetDir = assetDir ? basename(assetDir) : 'assets';
    const appendix = extractedAssets
      .filter((assetPath) => {
        const originalAssetPath = [...assetPathMap.entries()].find(([, outputPath]) => outputPath === assetPath)?.[0];
        return !originalAssetPath || !inlineRenderedAssetPaths.has(originalAssetPath);
      })
      .map((assetPath) => {
      const fileName = basename(assetPath);
      const relativePath = `${relativeAssetDir}/${fileName}`;
      if (looksLikeImage(fileName)) {
        return `![${fileName}](${relativePath})`;
      }
      return `- [${fileName}](${relativePath})`;
    });

    if (appendix.length > 0) {
      chunks.push('## Extracted Assets');
      chunks.push(appendix.join('\n'));
    }
  }

  if (options.mode === 'clean' && simplifiedBlocks.length > 0) {
    warnings.push({
      code: 'MARKDOWN_SIMPLIFIED',
      message: `${simplifiedBlocks.length} block(s) were simplified for clean Markdown output.`,
    });
  }

  const diagnostics = {
    sourceFormat: options.sourceFormat,
    projectionMode: parsed.features.projectionMode,
    markdownMode: options.mode,
    structured: true,
    packageValidation: parsed.packageValidation,
    features: parsed.features,
    headerTexts: parsed.headerTexts,
    footerTexts: parsed.footerTexts,
    imageReferences: parsed.imageReferences,
    warningCodes: warnings.map((warning) => warning.code),
    simplifiedBlocks,
    extractedAssets: extractedAssets.map((assetPath) => basename(assetPath)),
  };

  return {
    markdown: `${chunks.filter((chunk) => chunk.trim().length > 0).join('\n\n')}\n`,
    diagnostics,
    warnings,
  };
}

export async function convertStructuredHwpxToMarkdown(context: ConvertContext): Promise<ConvertedArtifact> {
  const markdownPath = resolveOutputPath(context);
  const parsed = await parseStructuredHwpx(context.sourcePath);
  const mode = context.options?.markdownMode ?? 'fidelity';
  const includeDiagnostics = context.options?.includeDiagnostics ?? true;

  const { assetDir, assetFiles, assetPathMap } = await extractPackageAssets(parsed, markdownPath);
  const { markdown, diagnostics, warnings } = buildMarkdownDocument(parsed, {
    sourceFormat: context.sourceFormat,
    mode,
    includeDiagnostics,
  }, assetFiles, assetDir, assetPathMap);

  await ensureParentDirectory(markdownPath);
  await writeFileAtomically(markdownPath, Buffer.from(markdown, 'utf-8'));

  const sidecarPaths = [...assetFiles];
  if (includeDiagnostics) {
    const diagnosticsPath = join(
      dirname(markdownPath),
      `${basename(markdownPath, extname(markdownPath))}.diagnostics.json`
    );
    await writeFileAtomically(diagnosticsPath, Buffer.from(`${JSON.stringify(diagnostics, null, 2)}\n`, 'utf-8'));
    sidecarPaths.push(diagnosticsPath);
  }

  return {
    outputPath: markdownPath,
    targetFormat: 'md',
    sidecarPaths,
    warnings,
    details: {
      markdownMode: mode,
      assetDirectory: assetDir,
      diagnosticsIncluded: includeDiagnostics,
    },
  };
}
