import { basename } from 'node:path';
import { nanoid } from 'nanoid';
import { extractHwpText } from '../converters/hwp.js';
import { extractHwpxText, saveHwpxText } from '../converters/hwpx.js';
import type {
  ApiWarning,
  DocumentTemplateId,
  DocumentFeatureSummary,
  DocumentMode,
  DocumentSession,
  EmbeddedImageReference,
  NormalizedDocumentBody,
  ProjectionMode,
  SupplementalTextBlock,
  SourceFormat,
  TextExtractionResult,
  TextMutationOperation,
} from '../types.js';
import { copyStructuredHwpx, parseStructuredHwpx, saveStructuredHwpxWithMutations, saveStructuredHwpxWithPlainText } from './structuredHwpx.js';
import { getUniqueSiblingPath, normalizeHwpxOutputPath } from '../utils/files.js';
import { convertHwpToHwpxWithJvm, createBlankHwpxWithJvm, extractHwpTextWithJvm } from './jvmCore.js';
import { applyDocumentTemplate } from './hwpxTemplates.js';

export type CoreDocumentData = {
  sourceFormat: SourceFormat;
  projectionMode: ProjectionMode;
  documentMode: DocumentMode;
  readOnly: boolean;
  title: string;
  body: NormalizedDocumentBody;
  supplementalText: SupplementalTextBlock[];
  imageReferences: EmbeddedImageReference[];
  paragraphs: string[];
  rawText: string;
  warnings: ApiWarning[];
  features: DocumentFeatureSummary;
  canonicalPath?: string | null;
};

function buildTextProjectionBody(paragraphs: string[]): NormalizedDocumentBody {
  return {
    sections: [
      {
        sectionId: 'section0',
        blocks: paragraphs.map((text, index) => ({
          type: 'paragraph' as const,
          blockId: `0:${index}`,
          text,
          styleRef: 'Body',
          editable: true,
          containsObjects: false,
        })),
      },
    ],
  };
}

function buildTextProjectionFeatures(documentMode: DocumentMode): DocumentFeatureSummary {
  return {
    projectionMode: 'text_projection',
    authoritative: false,
    editable: documentMode === 'editable' ? ['paragraph'] : [],
    preservedReadOnly: [],
    unsupported: ['table', 'image', 'comment', 'header', 'footer', 'style'],
    hasUnsupportedEditableFeatures: true,
    allowPlainTextSave: true,
  };
}

function buildTextProjectionWarnings(sourceFormat: SourceFormat, documentMode: DocumentMode): ApiWarning[] {
  const warnings: ApiWarning[] = [
    {
      code: 'TEXT_PROJECTION',
      message: 'This document is open in text-projection mode. Rich structure is not preserved during editable-copy generation.',
    },
  ];

  if (documentMode !== 'editable') {
    warnings.push({
      code: 'WORKING_COPY_REQUIRED',
      message: sourceFormat === 'hwp'
        ? 'Create an editable .hwpx working copy before editing this .hwp file.'
        : 'Create an editable .hwpx working copy before editing this source document.',
    });
  }

  return warnings;
}

function createTextProjectionDocumentData(
  sourcePath: string,
  sourceFormat: SourceFormat,
  extracted: TextExtractionResult
): CoreDocumentData {
  const documentMode: DocumentMode = sourceFormat === 'hwp' ? 'import_required' : 'read_only';
  return {
    sourceFormat,
    projectionMode: 'text_projection',
    documentMode,
    readOnly: true,
    title: basename(sourcePath),
    body: buildTextProjectionBody(extracted.paragraphs),
    supplementalText: [],
    imageReferences: [],
    paragraphs: extracted.paragraphs,
    rawText: extracted.rawText,
    warnings: buildTextProjectionWarnings(sourceFormat, documentMode),
    features: buildTextProjectionFeatures(documentMode),
    canonicalPath: null,
  };
}

export async function openDocumentCore(sourcePath: string, sourceFormat: SourceFormat): Promise<CoreDocumentData> {
  if (sourceFormat === 'hwpx') {
    const parsed = await parseStructuredHwpx(sourcePath);
    return {
      sourceFormat,
      projectionMode: 'structured_hwpx',
      documentMode: 'read_only',
      readOnly: true,
      title: basename(sourcePath),
      body: parsed.body,
      supplementalText: parsed.supplementalText,
      imageReferences: parsed.imageReferences,
      paragraphs: parsed.paragraphs,
      rawText: parsed.rawText,
      warnings: parsed.warnings,
      features: parsed.features,
      canonicalPath: null,
    };
  }

  const extracted = await (async () => {
    try {
      const rawText = await extractHwpTextWithJvm(sourcePath);
      const paragraphs = rawText
        .replace(/\r\n/g, '\n')
        .split(/\n\s*\n/g)
        .map((entry) => entry.replace(/\s+/g, ' ').trim())
        .filter((entry) => entry.length > 0);
      return { rawText: rawText.trim(), paragraphs };
    } catch {
      return extractHwpText(sourcePath);
    }
  })();
  return createTextProjectionDocumentData(sourcePath, sourceFormat, extracted);
}

export async function createHwpxDocumentCore(outputPath: string, templateId: DocumentTemplateId = 'blank'): Promise<CoreDocumentData> {
  try {
    await createBlankHwpxWithJvm(outputPath);
  } catch {
    await saveHwpxText(outputPath, '', outputPath);
  }
  const templateWarnings = await applyDocumentTemplate(outputPath, templateId);
  const parsed = await parseStructuredHwpx(outputPath);

  return {
    sourceFormat: 'hwpx',
    projectionMode: parsed.features.projectionMode,
    documentMode: 'editable',
    readOnly: false,
    title: basename(outputPath),
    body: parsed.body,
    supplementalText: parsed.supplementalText,
    imageReferences: parsed.imageReferences,
    paragraphs: parsed.paragraphs,
    rawText: parsed.rawText,
    warnings: [...templateWarnings, ...parsed.warnings],
    features: {
      ...parsed.features,
      editable: parsed.features.editable.length > 0 ? parsed.features.editable : ['paragraph'],
    },
    canonicalPath: outputPath,
  };
}

export async function createHwpxDocumentFromReferenceCore(
  referencePath: string,
  referenceFormat: SourceFormat,
  outputPath: string
): Promise<CoreDocumentData> {
  let warnings: ApiWarning[];

  if (referenceFormat === 'hwpx') {
    const parsed = await copyStructuredHwpx(referencePath, outputPath);
    warnings = [
      {
        code: 'REFERENCE_TEMPLATE_APPLIED',
        message: 'Created a new editable HWPX document by copying a validated reference HWPX template.',
      },
    ];

    return {
      sourceFormat: 'hwpx',
      projectionMode: parsed.features.projectionMode,
      documentMode: 'editable',
      readOnly: false,
      title: basename(outputPath),
      body: parsed.body,
      supplementalText: parsed.supplementalText,
      imageReferences: parsed.imageReferences,
      paragraphs: parsed.paragraphs,
      rawText: parsed.rawText,
      warnings: [...warnings, ...parsed.warnings],
      features: parsed.features,
      canonicalPath: outputPath,
    };
  }

  await convertHwpToHwpxWithJvm(referencePath, outputPath);
  const parsed = await parseStructuredHwpx(outputPath);
  warnings = [
    {
      code: 'REFERENCE_TEMPLATE_APPLIED',
      message: 'Created a new editable HWPX document by importing an HWP reference into canonical HWPX.',
    },
    {
      code: 'HWP_IMPORTED_TO_HWPX',
      message: 'The original HWP reference remains unchanged. The new document was created as canonical HWPX.',
    },
  ];

  return {
    sourceFormat: 'hwpx',
    projectionMode: parsed.features.projectionMode,
    documentMode: 'editable',
    readOnly: false,
    title: basename(outputPath),
    body: parsed.body,
    supplementalText: parsed.supplementalText,
    imageReferences: parsed.imageReferences,
    paragraphs: parsed.paragraphs,
    rawText: parsed.rawText,
    warnings: [...warnings, ...parsed.warnings],
    features: parsed.features,
    canonicalPath: outputPath,
  };
}

export async function createEditableWorkingCopy(
  session: DocumentSession,
  requestedOutputPath?: string
): Promise<{ targetPath: string; document: CoreDocumentData; warnings: ApiWarning[] }> {
  const targetPath = requestedOutputPath && requestedOutputPath.trim().length > 0
    ? normalizeHwpxOutputPath(requestedOutputPath)
    : await getUniqueSiblingPath(session.sourcePath, '.edit', '.hwpx');

  if (session.sourceFormat === 'hwpx' && session.projectionMode === 'structured_hwpx') {
    const parsed = await copyStructuredHwpx(session.sourcePath, targetPath);
    const warnings: ApiWarning[] = [
      {
        code: 'WORKING_COPY_CREATED',
        message: 'A managed .hwpx working copy was created. The original source file remains unchanged.',
      },
    ];
    return {
      targetPath,
      document: {
        sourceFormat: 'hwpx',
        projectionMode: 'structured_hwpx',
        documentMode: 'editable',
        readOnly: false,
        title: basename(targetPath),
        body: parsed.body,
        supplementalText: parsed.supplementalText,
        imageReferences: parsed.imageReferences,
        paragraphs: parsed.paragraphs,
        rawText: parsed.rawText,
        warnings,
        features: parsed.features,
        canonicalPath: targetPath,
      },
      warnings,
    };
  }

  let warnings: ApiWarning[];
  try {
    await convertHwpToHwpxWithJvm(session.sourcePath, targetPath);
    warnings = [
      {
        code: 'WORKING_COPY_CREATED',
        message: 'The original .hwp file remains unchanged. Edits will apply to an imported .hwpx working copy.',
      },
      {
        code: 'HWP_IMPORTED_TO_HWPX',
        message: 'The editable working copy was created by converting the original .hwp into canonical .hwpx.',
      },
    ];
  } catch {
    await saveHwpxText(targetPath, session.rawText, targetPath);
    warnings = [
      {
        code: 'WORKING_COPY_CREATED',
        message: 'The original .hwp file remains unchanged. Edits will apply to a generated .hwpx working copy.',
      },
      {
        code: 'TEXT_PROJECTION_WORKING_COPY',
        message: 'The editable copy was generated from text projection because the JVM HWP import core was unavailable.',
      },
    ];
  }

  const parsed = await parseStructuredHwpx(targetPath);

  return {
    targetPath,
    document: {
      sourceFormat: 'hwpx',
      projectionMode: 'structured_hwpx',
      documentMode: 'editable',
      readOnly: false,
      title: basename(targetPath),
      body: parsed.body,
      supplementalText: parsed.supplementalText,
      imageReferences: parsed.imageReferences,
      paragraphs: parsed.paragraphs,
      rawText: parsed.rawText,
      warnings,
      features: parsed.features,
      canonicalPath: targetPath,
    },
    warnings,
  };
}

export async function saveEditableDocumentCore(
  session: DocumentSession,
  payload: { text?: string; mutations?: TextMutationOperation[] }
): Promise<CoreDocumentData> {
  if (!session.workingCopyPath) {
    throw new Error('This document has no editable working copy.');
  }

  if (session.projectionMode === 'structured_hwpx') {
    const parsed = typeof payload.text === 'string'
      ? await saveStructuredHwpxWithPlainText(session.workingCopyPath, payload.text)
      : await saveStructuredHwpxWithMutations(session.workingCopyPath, payload.mutations ?? []);

    return {
      sourceFormat: 'hwpx',
      projectionMode: 'structured_hwpx',
      documentMode: 'editable',
      readOnly: false,
      title: basename(session.workingCopyPath),
      body: parsed.body,
      supplementalText: parsed.supplementalText,
      imageReferences: parsed.imageReferences,
      paragraphs: parsed.paragraphs,
      rawText: parsed.rawText,
      warnings: parsed.warnings,
      features: parsed.features,
      canonicalPath: session.workingCopyPath,
    };
  }

  const text = typeof payload.text === 'string'
    ? payload.text
    : payload.mutations?.map((mutation) => {
        switch (mutation.op) {
          case 'replace_text_in_paragraph':
          case 'insert_paragraph_after':
            return mutation.text;
          default:
            return '';
        }
      }).join('\n\n') ?? session.rawText;

  await saveHwpxText(session.workingCopyPath, text, session.workingCopyPath);
  const extracted = await extractHwpxText(session.workingCopyPath);

  return {
    sourceFormat: 'hwpx',
    projectionMode: 'text_projection',
    documentMode: 'editable',
    readOnly: false,
    title: basename(session.workingCopyPath),
    body: buildTextProjectionBody(extracted.paragraphs),
    supplementalText: [],
    imageReferences: [],
    paragraphs: extracted.paragraphs,
    rawText: extracted.rawText,
    warnings: session.warnings,
    features: buildTextProjectionFeatures('editable'),
    canonicalPath: session.workingCopyPath,
  };
}

export function createDocumentSession(
  sourcePath: string,
  document: CoreDocumentData,
  overrides?: Partial<Pick<DocumentSession, 'workingCopyPath' | 'currentCheckpointId'>>
): DocumentSession {
  const now = new Date().toISOString();
  return {
    documentId: `doc_${nanoid(12)}`,
    sourcePath,
    sourceFormat: document.sourceFormat,
    canonicalFormat: 'hwpx',
    workingCopyPath: overrides?.workingCopyPath ?? document.canonicalPath ?? null,
    documentMode: document.documentMode,
    readOnly: document.readOnly,
    title: document.title,
    projectionMode: document.projectionMode,
    body: document.body,
    supplementalText: document.supplementalText,
    imageReferences: document.imageReferences,
    paragraphs: document.paragraphs,
    rawText: document.rawText,
    warnings: document.warnings,
    features: document.features,
    currentCheckpointId: overrides?.currentCheckpointId ?? null,
    createdAt: now,
    updatedAt: now,
  };
}
