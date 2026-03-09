import express, { type NextFunction, type Request, type Response } from 'express';
import { readFileSync } from 'node:fs';
import { unlink } from 'node:fs/promises';
import { extname } from 'node:path';
import { nanoid } from 'nanoid';
import {
  createDocumentSession,
  createEditableWorkingCopy,
  createHwpxDocumentCore,
  createHwpxDocumentFromReferenceCore,
  openDocumentCore,
  saveEditableDocumentCore,
} from './core/documentCore.js';
import { checkJvmCoreHealth } from './core/jvmCore.js';
import { isDocumentTemplateId, listDocumentTemplateIds } from './core/hwpxTemplates.js';
import { analyzeReferenceTemplate } from './core/referenceTemplate.js';
import { buildImageResourceSummaries } from './core/imageResources.js';
import { buildImageTargetSummaries } from './core/imageTargets.js';
import { normalizeTemplateFillInstructions, planTemplateFill, validateTemplateFillRequirements } from './core/templateFill.js';
import { analyzeTableRepeatRegion } from './core/tableRepeatAnalysis.js';
import type {
  ApiWarning,
  DocumentSession,
  ErrorCode,
  MarkdownMode,
  OfficeExtensionConvertRequest,
  SourceFormat,
  TargetFormat,
  TextMutationOperation,
} from './types.js';
import { convertWithSource } from './converters/index.js';
import {
  fileExists,
  guessSourceFormat,
  normalizeAbsolutePath,
  normalizeTargetFormat,
  statSafe,
  writeTempFile,
} from './utils/files.js';
import { EphemeralFileStore } from './utils/fileStore.js';
import { DocumentSessionStore } from './sessionStore.js';
import { CheckpointStore } from './checkpointStore.js';

const packageJson = JSON.parse(
  readFileSync(new URL('../package.json', import.meta.url), 'utf-8')
) as { version?: string };

const app = express();
const port = Number(process.env.PORT || 8090);
const engineVersion = process.env.HWP_EXTENSION_VERSION || packageJson.version || '0.1.0';
const maxDocumentBytes = Number(process.env.MAX_DOCUMENT_BYTES || 32 * 1024 * 1024);
const fileStore = new EphemeralFileStore(15 * 60 * 1000);
const sessionStore = new DocumentSessionStore();
const checkpointStore = new CheckpointStore();

const SUPPORTED_TARGETS = new Set<TargetFormat>(['txt', 'md', 'html', 'json', 'docx', 'pdf']);

app.use(express.json({ limit: '25mb' }));
app.use((req: Request, res: Response, next: NextFunction) => {
  const requestId = req.header('x-request-id') || `req_${nanoid(10)}`;
  res.locals.requestId = requestId;
  res.setHeader('x-request-id', requestId);
  next();
});

class RequestError extends Error {
  constructor(
    public readonly status: number,
    public readonly code: ErrorCode,
    message: string,
    public readonly details?: Record<string, unknown>,
    public readonly retryable: boolean = false,
    public readonly warnings: ApiWarning[] = []
  ) {
    super(message);
  }
}

function baseUrl(req: Request): string {
  const forwardedProto = req.headers['x-forwarded-proto'];
  const proto = typeof forwardedProto === 'string' ? forwardedProto.split(',')[0] : req.protocol;
  return `${proto}://${req.get('host')}`;
}

function requestIdOf(res: Response): string {
  return String(res.locals.requestId || `req_${nanoid(10)}`);
}

function sendSuccess(res: Response, payload: Record<string, unknown>, status: number = 200): void {
  res.status(status).json({
    success: true,
    requestId: requestIdOf(res),
    engineVersion,
    warnings: [],
    ...payload,
  });
}

function sendError(res: Response, error: unknown): void {
  if (error instanceof RequestError) {
    res.status(error.status).json({
      success: false,
      requestId: requestIdOf(res),
      engineVersion,
      error: {
        code: error.code,
        message: error.message,
        details: error.details,
        retryable: error.retryable,
      },
      warnings: error.warnings,
    });
    return;
  }

  const message = error instanceof Error ? error.message : 'Internal error';
  res.status(500).json({
    success: false,
    requestId: requestIdOf(res),
    engineVersion,
    error: {
      code: 'INTERNAL_ERROR',
      message,
      retryable: false,
    },
    warnings: [],
  });
}

function downloadContentType(path: string): string {
  const ext = extname(path).toLowerCase();
  if (ext === '.pdf') return 'application/pdf';
  if (ext === '.docx') return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  if (ext === '.json') return 'application/json; charset=utf-8';
  if (ext === '.html') return 'text/html; charset=utf-8';
  if (ext === '.md') return 'text/markdown; charset=utf-8';
  return 'text/plain; charset=utf-8';
}

async function resolveSourceFile(request: OfficeExtensionConvertRequest): Promise<string> {
  if (request.filePath && request.filePath.trim().length > 0) {
    return normalizeAbsolutePath(request.filePath);
  }

  if (!request.url || request.url.trim().length === 0) {
    throw new RequestError(400, 'INVALID_REQUEST', 'Either filePath or url is required.');
  }

  const response = await fetch(request.url);
  if (!response.ok) {
    throw new RequestError(502, 'INTERNAL_ERROR', `Failed to download source file from url: ${response.status} ${response.statusText}`);
  }

  const sourceHint = request.filetype?.toLowerCase() || extname(new URL(request.url).pathname).replace('.', '').toLowerCase() || 'bin';
  const extension = sourceHint.startsWith('.') ? sourceHint : `.${sourceHint}`;
  const data = Buffer.from(await response.arrayBuffer());
  return writeTempFile(data, extension);
}

async function assertSupportedPath(filePathRaw: string): Promise<{ filePath: string; sourceFormat: SourceFormat }> {
  const filePath = normalizeAbsolutePath(filePathRaw);
  const exists = await fileExists(filePath);
  if (!exists) {
    throw new RequestError(404, 'FILE_NOT_FOUND', `File not found: ${filePath}`, { path: filePath });
  }

  const sourceFormat = guessSourceFormat(undefined, filePath);
  if (!sourceFormat) {
    throw new RequestError(400, 'UNSUPPORTED_EXTENSION', 'Only .hwp/.hwpx files are supported.', { path: filePath });
  }

  const fileStat = await statSafe(filePath);
  if (!fileStat) {
    throw new RequestError(404, 'FILE_NOT_FOUND', `File not found: ${filePath}`, { path: filePath });
  }

  if (fileStat.size > maxDocumentBytes) {
    throw new RequestError(413, 'DOCUMENT_TOO_LARGE', `Document exceeds the size limit (${maxDocumentBytes} bytes).`, {
      path: filePath,
      size: fileStat.size,
      limit: maxDocumentBytes,
    });
  }

  return { filePath, sourceFormat };
}

function serializeSessionDocument(session: DocumentSession) {
  return {
    documentId: session.documentId,
    sourcePath: session.sourcePath,
    workingCopyPath: session.workingCopyPath,
    sourceFormat: session.sourceFormat,
    canonicalFormat: session.canonicalFormat,
    documentMode: session.documentMode,
    readOnly: session.readOnly,
    title: session.title,
    projectionMode: session.projectionMode,
    body: session.body,
    supplementalText: session.supplementalText,
    imageReferences: session.imageReferences,
    imageTargets: buildImageTargetSummaries(session.imageReferences),
    imageResources: buildImageResourceSummaries(session.imageReferences),
    text: session.rawText,
    paragraphs: session.paragraphs,
    features: session.features,
    currentCheckpointId: session.currentCheckpointId,
    createdAt: session.createdAt,
    updatedAt: session.updatedAt,
  };
}

function normalizeMarkdownMode(value: unknown): MarkdownMode | undefined {
  if (value === undefined || value === null || value === '') {
    return undefined;
  }
  if (value === 'clean' || value === 'fidelity') {
    return value;
  }
  throw new RequestError(400, 'INVALID_REQUEST', `Unsupported markdownMode: ${String(value)}`, {
    supportedMarkdownModes: ['clean', 'fidelity'],
  });
}

function requireSession(documentId: unknown): DocumentSession {
  if (typeof documentId !== 'string' || documentId.trim().length === 0) {
    throw new RequestError(400, 'INVALID_REQUEST', 'documentId is required.');
  }

  const session = sessionStore.get(documentId);
  if (!session) {
    throw new RequestError(404, 'FILE_NOT_FOUND', `Document session not found: ${documentId}`, { documentId });
  }

  return session;
}

function paragraphsToText(paragraphs: string[]): string {
  return `${paragraphs.join('\n\n')}${paragraphs.length > 0 ? '\n' : ''}`;
}

function applyTextProjectionMutations(paragraphs: string[], mutations: TextMutationOperation[]): string[] {
  const next = [...paragraphs];

  for (const mutation of mutations) {
    switch (mutation.op) {
      case 'replace_image_asset':
      case 'replace_image_instance_asset':
      case 'delete_image_instance':
      case 'insert_image_from_prototype':
      case 'update_image_placement':
      case 'replace_text_in_supplemental_paragraph':
      case 'replace_table_cell_text':
      case 'insert_table_row':
      case 'delete_table_row':
      case 'clone_table_region':
        throw new RequestError(400, 'INVALID_REQUEST', `Mutation "${mutation.op}" is not supported by the text-projection fallback engine.`);

      case 'replace_text_in_paragraph':
      case 'insert_paragraph_after':
      case 'delete_paragraph': {
        if (mutation.sectionIndex !== 0) {
          throw new RequestError(400, 'INVALID_REQUEST', 'Only sectionIndex=0 is supported by the text-projection fallback engine.');
        }

        const paragraphIndex = mutation.blockIndex;
        if (paragraphIndex < 0 || paragraphIndex >= next.length) {
          throw new RequestError(400, 'INVALID_REQUEST', `blockIndex ${paragraphIndex} is out of range.`, { blockIndex: paragraphIndex });
        }

        if (mutation.op === 'replace_text_in_paragraph') {
          next[paragraphIndex] = mutation.text.trim();
          continue;
        }

        if (mutation.op === 'insert_paragraph_after') {
          next.splice(paragraphIndex + 1, 0, mutation.text.trim());
          continue;
        }

        next.splice(paragraphIndex, 1);
        continue;
      }
    }
  }

  return next.filter((line) => line.length > 0);
}

type ImageReplacementMutation = Extract<TextMutationOperation, { op: 'replace_image_asset' }>;
type ImageInstanceReplacementMutation = Extract<TextMutationOperation, { op: 'replace_image_instance_asset' }>;
type ImageDeletionMutation = Extract<TextMutationOperation, { op: 'delete_image_instance' }>;
type ImageInsertionMutation = Extract<TextMutationOperation, { op: 'insert_image_from_prototype' }>;
type ImagePlacementMutation = Extract<TextMutationOperation, { op: 'update_image_placement' }>;

function normalizeLocatorField(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim().length > 0 ? value.trim() : undefined;
}

function validateImageLocator(
  session: DocumentSession,
  locator: { targetId?: string; binaryItemId?: string; objectId?: string; instanceId?: string },
  mode: 'asset_replacement' | 'instance_asset_replacement' | 'placement_update'
): void {
  const locatorCount = [locator.targetId, locator.binaryItemId, locator.objectId, locator.instanceId]
    .filter((value) => typeof value === 'string' && value.length > 0)
    .length;
  if (locatorCount !== 1) {
    const operationLabel = mode === 'asset_replacement'
      ? 'replace_image_asset'
      : mode === 'instance_asset_replacement'
        ? 'replace_image_instance_asset'
        : 'update_image_placement';
    const locatorLabel = mode === 'asset_replacement'
      ? 'targetId or binaryItemId'
      : 'targetId, binaryItemId, objectId, or instanceId';
    throw new RequestError(400, 'INVALID_REQUEST', `Specify exactly one locator for ${operationLabel}: ${locatorLabel}.`);
  }

  if (locator.targetId) {
    const matchedReferences = session.imageReferences.filter((reference) => reference.targetId === locator.targetId);
    if (matchedReferences.length === 0) {
      throw new RequestError(404, 'IMAGE_TARGET_NOT_FOUND', `No image reference matched targetId "${locator.targetId}".`, {
        documentId: session.documentId,
        targetId: locator.targetId,
      });
    }
    if (matchedReferences.length > 1) {
      throw new RequestError(
        400,
        'IMAGE_TARGET_AMBIGUOUS',
        `Target "${locator.targetId}" contains ${matchedReferences.length} image references. Use ${mode === 'asset_replacement' ? 'binaryItemId for deterministic replacement' : 'a deterministic single-image locator'}.`,
        {
          documentId: session.documentId,
          targetId: locator.targetId,
          imageCount: matchedReferences.length,
          binaryItemIds: [...new Set(matchedReferences.map((reference) => reference.binaryItemId))],
        }
      );
    }
  }

  if (locator.binaryItemId) {
    const matchedReferences = session.imageReferences.filter((reference) => reference.binaryItemId === locator.binaryItemId);
    if (matchedReferences.length === 0) {
      throw new RequestError(404, 'IMAGE_TARGET_NOT_FOUND', `No image reference matched binaryItemId "${locator.binaryItemId}".`, {
        documentId: session.documentId,
        binaryItemId: locator.binaryItemId,
      });
    }
    if ((mode === 'placement_update' || mode === 'instance_asset_replacement') && matchedReferences.length > 1) {
      throw new RequestError(
        400,
        'IMAGE_TARGET_AMBIGUOUS',
        `binaryItemId "${locator.binaryItemId}" matched ${matchedReferences.length} image references. Use targetId, objectId, or instanceId for deterministic single-image updates.`,
        {
          documentId: session.documentId,
          binaryItemId: locator.binaryItemId,
          imageCount: matchedReferences.length,
          targetIds: [...new Set(matchedReferences.map((reference) => reference.targetId))],
        }
      );
    }
  }

  if ((mode === 'placement_update' || mode === 'instance_asset_replacement') && locator.objectId) {
    const matchedReferences = session.imageReferences.filter((reference) => reference.placement?.objectId === locator.objectId);
    if (matchedReferences.length === 0) {
      throw new RequestError(404, 'IMAGE_TARGET_NOT_FOUND', `No image reference matched objectId "${locator.objectId}".`, {
        documentId: session.documentId,
        objectId: locator.objectId,
      });
    }
    if (matchedReferences.length > 1) {
      throw new RequestError(400, 'IMAGE_TARGET_AMBIGUOUS', `objectId "${locator.objectId}" matched ${matchedReferences.length} image references. Use a deterministic single-image locator.`, {
        documentId: session.documentId,
        objectId: locator.objectId,
        imageCount: matchedReferences.length,
      });
    }
  }

  if ((mode === 'placement_update' || mode === 'instance_asset_replacement') && locator.instanceId) {
    const matchedReferences = session.imageReferences.filter((reference) => reference.placement?.instanceId === locator.instanceId);
    if (matchedReferences.length === 0) {
      throw new RequestError(404, 'IMAGE_TARGET_NOT_FOUND', `No image reference matched instanceId "${locator.instanceId}".`, {
        documentId: session.documentId,
        instanceId: locator.instanceId,
      });
    }
    if (matchedReferences.length > 1) {
      throw new RequestError(400, 'IMAGE_TARGET_AMBIGUOUS', `instanceId "${locator.instanceId}" matched ${matchedReferences.length} image references. Use a deterministic single-image locator.`, {
        documentId: session.documentId,
        instanceId: locator.instanceId,
        imageCount: matchedReferences.length,
      });
    }
  }
}

function assertNonNegativeInteger(value: number | undefined, fieldPath: string): void {
  if (value === undefined) {
    return;
  }
  if (!Number.isInteger(value) || value < 0) {
    throw new RequestError(400, 'INVALID_REQUEST', `${fieldPath} must be a non-negative integer.`);
  }
}

function resolveAssetReplacementImpact(
  session: DocumentSession,
  locator: { targetId?: string; binaryItemId?: string }
): { binaryItemId: string; references: typeof session.imageReferences } {
  if (locator.binaryItemId) {
    const references = session.imageReferences.filter((reference) => reference.binaryItemId === locator.binaryItemId);
    return {
      binaryItemId: locator.binaryItemId,
      references,
    };
  }

  const targetReference = session.imageReferences.find((reference) => reference.targetId === locator.targetId);
  if (!targetReference) {
    return {
      binaryItemId: '',
      references: [],
    };
  }
  return {
    binaryItemId: targetReference.binaryItemId,
    references: session.imageReferences.filter((reference) => reference.binaryItemId === targetReference.binaryItemId),
  };
}

function hasDocumentTarget(session: DocumentSession, targetId: string): boolean {
  return session.body.sections.some((section) => section.blocks.some((block) => {
    if (block.type === 'paragraph') {
      return block.blockId === targetId;
    }
    return block.cells.some((cell) => cell.cellId === targetId);
  })) || session.supplementalText.some((entry) => entry.targetId === targetId);
}

function validatePlacementPatch(patch: ImagePlacementMutation['patch'], operationLabel: string): void {
  const hasPatch =
    typeof patch?.textWrap === 'string' ||
    typeof patch?.textFlow === 'string' ||
    typeof patch?.zOrder === 'number' ||
    typeof patch?.width === 'number' ||
    typeof patch?.height === 'number' ||
    patch?.margins !== undefined ||
    patch?.clip !== undefined;

  if (!hasPatch) {
    throw new RequestError(400, 'INVALID_REQUEST', `${operationLabel} requires at least one patch field.`);
  }

  if (typeof patch.textWrap === 'string') {
    patch.textWrap = patch.textWrap.trim();
    if (patch.textWrap.length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'patch.textWrap cannot be empty.');
    }
  }
  if (typeof patch.textFlow === 'string') {
    patch.textFlow = patch.textFlow.trim();
    if (patch.textFlow.length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'patch.textFlow cannot be empty.');
    }
  }

  assertNonNegativeInteger(patch.zOrder, 'patch.zOrder');
  assertNonNegativeInteger(patch.width, 'patch.width');
  assertNonNegativeInteger(patch.height, 'patch.height');
  assertNonNegativeInteger(patch.margins?.left, 'patch.margins.left');
  assertNonNegativeInteger(patch.margins?.right, 'patch.margins.right');
  assertNonNegativeInteger(patch.margins?.top, 'patch.margins.top');
  assertNonNegativeInteger(patch.margins?.bottom, 'patch.margins.bottom');
  assertNonNegativeInteger(patch.clip?.left, 'patch.clip.left');
  assertNonNegativeInteger(patch.clip?.right, 'patch.clip.right');
  assertNonNegativeInteger(patch.clip?.top, 'patch.clip.top');
  assertNonNegativeInteger(patch.clip?.bottom, 'patch.clip.bottom');
}

async function validateImageMutations(session: DocumentSession, mutations: TextMutationOperation[]): Promise<void> {
  for (const mutation of mutations) {
    if (mutation.op === 'delete_image_instance') {
      mutation.targetId = normalizeLocatorField(mutation.targetId);
      mutation.binaryItemId = normalizeLocatorField(mutation.binaryItemId);
      mutation.objectId = normalizeLocatorField(mutation.objectId);
      mutation.instanceId = normalizeLocatorField(mutation.instanceId);
      validateImageLocator(session, mutation, 'instance_asset_replacement');
      continue;
    }

    if (mutation.op === 'replace_image_asset' || mutation.op === 'replace_image_instance_asset' || mutation.op === 'insert_image_from_prototype') {
      if (typeof mutation.imagePath !== 'string' || mutation.imagePath.trim().length === 0) {
        throw new RequestError(400, 'INVALID_REQUEST', `${mutation.op} requires imagePath.`);
      }

      mutation.imagePath = normalizeAbsolutePath(mutation.imagePath);
      const imageExists = await fileExists(mutation.imagePath);
      if (!imageExists) {
        throw new RequestError(404, 'FILE_NOT_FOUND', `Replacement image not found: ${mutation.imagePath}`, {
          imagePath: mutation.imagePath,
        });
      }

      const imageStat = await statSafe(mutation.imagePath);
      if (!imageStat) {
        throw new RequestError(404, 'FILE_NOT_FOUND', `Replacement image not found: ${mutation.imagePath}`, {
          imagePath: mutation.imagePath,
        });
      }

      if (imageStat.size > maxDocumentBytes) {
        throw new RequestError(413, 'DOCUMENT_TOO_LARGE', `Replacement image exceeds the size limit (${maxDocumentBytes} bytes).`, {
          imagePath: mutation.imagePath,
          size: imageStat.size,
          limit: maxDocumentBytes,
        });
      }

      if (mutation.op === 'insert_image_from_prototype') {
        mutation.destinationTargetId = normalizeLocatorField(mutation.destinationTargetId) ?? '';
        mutation.prototypeTargetId = normalizeLocatorField(mutation.prototypeTargetId);
        mutation.prototypeBinaryItemId = normalizeLocatorField(mutation.prototypeBinaryItemId);
        mutation.prototypeObjectId = normalizeLocatorField(mutation.prototypeObjectId);
        mutation.prototypeInstanceId = normalizeLocatorField(mutation.prototypeInstanceId);

        if (mutation.destinationTargetId.length === 0) {
          throw new RequestError(400, 'INVALID_REQUEST', 'insert_image_from_prototype requires destinationTargetId.');
        }
        if (!hasDocumentTarget(session, mutation.destinationTargetId)) {
          throw new RequestError(404, 'FILE_NOT_FOUND', `Destination target not found: ${mutation.destinationTargetId}`, {
            destinationTargetId: mutation.destinationTargetId,
          });
        }

        const prototypeLocatorCount = [
          mutation.prototypeTargetId,
          mutation.prototypeBinaryItemId,
          mutation.prototypeObjectId,
          mutation.prototypeInstanceId,
        ].filter((value) => typeof value === 'string' && value.length > 0).length;
        if (prototypeLocatorCount > 0) {
          validateImageLocator(session, {
            targetId: mutation.prototypeTargetId,
            binaryItemId: mutation.prototypeBinaryItemId,
            objectId: mutation.prototypeObjectId,
            instanceId: mutation.prototypeInstanceId,
          }, 'instance_asset_replacement');
        }

        if (mutation.patch) {
          validatePlacementPatch(mutation.patch, 'insert_image_from_prototype');
        }
        continue;
      }

      mutation.targetId = normalizeLocatorField(mutation.targetId);
      mutation.binaryItemId = normalizeLocatorField(mutation.binaryItemId);
      if (mutation.op === 'replace_image_instance_asset') {
        mutation.objectId = normalizeLocatorField(mutation.objectId);
        mutation.instanceId = normalizeLocatorField(mutation.instanceId);
      }
      validateImageLocator(session, mutation, mutation.op === 'replace_image_asset' ? 'asset_replacement' : 'instance_asset_replacement');
      continue;
    }

    if (mutation.op === 'update_image_placement') {
      mutation.targetId = normalizeLocatorField(mutation.targetId);
      mutation.binaryItemId = normalizeLocatorField(mutation.binaryItemId);
      mutation.objectId = normalizeLocatorField(mutation.objectId);
      mutation.instanceId = normalizeLocatorField(mutation.instanceId);
      validateImageLocator(session, mutation, 'placement_update');
      validatePlacementPatch(mutation.patch, 'update_image_placement');
    }
  }
}

function savePayloadForSession(session: DocumentSession, body: { text?: unknown; mutations?: unknown }): { text?: string; mutations?: TextMutationOperation[] } {
  if (typeof body.text === 'string') {
    if (session.projectionMode === 'structured_hwpx' && !session.features.allowPlainTextSave) {
      throw new RequestError(
        400,
        'INVALID_REQUEST',
        'This document contains structured content. Save it via structured mutations instead of replacing the full text projection.',
        { projectionMode: session.projectionMode }
      );
    }
    return { text: body.text };
  }

  if (Array.isArray(body.mutations)) {
    const mutations = body.mutations as TextMutationOperation[];
    if (session.projectionMode === 'text_projection') {
      const nextParagraphs = applyTextProjectionMutations(session.paragraphs, mutations);
      return { text: paragraphsToText(nextParagraphs) };
    }
    return { mutations };
  }

  throw new RequestError(400, 'INVALID_REQUEST', 'Either text or mutations is required for save.');
}

function renderOpenPage(filePath: string): string {
  const safePath = JSON.stringify(filePath);
  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>HWP Extension Editor</title>
  <style>
    :root { color-scheme: light dark; }
    body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #111; color: #eee; }
    .wrap { display: flex; flex-direction: column; height: 100vh; }
    .toolbar { display: flex; align-items: center; gap: 8px; padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,0.15); background: rgba(0,0,0,0.2); flex-wrap: wrap; }
    .path { font-size: 12px; opacity: 0.8; flex: 1 1 320px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    button { border: 1px solid rgba(255,255,255,0.3); border-radius: 6px; padding: 6px 10px; background: rgba(255,255,255,0.06); color: inherit; cursor: pointer; }
    button:hover { background: rgba(255,255,255,0.14); }
    button[disabled] { opacity: 0.45; cursor: not-allowed; }
    .status { font-size: 12px; opacity: 0.85; }
    .mode { font-size: 12px; padding: 4px 8px; border: 1px solid rgba(255,255,255,0.2); border-radius: 999px; }
    .warnings { padding: 10px 12px; font-size: 12px; line-height: 1.5; border-bottom: 1px solid rgba(255,255,255,0.12); background: rgba(255,255,255,0.04); display: none; white-space: pre-wrap; }
    .supplemental { padding: 10px 12px; font-size: 12px; line-height: 1.6; border-bottom: 1px solid rgba(255,255,255,0.12); background: rgba(255,255,255,0.03); display: none; }
    .image-targets { padding: 10px 12px; font-size: 12px; line-height: 1.6; border-bottom: 1px solid rgba(255,255,255,0.12); background: rgba(255,255,255,0.03); display: none; }
    .image-resources { padding: 10px 12px; font-size: 12px; line-height: 1.6; border-bottom: 1px solid rgba(255,255,255,0.12); background: rgba(255,255,255,0.03); display: none; }
    .supplemental h3 { margin: 0 0 6px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.04em; opacity: 0.8; }
    .image-targets h3 { margin: 0 0 6px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.04em; opacity: 0.8; }
    .image-resources h3 { margin: 0 0 6px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.04em; opacity: 0.8; }
    .supplemental-group + .supplemental-group { margin-top: 10px; }
    .supplemental-item { white-space: pre-wrap; opacity: 0.92; }
    .supplemental-meta { opacity: 0.6; margin-bottom: 4px; }
    .image-item + .image-item { margin-top: 10px; }
    .image-meta { opacity: 0.82; margin-bottom: 2px; }
    .image-detail { opacity: 0.62; white-space: pre-wrap; }
    textarea { flex: 1; width: 100%; box-sizing: border-box; border: 0; outline: none; padding: 14px; resize: none; background: #161616; color: #f4f4f4; font: 14px/1.6 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="toolbar">
      <div class="path" id="path"></div>
      <div class="mode" id="mode">Opening...</div>
      <button id="reloadBtn">Reload</button>
      <button id="createCopyBtn" hidden>Create Editable Copy</button>
      <button id="saveBtn" disabled>Save</button>
      <div class="status" id="status">Loading...</div>
    </div>
    <div class="warnings" id="warnings"></div>
    <div class="supplemental" id="supplemental"></div>
    <div class="image-targets" id="imageTargets"></div>
    <div class="image-resources" id="imageResources"></div>
    <textarea id="editor" spellcheck="false"></textarea>
  </div>

  <script>
    const initialPath = ${safePath};
    const pathEl = document.getElementById('path');
    const modeEl = document.getElementById('mode');
    const statusEl = document.getElementById('status');
    const warningsEl = document.getElementById('warnings');
    const supplementalEl = document.getElementById('supplemental');
    const imageTargetsEl = document.getElementById('imageTargets');
    const imageResourcesEl = document.getElementById('imageResources');
    const editorEl = document.getElementById('editor');
    const reloadBtn = document.getElementById('reloadBtn');
    const createCopyBtn = document.getElementById('createCopyBtn');
    const saveBtn = document.getElementById('saveBtn');

    let documentId = null;
    let currentOutputPath = initialPath;

    pathEl.textContent = initialPath;

    function setStatus(message) {
      statusEl.textContent = message;
    }

    function publishDocumentState(document) {
      if (!document) return;
      window.parent?.postMessage({
        type: 'HWP_EXTENSION_DOCUMENT_STATE',
        filePath: currentOutputPath,
        originalPath: initialPath,
        document,
      }, '*');
    }

    function setWarnings(warnings) {
      if (!warnings || warnings.length === 0) {
        warningsEl.style.display = 'none';
        warningsEl.textContent = '';
        return;
      }
      warningsEl.style.display = 'block';
      warningsEl.textContent = warnings.map((warning) => '- ' + warning.message).join('\\n');
    }

    function escapeHtml(value) {
      return String(value)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll(\"'\", '&#39;');
    }

    function setSupplementalText(supplementalText) {
      if (!Array.isArray(supplementalText) || supplementalText.length === 0) {
        supplementalEl.style.display = 'none';
        supplementalEl.innerHTML = '';
        return;
      }

      const groups = new Map();
      supplementalText.forEach((entry) => {
        const bucket = groups.get(entry.kind) || [];
        bucket.push(entry);
        groups.set(entry.kind, bucket);
      });

      const parts = [];
      for (const [kind, entries] of groups.entries()) {
        const items = entries.map((entry) => {
          const meta = entry.editable ? 'editable target' : 'read-only target';
          return '<div class=\"supplemental-item\"><div class=\"supplemental-meta\">' + escapeHtml(entry.targetId) + ' · ' + meta + '</div>' + escapeHtml(entry.text || '') + '</div>';
        }).join('');
        parts.push('<div class=\"supplemental-group\"><h3>' + escapeHtml(kind) + '</h3>' + items + '</div>');
      }

      supplementalEl.innerHTML = parts.join('');
      supplementalEl.style.display = 'block';
    }

    function setImageTargets(imageTargets, imageReferences) {
      if (!Array.isArray(imageTargets) || imageTargets.length === 0) {
        imageTargetsEl.style.display = 'none';
        imageTargetsEl.innerHTML = '';
        return;
      }

      const referencesByTarget = new Map();
      if (Array.isArray(imageReferences)) {
        imageReferences.forEach((reference) => {
          const bucket = referencesByTarget.get(reference.targetId) || [];
          bucket.push(reference);
          referencesByTarget.set(reference.targetId, bucket);
        });
      }

      const items = imageTargets.map((entry) => {
        const assetLocator = entry.recommendedAssetLocator || entry.recommendedLocator || 'targetId';
        const placementLocator = entry.recommendedPlacementLocator || 'none';
        const assets = Array.isArray(entry.assetFileNames) && entry.assetFileNames.length > 0
          ? entry.assetFileNames.join(', ')
          : 'no resolved asset path';
        const binaryIds = Array.isArray(entry.binaryItemIds) && entry.binaryItemIds.length > 0
          ? entry.binaryItemIds.join(', ')
          : 'none';
        const objectIds = Array.isArray(entry.objectIds) && entry.objectIds.length > 0
          ? entry.objectIds.join(', ')
          : 'none';
        const instanceIds = Array.isArray(entry.instanceIds) && entry.instanceIds.length > 0
          ? entry.instanceIds.join(', ')
          : 'none';
        const refs = referencesByTarget.get(entry.targetId) || [];
        const placement = refs.length === 1 && refs[0].placement
          ? 'wrap=' + (refs[0].placement.textWrap || 'n/a')
            + ', flow=' + (refs[0].placement.textFlow || 'n/a')
            + ', size=' + String(refs[0].width || 0) + 'x' + String(refs[0].height || 0)
          : 'multiple images or no placement metadata';
        return '<div class=\"image-item\">'
          + '<div class=\"image-meta\">'
          + escapeHtml(entry.targetId) + ' · ' + escapeHtml(entry.kind || 'image')
          + '</div>'
          + '<div class=\"image-detail\">assets: ' + escapeHtml(assets) + '\\n'
          + 'binaryItemIds: ' + escapeHtml(binaryIds) + '\\n'
          + 'objectIds: ' + escapeHtml(objectIds) + '\\n'
          + 'instanceIds: ' + escapeHtml(instanceIds) + '\\n'
          + 'count: ' + escapeHtml(String(entry.imageCount || 0)) + '\\n'
          + 'asset locator: ' + escapeHtml(String(assetLocator)) + '\\n'
          + 'placement locator: ' + escapeHtml(String(placementLocator)) + '\\n'
          + 'placement: ' + escapeHtml(placement)
          + '</div>'
          + '</div>';
      }).join('');

      imageTargetsEl.innerHTML = '<h3>image targets</h3>' + items;
      imageTargetsEl.style.display = 'block';
    }

    function setImageResources(imageResources) {
      if (!Array.isArray(imageResources) || imageResources.length === 0) {
        imageResourcesEl.style.display = 'none';
        imageResourcesEl.innerHTML = '';
        return;
      }

      const items = imageResources.map((entry) => {
        const targets = Array.isArray(entry.targetIds) && entry.targetIds.length > 0
          ? entry.targetIds.join(', ')
          : 'none';
        const scope = entry.assetReplacementScope || 'single_target';
        const placement = entry.deterministicPlacementUpdate ? 'exact placement allowed' : 'shared instance locator required';
        return '<div class=\"image-item\">'
          + '<div class=\"image-meta\">'
          + escapeHtml(entry.binaryItemId || 'unknown') + ' · ' + escapeHtml(entry.mediaType || 'unknown')
          + '</div>'
          + '<div class=\"image-detail\">asset: ' + escapeHtml(entry.assetFileName || entry.assetPath || 'unresolved') + '\\n'
          + 'references: ' + escapeHtml(String(entry.referenceCount || 0)) + '\\n'
          + 'targets: ' + escapeHtml(targets) + '\\n'
          + 'asset scope: ' + escapeHtml(String(scope)) + '\\n'
          + 'placement: ' + escapeHtml(placement)
          + '</div>'
          + '</div>';
      }).join('');

      imageResourcesEl.innerHTML = '<h3>image resources</h3>' + items;
      imageResourcesEl.style.display = 'block';
    }

    function updateMode(document) {
      const plainTextEditable = document.documentMode === 'editable' && document.features?.allowPlainTextSave !== false;
      modeEl.textContent = document.documentMode + ' / ' + document.projectionMode;
      editorEl.disabled = !plainTextEditable;
      saveBtn.disabled = !plainTextEditable;
      createCopyBtn.hidden = document.documentMode === 'editable';
      if (document.documentMode === 'editable' && !plainTextEditable) {
        setStatus('Structured content: use API mutations from the host editor');
      }
    }

    async function loadByDocumentId() {
      const res = await fetch('/document?documentId=' + encodeURIComponent(documentId));
      const json = await res.json();
      if (!res.ok || !json.success) {
        throw new Error(json.error?.message || 'Failed to load document');
      }
      editorEl.value = json.document.text || '';
      updateMode(json.document);
      setSupplementalText(json.document.supplementalText || []);
      setImageTargets(json.document.imageTargets || [], json.document.imageReferences || []);
      setImageResources(json.document.imageResources || []);
      if (json.document.workingCopyPath) {
        currentOutputPath = json.document.workingCopyPath;
        pathEl.textContent = json.document.workingCopyPath;
      }
      setWarnings(json.warnings || []);
      publishDocumentState(json.document);
      return json.document;
    }

    async function openDocument() {
      setStatus('Opening...');
      const res = await fetch('/document/open', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path: initialPath }),
      });
      const json = await res.json();
      if (!res.ok || !json.success) {
        throw new Error(json.error?.message || 'Failed to open document');
      }
      documentId = json.document.documentId;
      currentOutputPath = json.document.workingCopyPath || json.document.sourcePath;
      pathEl.textContent = currentOutputPath;
      setWarnings(json.warnings || []);
      setSupplementalText(json.document.supplementalText || []);
      setImageTargets(json.document.imageTargets || [], json.document.imageReferences || []);
      setImageResources(json.document.imageResources || []);
      await loadByDocumentId();
      setStatus('Loaded');
      window.parent?.postMessage({ type: 'ONLYOFFICE_DOCUMENT_READY' }, '*');
    }

    async function createEditableCopy() {
      if (!documentId) return;
      setStatus('Creating working copy...');
      const res = await fetch('/document/fork-editable-copy', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ documentId }),
      });
      const json = await res.json();
      if (!res.ok || !json.success) {
        throw new Error(json.error?.message || 'Failed to create working copy');
      }
      currentOutputPath = json.document.workingCopyPath || currentOutputPath;
      pathEl.textContent = currentOutputPath;
      setWarnings(json.warnings || []);
      await loadByDocumentId();
      setStatus('Editable copy ready');
    }

    async function saveDocument() {
      if (!documentId) return;
      setStatus('Saving...');
      const res = await fetch('/document/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ documentId, text: editorEl.value }),
      });
      const json = await res.json();
      if (!res.ok || !json.success) {
        throw new Error(json.error?.message || 'Failed to save document');
      }
      setWarnings(json.warnings || []);
      currentOutputPath = json.document.workingCopyPath || currentOutputPath;
      pathEl.textContent = currentOutputPath;
      setStatus('Saved');
      publishDocumentState(json.document);
      window.parent?.postMessage({
        type: 'HWP_EXTENSION_DOCUMENT_SAVED',
        filePath: currentOutputPath,
        originalPath: initialPath,
      }, '*');
    }

    reloadBtn.addEventListener('click', () => {
      void (async () => {
        try {
          if (!documentId) {
            await openDocument();
          } else {
            await loadByDocumentId();
            setStatus('Reloaded');
          }
        } catch (error) {
          setStatus('Reload failed');
          alert(error.message || 'Failed to reload document');
        }
      })();
    });

    createCopyBtn.addEventListener('click', () => {
      void (async () => {
        try {
          await createEditableCopy();
        } catch (error) {
          setStatus('Create copy failed');
          alert(error.message || 'Failed to create editable copy');
        }
      })();
    });

    saveBtn.addEventListener('click', () => {
      void (async () => {
        try {
          await saveDocument();
        } catch (error) {
          setStatus('Save failed');
          alert(error.message || 'Failed to save document');
        }
      })();
    });

    window.addEventListener('message', (event) => {
      if (!event.data || typeof event.data !== 'object') {
        return;
      }
      if (event.data.type === 'HWP_EXTENSION_HOST_REFRESH' && (!event.data.documentId || event.data.documentId === documentId)) {
        void (async () => {
          try {
            if (documentId) {
              await loadByDocumentId();
              setStatus('Refreshed');
            }
          } catch (error) {
            console.error('[HWP open page] host refresh failed:', error);
          }
        })();
      }
    });

    void (async () => {
      try {
        await openDocument();
      } catch (error) {
        setStatus('Open failed');
        alert(error.message || 'Failed to open document');
      }
    })();
  </script>
</body>
</html>`;
}

app.get('/healthcheck', async (_req, res) => {
  const wantsJson = _req.query.full === '1' || _req.query.format === 'json';
  if (!wantsJson) {
    res.status(200).send('true');
    return;
  }

  const jvmHealthy = await checkJvmCoreHealth();

  sendSuccess(res, {
    status: 'ok',
    sidecar: {
      healthy: true,
      ready: true,
      mode: 'embedded-node-dom',
      authoritativeFormats: ['hwpx'],
      fallbackFormats: ['hwp'],
      jvmImportCore: {
        healthy: jvmHealthy,
        available: jvmHealthy,
      },
    },
  });
});

app.post('/document/open', async (req, res) => {
  try {
    const pathRaw = req.body?.path;
    if (typeof pathRaw !== 'string' || pathRaw.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'path is required.');
    }

    const { filePath, sourceFormat } = await assertSupportedPath(pathRaw);
    const document = await openDocumentCore(filePath, sourceFormat);
    const session = createDocumentSession(filePath, document);
    sessionStore.put(session);

    sendSuccess(res, {
      document: serializeSessionDocument(session),
      warnings: session.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.get('/document', async (req, res) => {
  try {
    const documentIdRaw = req.query.documentId;
    if (typeof documentIdRaw === 'string' && documentIdRaw.trim().length > 0) {
      const session = requireSession(documentIdRaw);
      sendSuccess(res, {
        document: serializeSessionDocument(session),
        warnings: session.warnings,
      });
      return;
    }

    const filePathRaw = req.query.filepath;
    if (typeof filePathRaw !== 'string' || filePathRaw.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'filepath or documentId query is required.');
    }

    const { filePath, sourceFormat } = await assertSupportedPath(filePathRaw);
    const document = await openDocumentCore(filePath, sourceFormat);

    sendSuccess(res, {
      sourceFormat,
      projectionMode: document.projectionMode,
      body: document.body,
      supplementalText: document.supplementalText,
      imageReferences: document.imageReferences,
      imageTargets: buildImageTargetSummaries(document.imageReferences),
      imageResources: buildImageResourceSummaries(document.imageReferences),
      text: document.rawText,
      paragraphs: document.paragraphs,
      features: document.features,
      warnings: document.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.get('/document/features', (req, res) => {
  try {
    const session = requireSession(req.query.documentId);
    sendSuccess(res, {
      features: session.features,
      warnings: session.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.get('/document/templates', (_req, res) => {
  sendSuccess(res, {
    templates: listDocumentTemplateIds(),
  });
});

app.post('/document/analyze-reference', async (req, res) => {
  try {
    const documentIdRaw = req.body?.documentId;
    const pathRaw = req.body?.path;

    if (documentIdRaw !== undefined && pathRaw !== undefined) {
      throw new RequestError(400, 'INVALID_REQUEST', 'Specify either documentId or path, not both.');
    }

    if (typeof documentIdRaw === 'string' && documentIdRaw.trim().length > 0) {
      const session = requireSession(documentIdRaw);
      const analysisPath = session.workingCopyPath ?? session.sourcePath;
      const analysisFormat: SourceFormat = session.workingCopyPath ? 'hwpx' : session.sourceFormat;
      const analysis = await analyzeReferenceTemplate(analysisPath, analysisFormat);
      sendSuccess(res, {
        analysis,
        warnings: analysis.warnings,
      });
      return;
    }

    if (typeof pathRaw !== 'string' || pathRaw.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'path or documentId is required.');
    }

    const { filePath, sourceFormat } = await assertSupportedPath(pathRaw);
    const analysis = await analyzeReferenceTemplate(filePath, sourceFormat);
    sendSuccess(res, {
      analysis,
      warnings: analysis.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/analyze-repeat-region', async (req, res) => {
  try {
    const documentId = req.body?.documentId;
    if (typeof documentId !== 'string' || documentId.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'documentId is required.');
    }
    if (typeof req.body?.tableBlockId !== 'string' || req.body.tableBlockId.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'tableBlockId is required.');
    }
    if (!Number.isInteger(req.body?.templateRowIndex)) {
      throw new RequestError(400, 'INVALID_REQUEST', 'templateRowIndex must be an integer.');
    }
    if (req.body?.templateEndRowIndex !== undefined && !Number.isInteger(req.body.templateEndRowIndex)) {
      throw new RequestError(400, 'INVALID_REQUEST', 'templateEndRowIndex must be an integer when provided.');
    }

    const session = requireSession(documentId);
    const templateRowIndex = req.body.templateRowIndex as number;
    const templateEndRowIndex = req.body.templateEndRowIndex as number | undefined;
    if (templateEndRowIndex !== undefined && templateEndRowIndex < templateRowIndex) {
      throw new RequestError(400, 'INVALID_REQUEST', 'templateEndRowIndex must be greater than or equal to templateRowIndex.');
    }

    const analyzed = analyzeTableRepeatRegion(session.body, {
      tableBlockId: req.body.tableBlockId,
      templateRowIndex,
      templateEndRowIndex,
    });
    if (!analyzed) {
      throw new RequestError(404, 'FILE_NOT_FOUND', `No table matched blockId "${String(req.body.tableBlockId)}".`);
    }

    if (analyzed.analysis.templateEndRowIndex >= analyzed.analysis.tableRowCount) {
      throw new RequestError(
        400,
        'INVALID_REQUEST',
        `Template region ${templateRowIndex}-${analyzed.analysis.templateEndRowIndex} is outside the table row range 0-${analyzed.analysis.tableRowCount - 1}.`
      );
    }

    sendSuccess(res, {
      repeatRegion: analyzed.analysis,
      warnings: session.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/create', async (req, res) => {
  try {
    const format = normalizeTargetFormat(String(req.body?.format || ''));
    const outputPathRaw = req.body?.outputPath;
    const templateIdRaw = req.body?.templateId;
    const referencePathRaw = req.body?.referencePath;

    if (format !== 'hwpx') {
      throw new RequestError(400, 'INVALID_REQUEST', 'Only format="hwpx" is supported for document creation in v1.');
    }

    if (typeof outputPathRaw !== 'string' || outputPathRaw.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'outputPath is required.');
    }

    if (templateIdRaw !== undefined && referencePathRaw !== undefined) {
      throw new RequestError(400, 'INVALID_REQUEST', 'Specify either templateId or referencePath, not both.');
    }

    if (templateIdRaw !== undefined && !isDocumentTemplateId(templateIdRaw)) {
      throw new RequestError(400, 'INVALID_REQUEST', `Unsupported templateId: ${String(templateIdRaw)}`, {
        supportedTemplates: listDocumentTemplateIds(),
      });
    }

    const outputPath = normalizeAbsolutePath(outputPathRaw);
    let document;

    if (typeof referencePathRaw === 'string' && referencePathRaw.trim().length > 0) {
      const { filePath: referencePath, sourceFormat: referenceFormat } = await assertSupportedPath(referencePathRaw);
      if (referencePath === outputPath) {
        throw new RequestError(400, 'INVALID_REQUEST', 'outputPath must differ from referencePath.');
      }
      document = await createHwpxDocumentFromReferenceCore(referencePath, referenceFormat, outputPath);
    } else {
      document = await createHwpxDocumentCore(outputPath, templateIdRaw ?? 'blank');
    }

    const session = createDocumentSession(outputPath, document, { workingCopyPath: outputPath });
    sessionStore.put(session);

    sendSuccess(res, {
      document: serializeSessionDocument(session),
      warnings: session.warnings,
    }, 201);
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/fork-editable-copy', async (req, res) => {
  try {
    const session = requireSession(req.body?.documentId);
    const outputPathRaw = typeof req.body?.outputPath === 'string' ? req.body.outputPath : undefined;
    const { targetPath, document, warnings } = await createEditableWorkingCopy(session, outputPathRaw);

    const updated = sessionStore.update(session.documentId, {
      workingCopyPath: targetPath,
      documentMode: 'editable',
      readOnly: false,
      title: document.title,
      projectionMode: document.projectionMode,
      body: document.body,
      supplementalText: document.supplementalText,
      imageReferences: document.imageReferences,
      paragraphs: document.paragraphs,
      rawText: document.rawText,
      warnings,
      features: document.features,
      currentCheckpointId: null,
    });

    if (!updated) {
      throw new RequestError(500, 'INTERNAL_ERROR', 'Failed to update document session after creating working copy.');
    }

    sendSuccess(res, {
      document: serializeSessionDocument(updated),
      warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.get('/document/checkpoints', (req, res) => {
  try {
    const session = requireSession(req.query.documentId);
    const checkpoints = checkpointStore.list(session.documentId);
    sendSuccess(res, { checkpoints });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/recover', async (req, res) => {
  try {
    const session = requireSession(req.body?.documentId);
    const checkpointId = req.body?.checkpointId;
    if (typeof checkpointId !== 'string' || checkpointId.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'checkpointId is required.');
    }

    if (!session.workingCopyPath) {
      throw new RequestError(400, 'READ_ONLY_SOURCE', 'This document has no editable working copy to recover into.');
    }

    const restored = await checkpointStore.restore(session.documentId, checkpointId, session.workingCopyPath);
    if (!restored) {
      throw new RequestError(404, 'CHECKPOINT_NOT_FOUND', `Checkpoint not found: ${checkpointId}`, { checkpointId });
    }

    const reopened = await openDocumentCore(session.workingCopyPath, 'hwpx');
    const updated = sessionStore.update(session.documentId, {
      documentMode: 'editable',
      readOnly: false,
      title: reopened.title,
      projectionMode: reopened.projectionMode,
      body: reopened.body,
      supplementalText: reopened.supplementalText,
      imageReferences: reopened.imageReferences,
      paragraphs: reopened.paragraphs,
      rawText: reopened.rawText,
      warnings: reopened.warnings,
      features: reopened.features,
      currentCheckpointId: restored.checkpointId,
    });

    if (!updated) {
      throw new RequestError(500, 'INTERNAL_ERROR', 'Failed to update document session after recovery.');
    }

    sendSuccess(res, {
      restored,
      document: serializeSessionDocument(updated),
      warnings: updated.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/save', async (req, res) => {
  try {
    const documentId = req.body?.documentId;
    if (typeof documentId !== 'string' || documentId.trim().length === 0) {
      throw new RequestError(
        400,
        'INVALID_REQUEST',
        'documentId is required. Open the document through /document/open and save via the returned session.'
      );
    }

    const session = requireSession(documentId);
    if (session.readOnly || session.documentMode !== 'editable' || !session.workingCopyPath) {
      throw new RequestError(400, 'READ_ONLY_SOURCE', 'This document is read-only. Create an editable .hwpx working copy first.');
    }

    const payload = savePayloadForSession(session, {
      text: req.body?.text,
      mutations: req.body?.mutations,
    });
    if (payload.mutations) {
      await validateImageMutations(session, payload.mutations);
    }
    const checkpoint = await checkpointStore.create(session.documentId, session.workingCopyPath);
    const saved = await saveEditableDocumentCore(session, payload);

    const updated = sessionStore.update(session.documentId, {
      documentMode: 'editable',
      readOnly: false,
      title: saved.title,
      projectionMode: saved.projectionMode,
      body: saved.body,
      supplementalText: saved.supplementalText,
      imageReferences: saved.imageReferences,
      paragraphs: saved.paragraphs,
      rawText: saved.rawText,
      warnings: saved.warnings,
      features: saved.features,
      currentCheckpointId: checkpoint?.checkpointId ?? session.currentCheckpointId,
    });

    if (!updated) {
      throw new RequestError(500, 'INTERNAL_ERROR', 'Failed to update document session after save.');
    }

    sendSuccess(res, {
      document: serializeSessionDocument(updated),
      save: {
        checkpointId: checkpoint?.checkpointId ?? null,
        validationSummary: {
          reopenVerified: true,
          projectionMode: updated.projectionMode,
          allowPlainTextSave: updated.features.allowPlainTextSave,
        },
      },
      warnings: updated.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/fill-template', async (req, res) => {
  try {
    const documentId = req.body?.documentId;
    if (typeof documentId !== 'string' || documentId.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'documentId is required.');
    }

    const session = requireSession(documentId);
    if (session.readOnly || session.documentMode !== 'editable' || !session.workingCopyPath) {
      throw new RequestError(400, 'READ_ONLY_SOURCE', 'This document is read-only. Create an editable .hwpx working copy first.');
    }

    let fillInput;
    try {
      fillInput = normalizeTemplateFillInstructions({
        fills: req.body?.fills,
        values: req.body?.values,
        tableRepeats: req.body?.tableRepeats,
        requiredPlaceholders: req.body?.requiredPlaceholders,
        requiredTargets: req.body?.requiredTargets,
      });
    } catch (error) {
      throw new RequestError(400, 'INVALID_REQUEST', error instanceof Error ? error.message : 'Invalid fill payload.');
    }

    if (fillInput.fills.length === 0 && fillInput.tableRepeats.length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'fills, values, or tableRepeats must include at least one replacement instruction.');
    }

    const validationIssues = validateTemplateFillRequirements(session.body, session.supplementalText, {
      requiredPlaceholders: fillInput.requiredPlaceholders,
      requiredTargets: fillInput.requiredTargets,
    });
    if (validationIssues.length > 0) {
      throw new RequestError(
        400,
        'VALIDATION_FAILED',
        'Template binding validation failed.',
        {
          issues: validationIssues,
        },
        false,
        validationIssues
      );
    }

    const { mutations, result } = planTemplateFill(session.body, session.supplementalText, fillInput);
    if (mutations.length === 0) {
      throw new RequestError(
        400,
        'INVALID_REQUEST',
        result.warnings.length > 0
          ? 'No deterministic fill or repeat operation could be applied to the current document.'
          : 'No placeholder or target matched the current document.',
        undefined,
        false,
        result.warnings
      );
    }

    const checkpoint = await checkpointStore.create(session.documentId, session.workingCopyPath);
    const saved = await saveEditableDocumentCore(session, { mutations });

    const updated = sessionStore.update(session.documentId, {
      documentMode: 'editable',
      readOnly: false,
      title: saved.title,
      projectionMode: saved.projectionMode,
      body: saved.body,
      supplementalText: saved.supplementalText,
      imageReferences: saved.imageReferences,
      paragraphs: saved.paragraphs,
      rawText: saved.rawText,
      warnings: [...saved.warnings, ...result.warnings],
      features: saved.features,
      currentCheckpointId: checkpoint?.checkpointId ?? session.currentCheckpointId,
    });

    if (!updated) {
      throw new RequestError(500, 'INTERNAL_ERROR', 'Failed to update document session after template fill.');
    }

    sendSuccess(res, {
      document: serializeSessionDocument(updated),
      fill: {
        ...result,
        checkpointId: checkpoint?.checkpointId ?? null,
      },
      warnings: updated.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/replace-image', async (req, res) => {
  try {
    const documentId = req.body?.documentId;
    if (typeof documentId !== 'string' || documentId.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'documentId is required.');
    }

    const imagePathRaw = req.body?.imagePath;
    if (typeof imagePathRaw !== 'string' || imagePathRaw.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'imagePath is required.');
    }

    const targetId = typeof req.body?.targetId === 'string' && req.body.targetId.trim().length > 0
      ? req.body.targetId.trim()
      : undefined;
    const binaryItemId = typeof req.body?.binaryItemId === 'string' && req.body.binaryItemId.trim().length > 0
      ? req.body.binaryItemId.trim()
      : undefined;

    if ((!targetId && !binaryItemId) || (targetId && binaryItemId)) {
      throw new RequestError(400, 'INVALID_REQUEST', 'Specify exactly one of targetId or binaryItemId.');
    }

    const session = requireSession(documentId);
    if (session.readOnly || session.documentMode !== 'editable' || !session.workingCopyPath) {
      throw new RequestError(400, 'READ_ONLY_SOURCE', 'This document is read-only. Create an editable .hwpx working copy first.');
    }
    const mutation: ImageReplacementMutation = {
      op: 'replace_image_asset',
      imagePath: imagePathRaw,
      targetId,
      binaryItemId,
    };
    await validateImageMutations(session, [mutation]);
    const impact = resolveAssetReplacementImpact(session, mutation);
    const operationWarnings: ApiWarning[] = impact.references.length > 1
      ? [{
          code: 'IMAGE_ASSET_SHARED_REPLACED',
          message: `This asset replacement affects ${impact.references.length} image references that currently share binaryItemId "${impact.binaryItemId}".`,
        }]
      : [];

    const checkpoint = await checkpointStore.create(session.documentId, session.workingCopyPath);
    const saved = await saveEditableDocumentCore(session, {
      mutations: [mutation],
    });

    const updated = sessionStore.update(session.documentId, {
      documentMode: 'editable',
      readOnly: false,
      title: saved.title,
      projectionMode: saved.projectionMode,
      body: saved.body,
      supplementalText: saved.supplementalText,
      imageReferences: saved.imageReferences,
      paragraphs: saved.paragraphs,
      rawText: saved.rawText,
      warnings: saved.warnings,
      features: saved.features,
      currentCheckpointId: checkpoint?.checkpointId ?? session.currentCheckpointId,
    });

    if (!updated) {
      throw new RequestError(500, 'INTERNAL_ERROR', 'Failed to update document session after image replacement.');
    }

    sendSuccess(res, {
      document: serializeSessionDocument(updated),
      imageReplacement: {
        checkpointId: checkpoint?.checkpointId ?? null,
        targetId: targetId ?? null,
        binaryItemId: binaryItemId ?? null,
        imagePath: mutation.imagePath,
        affectedReferenceCount: impact.references.length,
        affectedTargetIds: [...new Set(impact.references.map((reference) => reference.targetId))],
      },
      warnings: [...updated.warnings, ...operationWarnings],
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/replace-image-instance', async (req, res) => {
  try {
    const documentId = req.body?.documentId;
    if (typeof documentId !== 'string' || documentId.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'documentId is required.');
    }

    const session = requireSession(documentId);
    if (session.readOnly || session.documentMode !== 'editable' || !session.workingCopyPath) {
      throw new RequestError(400, 'READ_ONLY_SOURCE', 'This document is read-only. Create an editable .hwpx working copy first.');
    }

    const mutation: ImageInstanceReplacementMutation = {
      op: 'replace_image_instance_asset',
      imagePath: req.body?.imagePath,
      targetId: req.body?.targetId,
      binaryItemId: req.body?.binaryItemId,
      objectId: req.body?.objectId,
      instanceId: req.body?.instanceId,
    };
    await validateImageMutations(session, [mutation]);
    const previousReference = mutation.objectId
      ? session.imageReferences.find((reference) => reference.placement?.objectId === mutation.objectId)
      : mutation.instanceId
        ? session.imageReferences.find((reference) => reference.placement?.instanceId === mutation.instanceId)
        : mutation.targetId
          ? (() => {
              const refs = session.imageReferences.filter((reference) => reference.targetId === mutation.targetId);
              return refs.length === 1 ? refs[0] : undefined;
            })()
          : session.imageReferences.find((reference) => reference.binaryItemId === mutation.binaryItemId);

    const checkpoint = await checkpointStore.create(session.documentId, session.workingCopyPath);
    const saved = await saveEditableDocumentCore(session, {
      mutations: [mutation],
    });

    const updated = sessionStore.update(session.documentId, {
      documentMode: 'editable',
      readOnly: false,
      title: saved.title,
      projectionMode: saved.projectionMode,
      body: saved.body,
      supplementalText: saved.supplementalText,
      imageReferences: saved.imageReferences,
      paragraphs: saved.paragraphs,
      rawText: saved.rawText,
      warnings: saved.warnings,
      features: saved.features,
      currentCheckpointId: checkpoint?.checkpointId ?? session.currentCheckpointId,
    });

    if (!updated) {
      throw new RequestError(500, 'INTERNAL_ERROR', 'Failed to update document session after image-instance replacement.');
    }

    sendSuccess(res, {
      document: serializeSessionDocument(updated),
      imageInstanceReplacement: {
        checkpointId: checkpoint?.checkpointId ?? null,
        targetId: mutation.targetId ?? null,
        binaryItemId: mutation.binaryItemId ?? null,
        objectId: mutation.objectId ?? null,
        instanceId: mutation.instanceId ?? null,
        imagePath: mutation.imagePath,
        previousBinaryItemId: previousReference?.binaryItemId ?? null,
        nextBinaryItemId: mutation.objectId
          ? saved.imageReferences.find((reference) => reference.placement?.objectId === mutation.objectId)?.binaryItemId ?? null
          : mutation.instanceId
            ? saved.imageReferences.find((reference) => reference.placement?.instanceId === mutation.instanceId)?.binaryItemId ?? null
            : null,
      },
      warnings: updated.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/delete-image', async (req, res) => {
  try {
    const documentId = req.body?.documentId;
    if (typeof documentId !== 'string' || documentId.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'documentId is required.');
    }

    const session = requireSession(documentId);
    if (session.readOnly || session.documentMode !== 'editable' || !session.workingCopyPath) {
      throw new RequestError(400, 'READ_ONLY_SOURCE', 'This document is read-only. Create an editable .hwpx working copy first.');
    }

    const mutation: ImageDeletionMutation = {
      op: 'delete_image_instance',
      targetId: req.body?.targetId,
      binaryItemId: req.body?.binaryItemId,
      objectId: req.body?.objectId,
      instanceId: req.body?.instanceId,
    };
    await validateImageMutations(session, [mutation]);

    const previousReference = mutation.objectId
      ? session.imageReferences.find((reference) => reference.placement?.objectId === mutation.objectId)
      : mutation.instanceId
        ? session.imageReferences.find((reference) => reference.placement?.instanceId === mutation.instanceId)
        : mutation.targetId
          ? (() => {
              const refs = session.imageReferences.filter((reference) => reference.targetId === mutation.targetId);
              return refs.length === 1 ? refs[0] : undefined;
            })()
          : session.imageReferences.find((reference) => reference.binaryItemId === mutation.binaryItemId);

    const checkpoint = await checkpointStore.create(session.documentId, session.workingCopyPath);
    const saved = await saveEditableDocumentCore(session, {
      mutations: [mutation],
    });

    const updated = sessionStore.update(session.documentId, {
      documentMode: 'editable',
      readOnly: false,
      title: saved.title,
      projectionMode: saved.projectionMode,
      body: saved.body,
      supplementalText: saved.supplementalText,
      imageReferences: saved.imageReferences,
      paragraphs: saved.paragraphs,
      rawText: saved.rawText,
      warnings: saved.warnings,
      features: saved.features,
      currentCheckpointId: checkpoint?.checkpointId ?? session.currentCheckpointId,
    });

    if (!updated) {
      throw new RequestError(500, 'INTERNAL_ERROR', 'Failed to update document session after image deletion.');
    }

    sendSuccess(res, {
      document: serializeSessionDocument(updated),
      imageDeletion: {
        checkpointId: checkpoint?.checkpointId ?? null,
        targetId: mutation.targetId ?? null,
        binaryItemId: mutation.binaryItemId ?? null,
        objectId: mutation.objectId ?? null,
        instanceId: mutation.instanceId ?? null,
        deletedBinaryItemId: previousReference?.binaryItemId ?? null,
        deletedObjectId: previousReference?.placement?.objectId ?? null,
        deletedInstanceId: previousReference?.placement?.instanceId ?? null,
      },
      warnings: updated.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/insert-image', async (req, res) => {
  try {
    const documentId = req.body?.documentId;
    if (typeof documentId !== 'string' || documentId.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'documentId is required.');
    }

    const session = requireSession(documentId);
    if (session.readOnly || session.documentMode !== 'editable' || !session.workingCopyPath) {
      throw new RequestError(400, 'READ_ONLY_SOURCE', 'This document is read-only. Create an editable .hwpx working copy first.');
    }

    const mutation: ImageInsertionMutation = {
      op: 'insert_image_from_prototype',
      imagePath: req.body?.imagePath,
      destinationTargetId: req.body?.destinationTargetId,
      prototypeTargetId: req.body?.prototypeTargetId,
      prototypeBinaryItemId: req.body?.prototypeBinaryItemId,
      prototypeObjectId: req.body?.prototypeObjectId,
      prototypeInstanceId: req.body?.prototypeInstanceId,
      patch: req.body?.patch ?? undefined,
    };
    await validateImageMutations(session, [mutation]);

    const checkpoint = await checkpointStore.create(session.documentId, session.workingCopyPath);
    const saved = await saveEditableDocumentCore(session, {
      mutations: [mutation],
    });

    const updated = sessionStore.update(session.documentId, {
      documentMode: 'editable',
      readOnly: false,
      title: saved.title,
      projectionMode: saved.projectionMode,
      body: saved.body,
      supplementalText: saved.supplementalText,
      imageReferences: saved.imageReferences,
      paragraphs: saved.paragraphs,
      rawText: saved.rawText,
      warnings: saved.warnings,
      features: saved.features,
      currentCheckpointId: checkpoint?.checkpointId ?? session.currentCheckpointId,
    });

    if (!updated) {
      throw new RequestError(500, 'INTERNAL_ERROR', 'Failed to update document session after image insertion.');
    }

    const previousKeys = new Set(session.imageReferences.map((reference) => [
      reference.targetId,
      reference.binaryItemId,
      reference.placement?.objectId ?? '',
      reference.placement?.instanceId ?? '',
    ].join('|')));
    const insertedReference = saved.imageReferences.find((reference) =>
      reference.targetId === mutation.destinationTargetId
      && !previousKeys.has([
        reference.targetId,
        reference.binaryItemId,
        reference.placement?.objectId ?? '',
        reference.placement?.instanceId ?? '',
      ].join('|'))
    ) ?? null;

    sendSuccess(res, {
      document: serializeSessionDocument(updated),
      imageInsertion: {
        checkpointId: checkpoint?.checkpointId ?? null,
        destinationTargetId: mutation.destinationTargetId,
        prototypeTargetId: mutation.prototypeTargetId ?? null,
        prototypeBinaryItemId: mutation.prototypeBinaryItemId ?? null,
        prototypeObjectId: mutation.prototypeObjectId ?? null,
        prototypeInstanceId: mutation.prototypeInstanceId ?? null,
        imagePath: mutation.imagePath,
        patch: mutation.patch ?? null,
        insertedBinaryItemId: insertedReference?.binaryItemId ?? null,
        insertedObjectId: insertedReference?.placement?.objectId ?? null,
        insertedInstanceId: insertedReference?.placement?.instanceId ?? null,
      },
      warnings: updated.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/document/update-image-placement', async (req, res) => {
  try {
    const documentId = req.body?.documentId;
    if (typeof documentId !== 'string' || documentId.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'documentId is required.');
    }

    const session = requireSession(documentId);
    if (session.readOnly || session.documentMode !== 'editable' || !session.workingCopyPath) {
      throw new RequestError(400, 'READ_ONLY_SOURCE', 'This document is read-only. Create an editable .hwpx working copy first.');
    }

    const mutation: ImagePlacementMutation = {
      op: 'update_image_placement',
      targetId: req.body?.targetId,
      binaryItemId: req.body?.binaryItemId,
      objectId: req.body?.objectId,
      instanceId: req.body?.instanceId,
      patch: req.body?.patch ?? {},
    };
    await validateImageMutations(session, [mutation]);

    const checkpoint = await checkpointStore.create(session.documentId, session.workingCopyPath);
    const saved = await saveEditableDocumentCore(session, {
      mutations: [mutation],
    });

    const updated = sessionStore.update(session.documentId, {
      documentMode: 'editable',
      readOnly: false,
      title: saved.title,
      projectionMode: saved.projectionMode,
      body: saved.body,
      supplementalText: saved.supplementalText,
      imageReferences: saved.imageReferences,
      paragraphs: saved.paragraphs,
      rawText: saved.rawText,
      warnings: saved.warnings,
      features: saved.features,
      currentCheckpointId: checkpoint?.checkpointId ?? session.currentCheckpointId,
    });

    if (!updated) {
      throw new RequestError(500, 'INTERNAL_ERROR', 'Failed to update document session after image placement update.');
    }

    sendSuccess(res, {
      document: serializeSessionDocument(updated),
      imagePlacementUpdate: {
        checkpointId: checkpoint?.checkpointId ?? null,
        targetId: mutation.targetId ?? null,
        binaryItemId: mutation.binaryItemId ?? null,
        objectId: mutation.objectId ?? null,
        instanceId: mutation.instanceId ?? null,
        patch: mutation.patch,
      },
      warnings: updated.warnings,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.post('/converter', async (req, res) => {
  const request = req.body as Partial<OfficeExtensionConvertRequest>;

  try {
    if (!request.filetype || !request.outputtype) {
      throw new RequestError(400, 'INVALID_REQUEST', 'filetype and outputtype are required.');
    }

    const markdownMode = normalizeMarkdownMode(request.markdownMode);
    const includeDiagnostics = typeof request.includeDiagnostics === 'boolean'
      ? request.includeDiagnostics
      : undefined;

    const sourcePath = await resolveSourceFile(request as OfficeExtensionConvertRequest);
    const sourceFormat = guessSourceFormat(request.filetype, sourcePath);
    if (!sourceFormat) {
      throw new RequestError(400, 'UNSUPPORTED_EXTENSION', `Unsupported source format: ${request.filetype}. Only hwp/hwpx are supported.`);
    }

    const fileStat = await statSafe(sourcePath);
    if (fileStat && fileStat.size > maxDocumentBytes) {
      throw new RequestError(413, 'DOCUMENT_TOO_LARGE', `Document exceeds the size limit (${maxDocumentBytes} bytes).`, {
        path: sourcePath,
        size: fileStat.size,
        limit: maxDocumentBytes,
      });
    }

    const target = normalizeTargetFormat(request.outputtype);
    if (!SUPPORTED_TARGETS.has(target as TargetFormat)) {
      throw new RequestError(400, 'INVALID_REQUEST', `Unsupported target format: ${request.outputtype}. Supported: ${Array.from(SUPPORTED_TARGETS).join(', ')}`);
    }

    const result = await convertWithSource({
      sourcePath,
      sourceFormat,
      targetFormat: target as TargetFormat,
      requestedOutputPath: request.outputPath,
      options: {
        markdownMode,
        includeDiagnostics,
      },
    });

    const conversionPayload = {
      outputPath: result.outputPath,
      sidecarPaths: result.sidecarPaths,
      details: result.details,
      warnings: result.warnings ?? [],
    };

    if (request.outputPath && request.outputPath.trim().length > 0) {
      sendSuccess(res, conversionPayload);
      return;
    }

    const token = nanoid(16);
    fileStore.put(token, result.outputPath);
    sendSuccess(res, {
      ...conversionPayload,
      url: `${baseUrl(req)}/download/${token}`,
    });
  } catch (error) {
    sendError(res, error);
  }
});

app.get('/download/:token', async (req, res) => {
  try {
    const token = req.params.token;
    const stored = await fileStore.consume(token);

    if (!stored) {
      throw new RequestError(404, 'FILE_NOT_FOUND', 'File token expired or does not exist.', { token });
    }

    res.status(200);
    res.setHeader('Content-Type', downloadContentType(stored.path));
    res.setHeader('Content-Disposition', `attachment; filename="${stored.path.split('/').pop() || 'converted'}"`);

    stored.stream.on('close', async () => {
      try {
        await unlink(stored.path);
      } catch {
        // Ignore cleanup errors
      }
    });

    stored.stream.pipe(res);
  } catch (error) {
    sendError(res, error);
  }
});

app.get('/open', async (req, res) => {
  try {
    const filePathRaw = req.query.filepath;
    if (typeof filePathRaw !== 'string' || filePathRaw.trim().length === 0) {
      throw new RequestError(400, 'INVALID_REQUEST', 'Missing filepath query parameter.');
    }

    const { filePath } = await assertSupportedPath(filePathRaw);
    res.status(200).setHeader('Content-Type', 'text/html; charset=utf-8').send(renderOpenPage(filePath));
  } catch (error) {
    if (error instanceof RequestError) {
      res.status(error.status).send(error.message);
      return;
    }
    res.status(500).send(error instanceof Error ? error.message : 'Failed to open document.');
  }
});

setInterval(() => {
  void fileStore.cleanupExpired();
}, 60_000).unref();

app.listen(port, () => {
  console.log(`[hwp-converter-extension] listening on http://localhost:${port} (engine=${engineVersion})`);
});
