export interface OfficeExtensionConvertRequest {
  filetype: string;
  outputtype: string;
  key?: string;
  title?: string;
  filePath?: string;
  url?: string;
  outputPath?: string;
  markdownMode?: MarkdownMode;
  includeDiagnostics?: boolean;
}

export interface OfficeExtensionConvertResponse {
  success: boolean;
  url?: string;
  outputPath?: string;
  details?: string;
  error?: string;
}

export type SourceFormat = 'hwp' | 'hwpx';

export type TargetFormat =
  | 'txt'
  | 'md'
  | 'html'
  | 'json'
  | 'docx'
  | 'pdf';

export type MarkdownMode = 'clean' | 'fidelity';

export interface ConvertOptions {
  markdownMode?: MarkdownMode;
  includeDiagnostics?: boolean;
}

export interface ConvertContext {
  sourcePath: string;
  sourceFormat: SourceFormat;
  targetFormat: TargetFormat;
  requestedOutputPath?: string;
  options?: ConvertOptions;
}

export interface ConvertedArtifact {
  outputPath: string;
  targetFormat: TargetFormat;
  sidecarPaths?: string[];
  warnings?: ApiWarning[];
  details?: Record<string, unknown>;
}

export interface TextExtractionResult {
  paragraphs: string[];
  rawText: string;
}

export type DocumentMode = 'read_only' | 'editable' | 'import_required';

export type ProjectionMode = 'text_projection' | 'structured_hwpx';

export type ErrorCode =
  | 'INVALID_REQUEST'
  | 'FILE_NOT_FOUND'
  | 'IMAGE_TARGET_NOT_FOUND'
  | 'IMAGE_TARGET_AMBIGUOUS'
  | 'UNSUPPORTED_EXTENSION'
  | 'DOCUMENT_TOO_LARGE'
  | 'ENGINE_TIMEOUT'
  | 'PARSE_FAILED'
  | 'IMPORT_FAILED'
  | 'SAVE_FAILED'
  | 'VALIDATION_FAILED'
  | 'READ_ONLY_SOURCE'
  | 'CHECKPOINT_NOT_FOUND'
  | 'RECOVERY_FAILED'
  | 'INTERNAL_ERROR';

export interface ApiWarning {
  code: string;
  message: string;
}

export interface ApiError {
  code: ErrorCode;
  message: string;
  details?: Record<string, unknown>;
  retryable: boolean;
}

export interface DocumentFeatureSummary {
  projectionMode: ProjectionMode;
  authoritative: boolean;
  editable: string[];
  preservedReadOnly: string[];
  unsupported: string[];
  hasUnsupportedEditableFeatures: boolean;
  allowPlainTextSave: boolean;
}

export interface ParagraphBlock {
  type: 'paragraph';
  blockId: string;
  text: string;
  styleRef: string | null;
  editable: boolean;
  containsObjects: boolean;
}

export interface TableCellBlock {
  cellId: string;
  rowIndex: number;
  columnIndex: number;
  rowSpan: number;
  colSpan: number;
  text: string;
  paragraphs: string[];
  editable: boolean;
}

export interface TableBlock {
  type: 'table';
  blockId: string;
  tableId: string | null;
  rowCount: number;
  columnCount: number;
  editable: boolean;
  cells: TableCellBlock[];
}

export type DocumentBlock = ParagraphBlock | TableBlock;

export interface DocumentSection {
  sectionId: string;
  blocks: DocumentBlock[];
}

export interface NormalizedDocumentBody {
  sections: DocumentSection[];
}

export type SupplementalTextKind = 'header' | 'footer';
export type EmbeddedImageKind = 'paragraph' | 'table_cell' | 'header_paragraph' | 'footer_paragraph';

export interface SupplementalTextBlock {
  kind: SupplementalTextKind;
  targetId: string;
  paragraphIndex: number;
  text: string;
  editable: boolean;
  containsObjects: boolean;
}

export interface EmbeddedImageReference {
  kind: EmbeddedImageKind;
  targetId: string;
  containerId: string;
  sectionIndex: number;
  blockIndex: number;
  rowIndex?: number;
  columnIndex?: number;
  paragraphIndex?: number;
  binaryItemId: string;
  assetPath: string | null;
  assetFileName: string | null;
  mediaType: string | null;
  width?: number | null;
  height?: number | null;
  placement?: EmbeddedImagePlacement | null;
}

export interface EmbeddedImagePlacement {
  objectId: string | null;
  instanceId: string | null;
  zOrder: number | null;
  textWrap: string | null;
  textFlow: string | null;
  numberingType: string | null;
  widthRelTo: string | null;
  heightRelTo: string | null;
  margins: {
    left: number | null;
    right: number | null;
    top: number | null;
    bottom: number | null;
  };
  clip: {
    left: number | null;
    right: number | null;
    top: number | null;
    bottom: number | null;
  };
}

export interface ImageTargetSummary {
  targetId: string;
  containerId: string;
  kind: EmbeddedImageKind;
  sectionIndex: number;
  blockIndex: number;
  rowIndex?: number;
  columnIndex?: number;
  paragraphIndex?: number;
  imageCount: number;
  deterministicTargetReplacement: boolean;
  recommendedLocator: 'targetId' | 'binaryItemId';
  recommendedAssetLocator: 'targetId' | 'binaryItemId';
  recommendedPlacementLocator: 'targetId' | 'binaryItemId' | 'objectId' | 'instanceId' | 'none';
  binaryItemIds: string[];
  objectIds: string[];
  instanceIds: string[];
  assetFileNames: string[];
}

export interface ImageResourceSummary {
  binaryItemId: string;
  assetPath: string | null;
  assetFileName: string | null;
  mediaType: string | null;
  referenceCount: number;
  targetIds: string[];
  sharedAcrossTargets: boolean;
  assetReplacementScope: 'single_target' | 'multi_target';
  deterministicPlacementUpdate: boolean;
}

export interface DocumentSession {
  documentId: string;
  sourcePath: string;
  sourceFormat: SourceFormat;
  canonicalFormat: 'hwpx';
  workingCopyPath: string | null;
  documentMode: DocumentMode;
  readOnly: boolean;
  title: string;
  projectionMode: ProjectionMode;
  body: NormalizedDocumentBody;
  supplementalText: SupplementalTextBlock[];
  imageReferences: EmbeddedImageReference[];
  paragraphs: string[];
  rawText: string;
  warnings: ApiWarning[];
  features: DocumentFeatureSummary;
  currentCheckpointId: string | null;
  createdAt: string;
  updatedAt: string;
}

export type DocumentTemplateId = 'blank' | 'report' | 'minutes' | 'proposal';

export interface ReferenceStyleSummary {
  styleRef: string | null;
  paragraphCount: number;
  sampleText: string | null;
}

export interface ReferenceTableSummary {
  blockId: string;
  rowCount: number;
  columnCount: number;
  mergedCellCount: number;
  firstRowPreview: string[];
}

export interface ReferenceRepeatableRegionSummary {
  tableBlockId: string;
  startRowIndex: number;
  endRowIndex: number;
  rowCount: number;
  columnCount: number;
  hasMergedCells: boolean;
  boundaryCount: number;
  supportedBoundaryPolicies: Array<'reject' | 'split_boundary_merges'>;
  recommendedBoundaryPolicy: 'reject' | 'split_boundary_merges';
}

export interface TableRepeatBoundaryCrossingSummary {
  rowIndex: number;
  columnIndex: number;
  rowSpan: number;
  colSpan: number;
  coversWholeRegion: boolean;
  textPreview: string;
}

export interface TableRepeatRegionAnalysis {
  tableBlockId: string;
  templateRowIndex: number;
  templateEndRowIndex: number;
  templateRowCount: number;
  tableRowCount: number;
  strictSafe: boolean;
  splitSafe: boolean;
  supportedBoundaryPolicies: Array<'reject' | 'split_boundary_merges'>;
  recommendedBoundaryPolicy: 'reject' | 'split_boundary_merges';
  topBoundaryCrossings: TableRepeatBoundaryCrossingSummary[];
  bottomBoundaryCrossings: TableRepeatBoundaryCrossingSummary[];
  internalMergedCellCount: number;
}

export interface ReferencePlaceholderSummary {
  kind: 'paragraph' | 'table_cell' | 'header_paragraph' | 'footer_paragraph';
  blockId: string;
  text: string;
}

export interface ReferenceTokenOccurrence {
  kind: 'paragraph' | 'table_cell' | 'header_paragraph' | 'footer_paragraph';
  targetId: string;
  containerId: string;
  sectionIndex: number;
  blockIndex: number;
  rowIndex?: number;
  columnIndex?: number;
  textPreview: string;
}

export interface ReferenceTokenSummary {
  token: string;
  occurrenceCount: number;
  ambiguous: boolean;
  occurrences: ReferenceTokenOccurrence[];
}

export interface ReferenceTemplateAnalysis {
  sourcePath: string;
  sourceFormat: SourceFormat;
  canonicalFormat: 'hwpx';
  importedFromHwp: boolean;
  title: string;
  sectionCount: number;
  paragraphCount: number;
  tableCount: number;
  assetCount: number;
  unsupportedFeatures: string[];
  packageValidation: {
    sectionEntries: string[];
    assetEntries: string[];
    missingRecommendedEntries: string[];
  };
  headerPreview: string[];
  footerPreview: string[];
  styles: ReferenceStyleSummary[];
  tables: ReferenceTableSummary[];
  repeatableRegions: ReferenceRepeatableRegionSummary[];
  repeatRegionRecommendations: TableRepeatRegionAnalysis[];
  placeholders: ReferencePlaceholderSummary[];
  imageReferences: EmbeddedImageReference[];
  imageTargets: ImageTargetSummary[];
  imageResources: ImageResourceSummary[];
  detectedTokens: string[];
  tokenRegistry: ReferenceTokenSummary[];
  textPreview: string[];
  warnings: ApiWarning[];
}

export interface TemplateFillInstruction {
  target?: string;
  placeholder?: string;
  value: string;
  allowMultipleMatches?: boolean;
}

export interface TemplateFillRequirements {
  requiredPlaceholders: string[];
  requiredTargets: string[];
}

export interface TableRepeatInstruction {
  tableBlockId: string;
  templateRowIndex: number;
  templateEndRowIndex?: number;
  rows: Array<string[] | Record<string, string>>;
  boundaryPolicy?: 'reject' | 'split_boundary_merges';
}

export interface TemplateFillMatch {
  instructionIndex: number;
  target: string;
  matchCount: number;
}

export interface TableRepeatMatch {
  tableBlockId: string;
  templateRowIndex: number;
  templateEndRowIndex?: number;
  templateRowCount: number;
  repeatedRegionCount: number;
  appliedRowCount: number;
  insertedRowCount: number;
}

export interface TemplateFillResult {
  applied: TemplateFillMatch[];
  repeatedTables: TableRepeatMatch[];
  warnings: ApiWarning[];
  mutationCount: number;
}

export interface CheckpointRecord {
  checkpointId: string;
  documentId: string;
  path: string;
  createdAt: string;
}

export interface ImagePlacementPatch {
  textWrap?: string;
  textFlow?: string;
  zOrder?: number;
  width?: number;
  height?: number;
  margins?: {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
  };
  clip?: {
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
  };
}

export type TextMutationOperation =
  | {
      op: 'replace_text_in_paragraph';
      sectionIndex: number;
      blockIndex: number;
      text: string;
    }
  | {
      op: 'replace_text_in_supplemental_paragraph';
      kind: SupplementalTextKind;
      paragraphIndex: number;
      text: string;
    }
  | {
      op: 'replace_image_asset';
      imagePath: string;
      targetId?: string;
      binaryItemId?: string;
    }
  | {
      op: 'replace_image_instance_asset';
      imagePath: string;
      targetId?: string;
      binaryItemId?: string;
      objectId?: string;
      instanceId?: string;
    }
  | {
      op: 'delete_image_instance';
      targetId?: string;
      binaryItemId?: string;
      objectId?: string;
      instanceId?: string;
    }
  | {
      op: 'insert_image_from_prototype';
      imagePath: string;
      destinationTargetId: string;
      prototypeTargetId?: string;
      prototypeBinaryItemId?: string;
      prototypeObjectId?: string;
      prototypeInstanceId?: string;
      patch?: ImagePlacementPatch;
    }
  | {
      op: 'update_image_placement';
      targetId?: string;
      binaryItemId?: string;
      objectId?: string;
      instanceId?: string;
      patch: ImagePlacementPatch;
    }
  | {
      op: 'insert_paragraph_after';
      sectionIndex: number;
      blockIndex: number;
      text: string;
    }
  | {
      op: 'delete_paragraph';
      sectionIndex: number;
      blockIndex: number;
    }
  | {
      op: 'replace_table_cell_text';
      sectionIndex: number;
      blockIndex: number;
      rowIndex: number;
      columnIndex: number;
      text: string;
    }
  | {
      op: 'insert_table_row' | 'delete_table_row';
      sectionIndex: number;
      blockIndex: number;
      rowIndex: number;
      boundaryPolicy?: 'reject' | 'split_boundary_merges';
    }
  | {
      op: 'clone_table_region';
      sectionIndex: number;
      blockIndex: number;
      templateStartRowIndex: number;
      templateEndRowIndex: number;
      insertAfterRowIndex: number;
      boundaryPolicy?: 'reject' | 'split_boundary_merges';
    };
