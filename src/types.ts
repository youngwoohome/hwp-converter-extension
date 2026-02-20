export interface OfficeExtensionConvertRequest {
  filetype: string;
  outputtype: string;
  key?: string;
  title?: string;
  filePath?: string;
  url?: string;
  outputPath?: string;
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

export interface ConvertContext {
  sourcePath: string;
  sourceFormat: SourceFormat;
  targetFormat: TargetFormat;
  requestedOutputPath?: string;
}

export interface ConvertedArtifact {
  outputPath: string;
  targetFormat: TargetFormat;
}

export interface TextExtractionResult {
  paragraphs: string[];
  rawText: string;
}
