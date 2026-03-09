import type { ApiWarning, DocumentTemplateId } from '../types.js';
import { saveStructuredHwpxWithPlainText } from './structuredHwpx.js';

const TEMPLATE_PARAGRAPHS: Record<DocumentTemplateId, string[]> = {
  blank: [''],
  report: [
    '[Report Title]',
    'Summary',
    '',
    'Background',
    '',
    'Details',
    '',
    'Action Items',
    '',
  ],
  minutes: [
    '[Meeting Title]',
    'Date',
    '',
    'Location',
    '',
    'Attendees',
    '',
    'Agenda',
    '',
    'Decisions and Follow-ups',
    '',
  ],
  proposal: [
    '[Proposal Title]',
    'Overview',
    '',
    'Scope',
    '',
    'Timeline',
    '',
    'Budget Notes',
    '',
    'Approval Notes',
    '',
  ],
};

export function isDocumentTemplateId(value: unknown): value is DocumentTemplateId {
  return typeof value === 'string' && value in TEMPLATE_PARAGRAPHS;
}

export function listDocumentTemplateIds(): DocumentTemplateId[] {
  return Object.keys(TEMPLATE_PARAGRAPHS) as DocumentTemplateId[];
}

export async function applyDocumentTemplate(filePath: string, templateId: DocumentTemplateId): Promise<ApiWarning[]> {
  if (templateId === 'blank') {
    return [];
  }

  const paragraphs = TEMPLATE_PARAGRAPHS[templateId];
  await saveStructuredHwpxWithPlainText(filePath, paragraphs.join('\n\n'));

  return [
    {
      code: 'TEMPLATE_APPLIED',
      message: `Applied canonical "${templateId}" HWPX skeleton template.`,
    },
  ];
}
