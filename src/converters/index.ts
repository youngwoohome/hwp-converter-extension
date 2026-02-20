import type { ConvertContext, ConvertedArtifact } from '../types.js';
import { convertHwp } from './hwp.js';
import { convertHwpx } from './hwpx.js';

export async function convertWithSource(context: ConvertContext): Promise<ConvertedArtifact> {
  if (context.sourceFormat === 'hwpx') {
    return convertHwpx(context);
  }

  if (context.sourceFormat === 'hwp') {
    return convertHwp(context);
  }

  const exhaustiveCheck: never = context.sourceFormat;
  throw new Error(`Unsupported source format: ${exhaustiveCheck}`);
}
