import type { EmbeddedImageReference, ImageResourceSummary } from '../types.js';

export function buildImageResourceSummaries(imageReferences: EmbeddedImageReference[]): ImageResourceSummary[] {
  const grouped = new Map<string, EmbeddedImageReference[]>();

  imageReferences.forEach((reference) => {
    const bucket = grouped.get(reference.binaryItemId) ?? [];
    bucket.push(reference);
    grouped.set(reference.binaryItemId, bucket);
  });

  return [...grouped.entries()]
    .map(([binaryItemId, references]) => {
      const first = references[0];
      const targetIds = [...new Set(references.map((reference) => reference.targetId))].sort();
      return {
        binaryItemId,
        assetPath: first.assetPath,
        assetFileName: first.assetFileName,
        mediaType: first.mediaType,
        referenceCount: references.length,
        targetIds,
        sharedAcrossTargets: targetIds.length > 1,
        assetReplacementScope: targetIds.length > 1 ? 'multi_target' as const : 'single_target' as const,
        deterministicPlacementUpdate: references.length === 1,
      };
    })
    .sort((left, right) => left.binaryItemId.localeCompare(right.binaryItemId));
}
