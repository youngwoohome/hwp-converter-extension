import type { EmbeddedImageReference, ImageTargetSummary } from '../types.js';

export function buildImageTargetSummaries(imageReferences: EmbeddedImageReference[]): ImageTargetSummary[] {
  const grouped = new Map<string, EmbeddedImageReference[]>();
  const binaryCounts = new Map<string, number>();
  const objectIdCounts = new Map<string, number>();
  const instanceIdCounts = new Map<string, number>();

  imageReferences.forEach((reference) => {
    const bucket = grouped.get(reference.targetId) ?? [];
    bucket.push(reference);
    grouped.set(reference.targetId, bucket);
    binaryCounts.set(reference.binaryItemId, (binaryCounts.get(reference.binaryItemId) ?? 0) + 1);
    if (reference.placement?.objectId) {
      objectIdCounts.set(reference.placement.objectId, (objectIdCounts.get(reference.placement.objectId) ?? 0) + 1);
    }
    if (reference.placement?.instanceId) {
      instanceIdCounts.set(reference.placement.instanceId, (instanceIdCounts.get(reference.placement.instanceId) ?? 0) + 1);
    }
  });

  return [...grouped.entries()]
    .map(([targetId, references]) => {
      const first = references[0];
      const uniqueBinaryIds = [...new Set(references.map((reference) => reference.binaryItemId))];
      const uniqueObjectIds = [...new Set(references
        .map((reference) => reference.placement?.objectId)
        .filter((value): value is string => typeof value === 'string' && value.length > 0))];
      const uniqueInstanceIds = [...new Set(references
        .map((reference) => reference.placement?.instanceId)
        .filter((value): value is string => typeof value === 'string' && value.length > 0))];
      const singleBinaryId = uniqueBinaryIds.length === 1 ? uniqueBinaryIds[0] : null;
      const recommendedAssetLocator = references.length === 1
        && singleBinaryId
        && binaryCounts.get(singleBinaryId) === 1
        ? ('targetId' as const)
        : ('binaryItemId' as const);

      let recommendedPlacementLocator: ImageTargetSummary['recommendedPlacementLocator'] = 'none';
      if (references.length === 1) {
        recommendedPlacementLocator = 'targetId';
      } else if (uniqueObjectIds.length === references.length && uniqueObjectIds.every((id) => objectIdCounts.get(id) === 1)) {
        recommendedPlacementLocator = 'objectId';
      } else if (uniqueInstanceIds.length === references.length && uniqueInstanceIds.every((id) => instanceIdCounts.get(id) === 1)) {
        recommendedPlacementLocator = 'instanceId';
      } else if (uniqueBinaryIds.length === references.length && uniqueBinaryIds.every((id) => binaryCounts.get(id) === 1)) {
        recommendedPlacementLocator = 'binaryItemId';
      }

      return {
        targetId,
        containerId: first.containerId,
        kind: first.kind,
        sectionIndex: first.sectionIndex,
        blockIndex: first.blockIndex,
        rowIndex: first.rowIndex,
        columnIndex: first.columnIndex,
        paragraphIndex: first.paragraphIndex,
        imageCount: references.length,
        deterministicTargetReplacement: references.length === 1,
        recommendedLocator: recommendedAssetLocator,
        recommendedAssetLocator,
        recommendedPlacementLocator,
        binaryItemIds: uniqueBinaryIds,
        objectIds: uniqueObjectIds,
        instanceIds: uniqueInstanceIds,
        assetFileNames: [...new Set(references
          .map((reference) => reference.assetFileName)
          .filter((value): value is string => typeof value === 'string' && value.length > 0))],
      };
    })
    .sort((left, right) => left.targetId.localeCompare(right.targetId));
}
