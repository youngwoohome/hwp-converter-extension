import { copyFile } from 'node:fs/promises';
import { basename, extname, join } from 'node:path';
import { tmpdir } from 'node:os';
import { nanoid } from 'nanoid';
import type { CheckpointRecord } from './types.js';
import { ensureParentDirectory, fileExists } from './utils/files.js';

export class CheckpointStore {
  private readonly store = new Map<string, CheckpointRecord[]>();

  constructor(private readonly rootDir: string = join(tmpdir(), 'hwp-converter-checkpoints')) {}

  async create(documentId: string, filePath: string): Promise<CheckpointRecord | null> {
    const exists = await fileExists(filePath);
    if (!exists) return null;

    const checkpointId = `ckpt_${nanoid(10)}`;
    const ext = extname(filePath) || '.bin';
    const name = basename(filePath, ext);
    const checkpointPath = join(this.rootDir, documentId, `${name}.${checkpointId}${ext}`);

    await ensureParentDirectory(checkpointPath);
    await copyFile(filePath, checkpointPath);

    const record: CheckpointRecord = {
      checkpointId,
      documentId,
      path: checkpointPath,
      createdAt: new Date().toISOString(),
    };

    const current = this.store.get(documentId) ?? [];
    current.push(record);
    this.store.set(documentId, current);
    return record;
  }

  list(documentId: string): CheckpointRecord[] {
    return [...(this.store.get(documentId) ?? [])];
  }

  get(documentId: string, checkpointId: string): CheckpointRecord | null {
    return this.list(documentId).find((item) => item.checkpointId === checkpointId) ?? null;
  }

  async restore(documentId: string, checkpointId: string, destinationPath: string): Promise<CheckpointRecord | null> {
    const checkpoint = this.get(documentId, checkpointId);
    if (!checkpoint) return null;

    await ensureParentDirectory(destinationPath);
    await copyFile(checkpoint.path, destinationPath);
    return checkpoint;
  }
}
