import { createReadStream } from 'node:fs';
import { unlink } from 'node:fs/promises';
import { fileExists } from './files.js';

interface StoredFile {
  path: string;
  expiresAt: number;
}

export class EphemeralFileStore {
  private readonly store = new Map<string, StoredFile>();

  constructor(private readonly ttlMs: number) {}

  put(token: string, path: string): void {
    this.store.set(token, {
      path,
      expiresAt: Date.now() + this.ttlMs,
    });
  }

  async consume(token: string): Promise<{ path: string; stream: ReturnType<typeof createReadStream> } | null> {
    await this.cleanupExpired();

    const entry = this.store.get(token);
    if (!entry) return null;

    const exists = await fileExists(entry.path);
    if (!exists) {
      this.store.delete(token);
      return null;
    }

    this.store.delete(token);
    return {
      path: entry.path,
      stream: createReadStream(entry.path),
    };
  }

  async cleanupExpired(): Promise<void> {
    const now = Date.now();
    const expiredEntries = Array.from(this.store.entries()).filter(([, value]) => value.expiresAt <= now);

    for (const [token, entry] of expiredEntries) {
      this.store.delete(token);
      try {
        await unlink(entry.path);
      } catch {
        // Ignore temp-file cleanup failures
      }
    }
  }
}
