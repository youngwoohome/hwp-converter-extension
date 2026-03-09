import type { DocumentSession } from './types.js';

export class DocumentSessionStore {
  private readonly store = new Map<string, DocumentSession>();

  put(session: DocumentSession): void {
    this.store.set(session.documentId, session);
  }

  get(documentId: string): DocumentSession | null {
    return this.store.get(documentId) ?? null;
  }

  update(documentId: string, updates: Partial<DocumentSession>): DocumentSession | null {
    const current = this.store.get(documentId);
    if (!current) return null;

    const next: DocumentSession = {
      ...current,
      ...updates,
      updatedAt: new Date().toISOString(),
    };
    this.store.set(documentId, next);
    return next;
  }
}
