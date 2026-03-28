import type { EngineOpBatch } from "@bilig/workbook-domain";

export interface StoredBatch {
  documentId: string;
  cursor: number;
  batch: EngineOpBatch;
  receivedAtUnixMs: number;
}

export interface StoredSnapshot {
  documentId: string;
  snapshotId: string;
  cursor: number;
  bytes: Uint8Array;
  contentType: string;
  createdAtUnixMs: number;
}

export interface BatchStore {
  append(documentId: string, batch: EngineOpBatch, receivedAtUnixMs?: number): Promise<StoredBatch>;
  listAfter(documentId: string, cursor: number, limit?: number): Promise<StoredBatch[]>;
  latestCursor(documentId: string): Promise<number>;
  reset(documentId: string, cursor?: number): Promise<void>;
}

export interface SnapshotStore {
  put(snapshot: StoredSnapshot): Promise<void>;
  latest(documentId: string): Promise<StoredSnapshot | null>;
}

export interface DocumentOwnershipStore {
  claim(documentId: string, ownerId: string, leaseExpiresAtUnixMs: number): Promise<boolean>;
  release(documentId: string, ownerId: string): Promise<void>;
  owner(documentId: string): Promise<string | null>;
}

export interface PresenceStore {
  join(documentId: string, sessionId: string): Promise<void>;
  leave(documentId: string, sessionId: string): Promise<void>;
  sessions(documentId: string): Promise<string[]>;
}

export class InMemoryBatchStore implements BatchStore {
  private readonly documents = new Map<string, StoredBatch[]>();
  private readonly baseCursors = new Map<string, number>();

  async append(
    documentId: string,
    batch: EngineOpBatch,
    receivedAtUnixMs = Date.now(),
  ): Promise<StoredBatch> {
    const entries = this.documents.get(documentId) ?? [];
    const record = {
      documentId,
      cursor:
        entries.length === 0
          ? (this.baseCursors.get(documentId) ?? 0) + 1
          : entries[entries.length - 1]!.cursor + 1,
      batch,
      receivedAtUnixMs,
    } satisfies StoredBatch;
    entries.push(record);
    this.documents.set(documentId, entries);
    return record;
  }

  async listAfter(documentId: string, cursor: number, limit = 256): Promise<StoredBatch[]> {
    return (this.documents.get(documentId) ?? [])
      .filter((entry) => entry.cursor > cursor)
      .slice(0, limit);
  }

  async latestCursor(documentId: string): Promise<number> {
    const entries = this.documents.get(documentId);
    return entries && entries.length > 0
      ? entries[entries.length - 1]!.cursor
      : (this.baseCursors.get(documentId) ?? 0);
  }

  async reset(documentId: string, cursor = 0): Promise<void> {
    this.documents.delete(documentId);
    this.baseCursors.set(documentId, cursor);
  }
}

export class InMemorySnapshotStore implements SnapshotStore {
  private readonly snapshots = new Map<string, StoredSnapshot>();

  async put(snapshot: StoredSnapshot): Promise<void> {
    this.snapshots.set(snapshot.documentId, snapshot);
  }

  async latest(documentId: string): Promise<StoredSnapshot | null> {
    return this.snapshots.get(documentId) ?? null;
  }
}

export class InMemoryDocumentOwnershipStore implements DocumentOwnershipStore {
  private readonly owners = new Map<string, { ownerId: string; leaseExpiresAtUnixMs: number }>();

  async claim(documentId: string, ownerId: string, leaseExpiresAtUnixMs: number): Promise<boolean> {
    const existing = this.owners.get(documentId);
    if (existing && existing.ownerId !== ownerId && existing.leaseExpiresAtUnixMs > Date.now()) {
      return false;
    }
    this.owners.set(documentId, { ownerId, leaseExpiresAtUnixMs });
    return true;
  }

  async release(documentId: string, ownerId: string): Promise<void> {
    const existing = this.owners.get(documentId);
    if (existing?.ownerId === ownerId) {
      this.owners.delete(documentId);
    }
  }

  async owner(documentId: string): Promise<string | null> {
    const existing = this.owners.get(documentId);
    if (!existing) {
      return null;
    }
    if (existing.leaseExpiresAtUnixMs <= Date.now()) {
      this.owners.delete(documentId);
      return null;
    }
    return existing.ownerId;
  }
}

export class InMemoryPresenceStore implements PresenceStore {
  private readonly documents = new Map<string, Set<string>>();

  async join(documentId: string, sessionId: string): Promise<void> {
    const sessions = this.documents.get(documentId) ?? new Set<string>();
    sessions.add(sessionId);
    this.documents.set(documentId, sessions);
  }

  async leave(documentId: string, sessionId: string): Promise<void> {
    const sessions = this.documents.get(documentId);
    if (!sessions) {
      return;
    }
    sessions.delete(sessionId);
    if (sessions.size === 0) {
      this.documents.delete(documentId);
    }
  }

  async sessions(documentId: string): Promise<string[]> {
    return [...(this.documents.get(documentId) ?? new Set<string>())].toSorted();
  }
}

export interface InMemoryDocumentPersistence {
  batches: InMemoryBatchStore;
  snapshots: InMemorySnapshotStore;
  ownership: InMemoryDocumentOwnershipStore;
  presence: InMemoryPresenceStore;
}

export function createInMemoryDocumentPersistence(): InMemoryDocumentPersistence {
  return {
    batches: new InMemoryBatchStore(),
    snapshots: new InMemorySnapshotStore(),
    ownership: new InMemoryDocumentOwnershipStore(),
    presence: new InMemoryPresenceStore(),
  };
}
