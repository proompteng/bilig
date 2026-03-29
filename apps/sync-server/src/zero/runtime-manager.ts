import { SpreadsheetEngine, type EngineReplicaSnapshot } from "@bilig/core";
import type { WorkbookSnapshot } from "@bilig/protocol";
import {
  loadWorkbookRuntimeMetadata,
  loadWorkbookState,
  type Queryable,
  type WorkbookRuntimeMetadata,
  type WorkbookRuntimeState,
} from "./store.js";

export interface WorkbookRuntime extends WorkbookRuntimeState {
  documentId: string;
  engine: SpreadsheetEngine;
}

export interface WorkbookRuntimeManagerOptions {
  maxEntries?: number;
  now?: () => number;
  loadMetadata?: (db: Queryable, documentId: string) => Promise<WorkbookRuntimeMetadata>;
  loadState?: (db: Queryable, documentId: string) => Promise<WorkbookRuntimeState>;
  createEngine?: (
    documentId: string,
    snapshot: WorkbookSnapshot,
    replicaSnapshot: EngineReplicaSnapshot | null,
  ) => Promise<SpreadsheetEngine>;
}

interface RuntimeSession extends WorkbookRuntime {
  lastAccessedAt: number;
}

interface MutationCommit {
  snapshot: WorkbookSnapshot;
  replicaSnapshot: EngineReplicaSnapshot | null;
  headRevision: number;
  calculatedRevision: number;
  ownerUserId: string;
}

interface RecalcCommit {
  calculatedRevision: number;
  snapshot?: WorkbookSnapshot;
  replicaSnapshot?: EngineReplicaSnapshot | null;
}

const DEFAULT_MAX_ENTRIES = 64;

async function createWorkbookEngine(
  documentId: string,
  snapshot: WorkbookSnapshot,
  replicaSnapshot: EngineReplicaSnapshot | null,
): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: documentId,
    replicaId: `server:${documentId}`,
  });
  await engine.ready();
  engine.importSnapshot(snapshot);
  if (replicaSnapshot) {
    engine.importReplicaSnapshot(replicaSnapshot);
  }
  return engine;
}

export class WorkbookRuntimeManager {
  private readonly sessions = new Map<string, RuntimeSession>();
  private readonly busyDocuments = new Set<string>();
  private readonly waiters = new Map<string, Array<() => void>>();
  private readonly maxEntries: number;
  private readonly now: () => number;
  private readonly loadMetadata;
  private readonly loadState;
  private readonly createEngine;

  constructor(options: WorkbookRuntimeManagerOptions = {}) {
    this.maxEntries = options.maxEntries ?? DEFAULT_MAX_ENTRIES;
    this.now = options.now ?? (() => Date.now());
    this.loadMetadata = options.loadMetadata ?? loadWorkbookRuntimeMetadata;
    this.loadState = options.loadState ?? loadWorkbookState;
    this.createEngine = options.createEngine ?? createWorkbookEngine;
  }

  async runExclusive<T>(documentId: string, task: () => Promise<T>): Promise<T> {
    const release = await this.acquire(documentId);
    try {
      return await task();
    } finally {
      release();
    }
  }

  async loadRuntime(
    db: Queryable,
    documentId: string,
    metadata?: WorkbookRuntimeMetadata,
  ): Promise<WorkbookRuntime> {
    const nextMetadata = metadata ?? (await this.loadMetadata(db, documentId));
    const cached = this.sessions.get(documentId);
    if (cached && cached.headRevision === nextMetadata.headRevision) {
      cached.calculatedRevision = nextMetadata.calculatedRevision;
      cached.ownerUserId = nextMetadata.ownerUserId;
      this.touchSession(cached);
      return cached;
    }

    const state = await this.loadState(db, documentId);
    const session: RuntimeSession = {
      documentId,
      engine: await this.createEngine(documentId, state.snapshot, state.replicaSnapshot),
      snapshot: state.snapshot,
      replicaSnapshot: state.replicaSnapshot,
      headRevision: state.headRevision,
      calculatedRevision: state.calculatedRevision,
      ownerUserId: state.ownerUserId,
      lastAccessedAt: this.now(),
    };
    this.sessions.set(documentId, session);
    this.touchSession(session);
    this.evictIfNeeded();
    return session;
  }

  commitMutation(documentId: string, commit: MutationCommit): void {
    const session = this.sessions.get(documentId);
    if (!session) {
      return;
    }
    session.snapshot = commit.snapshot;
    session.replicaSnapshot = commit.replicaSnapshot;
    session.headRevision = commit.headRevision;
    session.calculatedRevision = commit.calculatedRevision;
    session.ownerUserId = commit.ownerUserId;
    this.touchSession(session);
  }

  commitRecalc(documentId: string, commit: RecalcCommit): void {
    const session = this.sessions.get(documentId);
    if (!session) {
      return;
    }
    session.calculatedRevision = commit.calculatedRevision;
    if (commit.snapshot) {
      session.snapshot = commit.snapshot;
    }
    if (commit.replicaSnapshot !== undefined) {
      session.replicaSnapshot = commit.replicaSnapshot;
    }
    this.touchSession(session);
  }

  invalidate(documentId: string): void {
    this.sessions.delete(documentId);
  }

  async close(): Promise<void> {
    this.sessions.clear();
    this.waiters.clear();
    this.busyDocuments.clear();
  }

  private touchSession(session: RuntimeSession): void {
    session.lastAccessedAt = this.now();
    this.sessions.delete(session.documentId);
    this.sessions.set(session.documentId, session);
  }

  private evictIfNeeded(): void {
    while (this.sessions.size > this.maxEntries) {
      const oldest = this.sessions.keys().next().value;
      if (oldest === undefined) {
        return;
      }
      this.sessions.delete(oldest);
    }
  }

  private async acquire(documentId: string): Promise<() => void> {
    if (!this.busyDocuments.has(documentId)) {
      this.busyDocuments.add(documentId);
      return () => {
        this.release(documentId);
      };
    }

    return await new Promise<() => void>((resolve) => {
      const queue = this.waiters.get(documentId) ?? [];
      queue.push(() => {
        resolve(() => {
          this.release(documentId);
        });
      });
      this.waiters.set(documentId, queue);
    });
  }

  private release(documentId: string): void {
    const queue = this.waiters.get(documentId);
    const next = queue?.shift();
    if (queue && queue.length === 0) {
      this.waiters.delete(documentId);
    }
    if (next) {
      next();
      return;
    }
    this.busyDocuments.delete(documentId);
  }
}
