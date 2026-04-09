import { SpreadsheetEngine, type EngineReplicaSnapshot } from "@bilig/core";
import { parseCellAddress } from "@bilig/formula";
import type { WorkbookSnapshot } from "@bilig/protocol";
import {
  buildWorkbookSourceProjection,
  type AxisMetadataSourceRow,
  type CellSourceRow,
} from "../zero/projection.js";
import type {
  Queryable,
  WorkbookProjectionCommit,
  WorkbookProjectionState,
  WorkbookRuntimeMetadata,
  WorkbookRuntimeState,
} from "../zero/store.js";
import { loadWorkbookRuntimeMetadata, loadWorkbookState } from "../zero/workbook-runtime-store.js";

export interface WorkbookRuntime extends WorkbookProjectionState {
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
  projectionCommit: WorkbookProjectionCommit;
  headRevision: number;
  calculatedRevision: number;
  ownerUserId: string;
}

interface RecalcCommit {
  calculatedRevision: number;
}

const DEFAULT_MAX_ENTRIES = 64;
const LOADED_PROJECTION_UPDATED_AT = "1970-01-01T00:00:00.000Z";

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
      cached.projection.workbook = {
        ...cached.projection.workbook,
        headRevision: nextMetadata.headRevision,
        calculatedRevision: nextMetadata.calculatedRevision,
        ownerUserId: nextMetadata.ownerUserId,
      };
      this.touchSession(cached);
      return cached;
    }

    const state = await this.loadState(db, documentId);
    const session: RuntimeSession = {
      documentId,
      engine: await this.createEngine(documentId, state.snapshot, state.replicaSnapshot),
      projection: buildWorkbookSourceProjection(documentId, state.snapshot, {
        revision: state.headRevision,
        calculatedRevision: state.calculatedRevision,
        ownerUserId: state.ownerUserId,
        updatedBy: state.ownerUserId,
        updatedAt: LOADED_PROJECTION_UPDATED_AT,
      }),
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
    this.applyProjectionCommit(session, commit.projectionCommit);
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
    session.projection.workbook = {
      ...session.projection.workbook,
      calculatedRevision: commit.calculatedRevision,
    };
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

  private applyProjectionCommit(session: RuntimeSession, commit: WorkbookProjectionCommit): void {
    switch (commit.kind) {
      case "replace":
        session.projection = commit.projection;
        return;
      case "focused-cell":
        session.projection.workbook = commit.workbook;
        session.projection.calculationSettings = commit.calculationSettings;
        session.projection.cells = withUpsertedProjectionCell(
          session.projection.cells,
          commit.sheetName,
          commit.address,
          commit.cell,
        );
        return;
      case "cell-range":
        session.projection.workbook = commit.workbook;
        session.projection.calculationSettings = commit.calculationSettings;
        if (commit.styles) {
          session.projection.styles = [...commit.styles];
        }
        if (commit.numberFormats) {
          session.projection.numberFormats = [...commit.numberFormats];
        }
        session.projection.cells = withReplacedProjectionCellsInRange(
          session.projection.cells,
          commit.range,
          commit.cells,
        );
        return;
      case "column-metadata":
        session.projection.workbook = commit.workbook;
        session.projection.calculationSettings = commit.calculationSettings;
        session.projection.columnMetadata = withReplacedProjectionColumnMetadata(
          session.projection.columnMetadata,
          commit.sheetName,
          commit.columnMetadata,
        );
        return;
      case "row-metadata":
        session.projection.workbook = commit.workbook;
        session.projection.calculationSettings = commit.calculationSettings;
        session.projection.rowMetadata = withReplacedProjectionRowMetadata(
          session.projection.rowMetadata,
          commit.sheetName,
          commit.rowMetadata,
        );
        return;
      default: {
        const exhaustive: never = commit;
        throw new Error(`Unhandled projection commit: ${JSON.stringify(exhaustive)}`);
      }
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

function withUpsertedProjectionCell(
  current: readonly CellSourceRow[],
  sheetName: string,
  address: string,
  nextRow: CellSourceRow | null,
): CellSourceRow[] {
  const cells = current.slice();
  const index = cells.findIndex(
    (entry) => entry.sheetName === sheetName && entry.address === address,
  );
  if (!nextRow) {
    if (index >= 0) {
      cells.splice(index, 1);
    }
    return cells;
  }
  if (index >= 0) {
    cells[index] = nextRow;
    return cells;
  }
  cells.push(nextRow);
  return cells;
}

function withReplacedProjectionCellsInRange(
  current: readonly CellSourceRow[],
  range: { sheetName: string; startAddress: string; endAddress: string },
  nextRows: readonly CellSourceRow[],
): CellSourceRow[] {
  const cells = current.slice();
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  const rowStart = Math.min(start.row, end.row);
  const rowEnd = Math.max(start.row, end.row);
  const colStart = Math.min(start.col, end.col);
  const colEnd = Math.max(start.col, end.col);
  for (let index = cells.length - 1; index >= 0; index -= 1) {
    const entry = cells[index]!;
    if (
      entry.sheetName === range.sheetName &&
      entry.rowNum >= rowStart &&
      entry.rowNum <= rowEnd &&
      entry.colNum >= colStart &&
      entry.colNum <= colEnd
    ) {
      cells.splice(index, 1);
    }
  }
  cells.push(...nextRows);
  return cells;
}

function withReplacedProjectionColumnMetadata(
  current: readonly AxisMetadataSourceRow[],
  sheetName: string,
  nextRows: readonly AxisMetadataSourceRow[],
): AxisMetadataSourceRow[] {
  const rows = current.slice();
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    if (rows[index]!.sheetName === sheetName) {
      rows.splice(index, 1);
    }
  }
  rows.push(...nextRows);
  return rows;
}

function withReplacedProjectionRowMetadata(
  current: readonly AxisMetadataSourceRow[],
  sheetName: string,
  nextRows: readonly AxisMetadataSourceRow[],
): AxisMetadataSourceRow[] {
  const rows = current.slice();
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    if (rows[index]!.sheetName === sheetName) {
      rows.splice(index, 1);
    }
  }
  rows.push(...nextRows);
  return rows;
}
