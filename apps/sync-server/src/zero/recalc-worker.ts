import { SpreadsheetEngine } from "@bilig/core";
import {
  leaseNextRecalcJob,
  loadWorkbookState,
  markRecalcJobCompleted,
  markRecalcJobFailed,
  markRecalcJobSuperseded,
  type Queryable,
} from "./store.js";
import { materializeCellEvalProjection } from "./projection.js";

const IDLE_POLL_MS = 250;
const BUSY_POLL_MS = 25;

async function createWorkbookEngine(
  documentId: string,
  snapshot: Awaited<ReturnType<typeof loadWorkbookState>>["snapshot"],
  replicaSnapshot: Awaited<ReturnType<typeof loadWorkbookState>>["replicaSnapshot"],
) {
  const engine = new SpreadsheetEngine({
    workbookName: documentId,
    replicaId: `recalc:${documentId}`,
  });
  await engine.ready();
  engine.importSnapshot(snapshot);
  if (replicaSnapshot) {
    engine.importReplicaSnapshot(replicaSnapshot);
  }
  return engine;
}

export class ZeroRecalcWorker {
  private timer: ReturnType<typeof setTimeout> | null = null;
  private running = false;
  private closed = false;

  constructor(
    private readonly db: Queryable,
    private readonly workerId = `bilig-recalc:${process.pid}:${Math.random().toString(36).slice(2)}`,
  ) {}

  start(): void {
    if (this.closed || this.timer) {
      return;
    }
    this.schedule(0);
  }

  stop(): void {
    this.closed = true;
    if (this.timer) {
      clearTimeout(this.timer);
      this.timer = null;
    }
  }

  private schedule(delayMs: number): void {
    if (this.closed) {
      return;
    }
    this.timer = setTimeout(() => {
      void this.tick();
    }, delayMs);
  }

  private async tick(): Promise<void> {
    this.timer = null;
    if (this.closed || this.running) {
      return;
    }
    this.running = true;
    let processed = false;
    try {
      processed = await this.processNextJob();
    } finally {
      this.running = false;
      this.schedule(processed ? BUSY_POLL_MS : IDLE_POLL_MS);
    }
  }

  private async processNextJob(): Promise<boolean> {
    const lease = await leaseNextRecalcJob(this.db, this.workerId);
    if (!lease) {
      return false;
    }

    try {
      const state = await loadWorkbookState(this.db, lease.workbookId);
      if (state.headRevision !== lease.toRevision) {
        await markRecalcJobSuperseded(this.db, lease);
        return true;
      }

      const engine = await createWorkbookEngine(
        lease.workbookId,
        state.snapshot,
        state.replicaSnapshot,
      );
      const cellEvalRows = materializeCellEvalProjection(
        engine,
        lease.workbookId,
        lease.toRevision,
        new Date().toISOString(),
      );
      await markRecalcJobCompleted(
        this.db,
        lease,
        cellEvalRows,
        engine.exportSnapshot(),
        engine.exportReplicaSnapshot(),
      );
    } catch (error) {
      await markRecalcJobFailed(this.db, lease, error);
    }

    return true;
  }
}
