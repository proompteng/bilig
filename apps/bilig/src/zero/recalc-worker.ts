import { WorkbookRuntimeManager } from "../workbook-runtime/runtime-manager.js";
import { materializeCellEvalProjection } from "./projection.js";
import {
  leaseNextRecalcJob,
  markRecalcJobCompleted,
  markRecalcJobFailed,
  markRecalcJobSuperseded,
} from "./recalc-job-store.js";
import { shouldPersistWorkbookCheckpointRevision, type Queryable } from "./store.js";
import { loadWorkbookRuntimeMetadata } from "./workbook-runtime-store.js";

const IDLE_POLL_MS = 250;
const BUSY_POLL_MS = 25;

export class ZeroRecalcWorker {
  private timer: ReturnType<typeof setTimeout> | null = null;
  private running = false;
  private closed = false;

  constructor(
    private readonly db: Queryable,
    private readonly runtimeManager: WorkbookRuntimeManager,
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
      await this.runtimeManager.runExclusive(lease.workbookId, async () => {
        const metadata = await loadWorkbookRuntimeMetadata(this.db, lease.workbookId);
        if (metadata.headRevision !== lease.toRevision) {
          await markRecalcJobSuperseded(this.db, lease);
          return;
        }

        const runtime = await this.runtimeManager.loadRuntime(this.db, lease.workbookId, metadata);
        const changedCellIndices = lease.dirtyRegions
          ? runtime.engine.recalculateDirty(lease.dirtyRegions)
          : runtime.engine.recalculateNow();

        const cellEvalRows = materializeCellEvalProjection(
          runtime.engine,
          lease.workbookId,
          lease.toRevision,
          new Date().toISOString(),
          changedCellIndices,
        );
        const shouldCheckpoint = shouldPersistWorkbookCheckpointRevision(lease.toRevision);
        const nextSnapshot = shouldCheckpoint ? runtime.engine.exportSnapshot() : null;
        const nextReplicaSnapshot = shouldCheckpoint
          ? runtime.engine.exportReplicaSnapshot()
          : null;
        const completed = await markRecalcJobCompleted(
          this.db,
          lease,
          cellEvalRows,
          nextSnapshot,
          nextReplicaSnapshot,
          true,
        );
        if (!completed) {
          this.runtimeManager.invalidate(lease.workbookId);
          return;
        }
        this.runtimeManager.commitRecalc(lease.workbookId, {
          calculatedRevision: lease.toRevision,
          ...(nextSnapshot ? { snapshot: nextSnapshot } : {}),
          ...(shouldCheckpoint ? { replicaSnapshot: nextReplicaSnapshot } : {}),
        });
      });
    } catch (error) {
      this.runtimeManager.invalidate(lease.workbookId);
      await markRecalcJobFailed(this.db, lease, error);
    }

    return true;
  }
}
