import type { Queryable } from "../zero/store.js";
import { WorkbookRuntimeManager } from "../workbook-runtime/runtime-manager.js";
import { ZeroRecalcWorker } from "../zero/recalc-worker.js";

export class RecalcWorker {
  private readonly worker: ZeroRecalcWorker;

  constructor(db: Queryable, runtimeManager: WorkbookRuntimeManager, workerId: string) {
    this.worker = new ZeroRecalcWorker(db, runtimeManager, workerId);
  }

  async start(): Promise<void> {
    this.worker.start();
  }

  stop(): void {
    this.worker.stop();
  }
}
