import type { EngineReplicaSnapshot } from "@bilig/core";
import { SpreadsheetEngine } from "@bilig/core";
import type { EngineEvent, WorkbookSnapshot } from "@bilig/protocol";
import type { PendingWorkbookMutation } from "./workbook-sync.js";
import { applyPendingWorkbookMutationToEngine } from "./worker-runtime-mutation-replay.js";
import type { WorkerEngine } from "./worker-runtime-support.js";
import type { ProjectionOverlayScope } from "./worker-local-overlay.js";
import { collectProjectionOverlayScopeFromEngineEvents } from "./worker-local-overlay.js";

export async function createWorkbookEngineFromState(input: {
  workbookName: string;
  replicaId: string;
  snapshot: WorkbookSnapshot | null;
  replica: EngineReplicaSnapshot | null;
}): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: input.workbookName,
    replicaId: input.replicaId,
  });
  await engine.ready();
  if (input.snapshot) {
    engine.importSnapshot(input.snapshot);
  }
  if (input.replica) {
    engine.importReplicaSnapshot(input.replica);
  }
  if (engine.workbook.sheetsByName.size === 0) {
    engine.createSheet("Sheet1");
  }
  return engine;
}

export async function createProjectionEngineFromState(input: {
  workbookName: string;
  replicaId: string;
  snapshot: WorkbookSnapshot | null;
  replica: EngineReplicaSnapshot | null;
  pendingMutations: readonly PendingWorkbookMutation[];
}): Promise<{
  engine: SpreadsheetEngine;
  overlayScope: ProjectionOverlayScope | null;
}> {
  const engine = await createWorkbookEngineFromState(input);
  if (input.pendingMutations.length === 0) {
    return { engine, overlayScope: null };
  }
  const replayEvents: EngineEvent[] = [];
  const unsubscribe = engine.subscribe((event) => {
    replayEvents.push(event);
  });
  try {
    input.pendingMutations.forEach((mutation) => {
      applyPendingWorkbookMutationToEngine(engine, mutation);
    });
  } finally {
    unsubscribe();
  }
  return {
    engine,
    overlayScope: collectProjectionOverlayScopeFromEngineEvents(
      engine as SpreadsheetEngine & WorkerEngine,
      replayEvents,
    ),
  };
}
