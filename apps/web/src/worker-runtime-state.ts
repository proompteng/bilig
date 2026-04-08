import type { SpreadsheetEngine } from "@bilig/core";
import type { RecalcMetrics, SyncState } from "@bilig/protocol";
import type { WorkerEngine } from "./worker-runtime-support.js";

export const EMPTY_RUNTIME_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
};

interface WorkbookSheetLike {
  name: string;
  order: number;
}

interface WorkbookLike {
  sheetsByName: Map<string, WorkbookSheetLike>;
}

export function listOrderedSheetNames(workbook: WorkbookLike): string[] {
  return [...workbook.sheetsByName.values()]
    .toSorted((left, right) => left.order - right.order)
    .map((sheet) => sheet.name);
}

export function cloneRuntimeMetrics(metrics: RecalcMetrics = EMPTY_RUNTIME_METRICS): RecalcMetrics {
  return { ...metrics };
}

export function cloneWorkerRuntimeState(input: {
  workbookName: string;
  sheetNames: readonly string[];
  metrics: RecalcMetrics;
  syncState: SyncState;
}): {
  workbookName: string;
  sheetNames: string[];
  metrics: RecalcMetrics;
  syncState: SyncState;
} {
  return {
    workbookName: input.workbookName,
    sheetNames: [...input.sheetNames],
    metrics: cloneRuntimeMetrics(input.metrics),
    syncState: input.syncState,
  };
}

export function withExternalSyncState(
  state: {
    workbookName: string;
    sheetNames: readonly string[];
    metrics: RecalcMetrics;
    syncState: SyncState;
  },
  externalSyncState: SyncState | null,
): {
  workbookName: string;
  sheetNames: string[];
  metrics: RecalcMetrics;
  syncState: SyncState;
} {
  const nextState = cloneWorkerRuntimeState(state);
  nextState.syncState = externalSyncState ?? state.syncState;
  return nextState;
}

export function buildWorkerRuntimeStateFromBootstrap(input: {
  workbookName: string;
  sheetNames: readonly string[];
}): {
  workbookName: string;
  sheetNames: string[];
  metrics: RecalcMetrics;
  syncState: SyncState;
} {
  return {
    workbookName: input.workbookName,
    sheetNames: [...input.sheetNames],
    metrics: cloneRuntimeMetrics(),
    syncState: "syncing",
  };
}

export function buildWorkerRuntimeStateFromEngine(engine: SpreadsheetEngine & WorkerEngine): {
  workbookName: string;
  sheetNames: string[];
  metrics: RecalcMetrics;
  syncState: SyncState;
} {
  return {
    workbookName: engine.workbook.workbookName,
    sheetNames: listOrderedSheetNames(engine.workbook),
    metrics: cloneRuntimeMetrics(engine.getLastMetrics()),
    syncState: engine.getSyncState(),
  };
}
