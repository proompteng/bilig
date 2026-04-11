import type { SpreadsheetEngine } from "@bilig/core";
import type { RecalcMetrics, SyncState, WorkbookDefinedNameSnapshot } from "@bilig/protocol";
import type { WorkerEngine } from "./worker-runtime-support.js";

interface WorkbookFailedPendingMutationLike {
  readonly id: string;
  readonly method: string;
  readonly failureMessage: string;
  readonly attemptCount: number;
}

interface WorkbookPendingMutationSummaryLike {
  readonly activeCount: number;
  readonly failedCount: number;
  readonly firstFailed: WorkbookFailedPendingMutationLike | null;
}

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

function clonePendingMutationSummary(
  summary: WorkbookPendingMutationSummaryLike | undefined,
): WorkbookPendingMutationSummaryLike | undefined {
  if (!summary) {
    return undefined;
  }
  return {
    activeCount: summary.activeCount,
    failedCount: summary.failedCount,
    firstFailed: summary.firstFailed
      ? {
          id: summary.firstFailed.id,
          method: summary.firstFailed.method,
          failureMessage: summary.firstFailed.failureMessage,
          attemptCount: summary.firstFailed.attemptCount,
        }
      : null,
  };
}

export function cloneWorkerRuntimeState(input: {
  workbookName: string;
  sheetNames: readonly string[];
  definedNames: readonly WorkbookDefinedNameSnapshot[];
  metrics: RecalcMetrics;
  syncState: SyncState;
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike;
  localPersistenceMode?: "persistent" | "ephemeral" | "follower";
}): {
  workbookName: string;
  sheetNames: string[];
  definedNames: WorkbookDefinedNameSnapshot[];
  metrics: RecalcMetrics;
  syncState: SyncState;
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike;
  localPersistenceMode?: "persistent" | "ephemeral" | "follower";
} {
  const pendingMutationSummary = clonePendingMutationSummary(input.pendingMutationSummary);
  return {
    workbookName: input.workbookName,
    sheetNames: [...input.sheetNames],
    definedNames: input.definedNames.map((entry) => structuredClone(entry)),
    metrics: cloneRuntimeMetrics(input.metrics),
    syncState: input.syncState,
    ...(pendingMutationSummary ? { pendingMutationSummary } : {}),
    ...(input.localPersistenceMode ? { localPersistenceMode: input.localPersistenceMode } : {}),
  };
}

export function withExternalSyncState(
  state: {
    workbookName: string;
    sheetNames: readonly string[];
    definedNames: readonly WorkbookDefinedNameSnapshot[];
    metrics: RecalcMetrics;
    syncState: SyncState;
    pendingMutationSummary?: WorkbookPendingMutationSummaryLike;
    localPersistenceMode?: "persistent" | "ephemeral" | "follower";
  },
  externalSyncState: SyncState | null,
): {
  workbookName: string;
  sheetNames: string[];
  definedNames: WorkbookDefinedNameSnapshot[];
  metrics: RecalcMetrics;
  syncState: SyncState;
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike;
  localPersistenceMode?: "persistent" | "ephemeral" | "follower";
} {
  const nextState = cloneWorkerRuntimeState(state);
  nextState.syncState = externalSyncState ?? state.syncState;
  return nextState;
}

export function buildWorkerRuntimeStateFromBootstrap(input: {
  workbookName: string;
  sheetNames: readonly string[];
  definedNames?: readonly WorkbookDefinedNameSnapshot[];
  localPersistenceMode?: "persistent" | "ephemeral" | "follower";
}): {
  workbookName: string;
  sheetNames: string[];
  definedNames: WorkbookDefinedNameSnapshot[];
  metrics: RecalcMetrics;
  syncState: SyncState;
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike;
  localPersistenceMode?: "persistent" | "ephemeral" | "follower";
} {
  return {
    workbookName: input.workbookName,
    sheetNames: [...input.sheetNames],
    definedNames: (input.definedNames ?? []).map((entry) => structuredClone(entry)),
    metrics: cloneRuntimeMetrics(),
    syncState: "syncing",
    ...(input.localPersistenceMode ? { localPersistenceMode: input.localPersistenceMode } : {}),
  };
}

export function buildWorkerRuntimeStateFromEngine(engine: SpreadsheetEngine & WorkerEngine): {
  workbookName: string;
  sheetNames: string[];
  definedNames: WorkbookDefinedNameSnapshot[];
  metrics: RecalcMetrics;
  syncState: SyncState;
  pendingMutationSummary?: WorkbookPendingMutationSummaryLike;
  localPersistenceMode?: "persistent" | "ephemeral" | "follower";
} {
  return {
    workbookName: engine.workbook.workbookName,
    sheetNames: listOrderedSheetNames(engine.workbook),
    definedNames: engine.getDefinedNames().map((entry) => structuredClone(entry)),
    metrics: cloneRuntimeMetrics(engine.getLastMetrics()),
    syncState: engine.getSyncState(),
  };
}
