import type { Zero } from "@rocicorp/zero";
import { createWorkerEngineClient, type MessagePortLike } from "@bilig/worker-transport";
import { parseCellAddress } from "@bilig/formula";
import type { CellSnapshot, RecalcMetrics, Viewport, WorkbookSnapshot } from "@bilig/protocol";
import { ValueTag } from "@bilig/protocol";
import type {
  WorkbookWorkerBootstrapResult,
  WorkbookWorkerStateSnapshot,
} from "./worker-runtime.js";
import { WorkerViewportCache } from "./viewport-cache.js";
import { ZeroWorkbookLiveSync } from "./runtime-zero-live.js";

export interface WorkerHandle {
  readonly cache: WorkerViewportCache;
}

export interface WorkerRuntimeSelection {
  readonly sheetName: string;
  readonly address: string;
}

export interface WorkerRuntimeSessionCallbacks {
  readonly onRuntimeState: (runtimeState: WorkbookWorkerStateSnapshot) => void;
  readonly onSelection: (selection: WorkerRuntimeSelection) => void;
  readonly onError: (message: string) => void;
}

export type ZeroClient = Zero;

export interface ZeroWorkbookSyncSource {
  materialize(query: unknown): unknown;
}

export interface CreateWorkerRuntimeSessionInput {
  readonly documentId: string;
  readonly replicaId: string;
  readonly persistState: boolean;
  readonly initialSelection: WorkerRuntimeSelection;
  readonly zero?: ZeroWorkbookSyncSource;
  readonly fetchImpl?: typeof fetch;
  readonly createWorker?: () => WorkerSessionPort;
}

export interface WorkerRuntimeSessionController {
  readonly handle: WorkerHandle;
  readonly runtimeState: WorkbookWorkerStateSnapshot;
  readonly selection: WorkerRuntimeSelection;
  readonly invoke: (method: string, ...args: unknown[]) => Promise<unknown>;
  readonly setSelection: (selection: WorkerRuntimeSelection) => Promise<void>;
  readonly subscribeViewport: (
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
    sheetViewId?: string,
  ) => () => void;
  readonly dispose: () => void;
}

interface WorkerSessionPort extends MessagePortLike {
  terminate?: () => void;
}

const EMPTY_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
};
const EMPTY_UNSUBSCRIBE = () => {};

function createInitialRuntimeState(documentId: string): WorkbookWorkerStateSnapshot {
  return {
    workbookName: documentId,
    sheetNames: ["Sheet1"],
    metrics: EMPTY_METRICS,
    syncState: "syncing",
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  return (
    isRecord(value) &&
    value["version"] === 1 &&
    isRecord(value["workbook"]) &&
    typeof value["workbook"]["name"] === "string" &&
    Array.isArray(value["sheets"])
  );
}

function isCellSnapshot(value: unknown): value is CellSnapshot {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["address"] === "string" &&
    typeof value["flags"] === "number" &&
    typeof value["version"] === "number" &&
    isRecord(value["value"]) &&
    typeof value["value"]["tag"] === "number"
  );
}

function isWorkbookWorkerStateSnapshot(value: unknown): value is WorkbookWorkerStateSnapshot {
  return (
    isRecord(value) &&
    typeof value["workbookName"] === "string" &&
    Array.isArray(value["sheetNames"]) &&
    value["sheetNames"].every((sheetName) => typeof sheetName === "string") &&
    isRecord(value["metrics"]) &&
    typeof value["syncState"] === "string"
  );
}

function isWorkbookWorkerBootstrapResult(value: unknown): value is WorkbookWorkerBootstrapResult {
  return (
    isRecord(value) &&
    typeof value["restoredFromPersistence"] === "boolean" &&
    isWorkbookWorkerStateSnapshot(value["runtimeState"])
  );
}

async function invokeWorkerMethod<T>(
  client: ReturnType<typeof createWorkerEngineClient>,
  method: string,
  guard: (value: unknown) => value is T,
  ...args: unknown[]
): Promise<T> {
  const value = await client.invoke(method, ...args);
  if (!guard(value)) {
    throw new Error(`Worker method ${method} returned an unexpected payload`);
  }
  return value;
}

function emptyCellSnapshot(selection: WorkerRuntimeSelection): CellSnapshot {
  return {
    sheetName: selection.sheetName,
    address: selection.address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  };
}

function selectionViewport(selection: WorkerRuntimeSelection): Viewport {
  const parsed = parseCellAddress(selection.address, selection.sheetName);
  return {
    rowStart: parsed.row,
    rowEnd: parsed.row,
    colStart: parsed.col,
    colEnd: parsed.col,
  };
}

function toErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function sameSelection(left: WorkerRuntimeSelection, right: WorkerRuntimeSelection): boolean {
  return left.sheetName === right.sheetName && left.address === right.address;
}

function reconcileSelection(
  selection: WorkerRuntimeSelection,
  sheetNames: readonly string[],
): WorkerRuntimeSelection {
  if (sheetNames.length === 0) {
    return selection;
  }
  if (sheetNames.includes(selection.sheetName)) {
    return selection;
  }
  return {
    sheetName: sheetNames[0]!,
    address: "A1",
  };
}

async function loadLatestWorkbookSnapshot(
  documentId: string,
  fetchImpl: typeof fetch,
): Promise<WorkbookSnapshot | null> {
  const response = await fetchImpl(
    `/v2/documents/${encodeURIComponent(documentId)}/snapshot/latest`,
    {
      headers: {
        accept: "application/json, application/vnd.bilig.workbook+json",
      },
      cache: "no-store",
    },
  );
  if (response.status === 404) {
    return null;
  }
  if (!response.ok) {
    throw new Error(`Failed to load workbook snapshot (${response.status})`);
  }
  const parsed: unknown = JSON.parse(await response.text());
  if (!isWorkbookSnapshot(parsed)) {
    throw new Error("Workbook snapshot payload does not match the expected schema");
  }
  return parsed;
}

function createWorkbookWorker(): WorkerSessionPort {
  return new Worker(new URL("./workbook.worker.ts", import.meta.url), {
    type: "module",
  }) as WorkerSessionPort;
}

export async function createWorkerRuntimeSessionController(
  input: CreateWorkerRuntimeSessionInput,
  callbacks: WorkerRuntimeSessionCallbacks,
): Promise<WorkerRuntimeSessionController> {
  const workerPort = (input.createWorker ?? createWorkbookWorker)();
  const client = createWorkerEngineClient({ port: workerPort });
  const cache = new WorkerViewportCache(client);
  const handle: WorkerHandle = { cache };
  const fetchImpl = input.fetchImpl ?? fetch;
  let currentSelection = input.initialSelection;
  let currentRuntimeState = createInitialRuntimeState(input.documentId);
  let disposed = false;
  let selectionViewportCleanup = EMPTY_UNSUBSCRIBE;
  const liveSync = input.zero
    ? new ZeroWorkbookLiveSync({
        zero: input.zero,
        documentId: input.documentId,
        cache,
        onError(message) {
          if (!disposed) {
            callbacks.onError(message);
          }
        },
      })
    : null;

  const publishRuntimeState = (runtimeState: WorkbookWorkerStateSnapshot) => {
    currentRuntimeState = runtimeState;
    cache.setKnownSheets(runtimeState.sheetNames);
    callbacks.onRuntimeState(runtimeState);
  };

  const subscribeProjectedViewport = (
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
  ): (() => void) => {
    const unsubscribeWorker = cache.subscribeViewport(sheetName, viewport, listener);
    const unsubscribeLive =
      liveSync?.subscribeViewport(sheetName, viewport, listener) ?? EMPTY_UNSUBSCRIBE;
    return () => {
      unsubscribeLive();
      unsubscribeWorker();
    };
  };

  const updateSelectionViewport = (selection: WorkerRuntimeSelection): void => {
    selectionViewportCleanup();
    selectionViewportCleanup = subscribeProjectedViewport(
      selection.sheetName,
      selectionViewport(selection),
      () => {},
    );
  };

  const applySelection = async (selection: WorkerRuntimeSelection): Promise<CellSnapshot> => {
    currentSelection = selection;
    callbacks.onSelection(selection);
    const snapshot = await invokeWorkerMethod(
      client,
      "getCell",
      isCellSnapshot,
      selection.sheetName,
      selection.address,
    );
    cache.setCellSnapshot(snapshot ?? emptyCellSnapshot(selection));
    updateSelectionViewport(selection);
    return snapshot ?? emptyCellSnapshot(selection);
  };

  const refreshRuntimeState = async (): Promise<void> => {
    const runtimeState = await invokeWorkerMethod(
      client,
      "getRuntimeState",
      isWorkbookWorkerStateSnapshot,
    );
    publishRuntimeState(runtimeState);
    const reconciledSelection = reconcileSelection(currentSelection, runtimeState.sheetNames);
    if (
      reconciledSelection.sheetName !== currentSelection.sheetName ||
      reconciledSelection.address !== currentSelection.address
    ) {
      await applySelection(reconciledSelection);
    }
  };

  callbacks.onRuntimeState(currentRuntimeState);
  callbacks.onSelection(currentSelection);

  try {
    const bootstrap = await invokeWorkerMethod(
      client,
      "bootstrap",
      isWorkbookWorkerBootstrapResult,
      {
        documentId: input.documentId,
        replicaId: input.replicaId,
        persistState: input.persistState,
      },
    );
    publishRuntimeState(bootstrap.runtimeState);

    if (!bootstrap.restoredFromPersistence) {
      const snapshot = await loadLatestWorkbookSnapshot(input.documentId, fetchImpl);
      if (snapshot) {
        const hydratedState = await invokeWorkerMethod(
          client,
          "replaceSnapshot",
          isWorkbookWorkerStateSnapshot,
          snapshot,
        );
        publishRuntimeState(hydratedState);
      }
    }

    await applySelection(reconcileSelection(currentSelection, currentRuntimeState.sheetNames));
  } catch (error) {
    liveSync?.dispose();
    client.dispose();
    workerPort.terminate?.();
    throw error;
  }

  return {
    get handle() {
      return handle;
    },
    get runtimeState() {
      return currentRuntimeState;
    },
    get selection() {
      return currentSelection;
    },
    async invoke(method, ...args) {
      try {
        const result = await client.invoke(method, ...args);
        if (method === "renderCommit" || method === "replaceSnapshot") {
          await refreshRuntimeState();
        }
        return result;
      } catch (error) {
        if (!disposed) {
          callbacks.onError(toErrorMessage(error));
        }
        throw error;
      }
    },
    async setSelection(selection) {
      try {
        if (sameSelection(selection, currentSelection)) {
          return;
        }
        await applySelection(selection);
      } catch (error) {
        if (!disposed) {
          callbacks.onError(toErrorMessage(error));
        }
        throw error;
      }
    },
    subscribeViewport(sheetName, viewport, listener) {
      return subscribeProjectedViewport(sheetName, viewport, listener);
    },
    dispose() {
      if (disposed) {
        return;
      }
      disposed = true;
      selectionViewportCleanup();
      selectionViewportCleanup = EMPTY_UNSUBSCRIBE;
      liveSync?.dispose();
      client.dispose();
      workerPort.terminate?.();
    },
  };
}
