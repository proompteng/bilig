import type { CellSnapshot, Viewport } from "@bilig/protocol";
import { ValueTag } from "@bilig/protocol";
import {
  createWorkerEngineClient,
  type MessagePortLike,
  type WorkerEngineClient,
} from "@bilig/worker-transport";
import type {
  WorkbookWorkerBootstrapOptions,
  WorkbookWorkerStateSnapshot,
} from "./worker-runtime.js";
import { WorkerViewportCache } from "./viewport-cache.js";
import { ZeroWorkbookBridge, type ZeroWorkbookBridgeState } from "./zero/ZeroWorkbookBridge.js";

export interface WorkerLike {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  addEventListener(type: "message", listener: EventListener): void;
  removeEventListener(type: "message", listener: EventListener): void;
  terminate(): void;
}

export interface WorkerHandle {
  readonly worker: WorkerLike;
  readonly client: WorkerEngineClient;
  readonly cache: WorkerViewportCache;
}

export interface WorkerRuntimeSelection {
  readonly sheetName: string;
  readonly address: string;
}

export interface WorkerRuntimeSessionCallbacks {
  readonly onRuntimeState: (runtimeState: WorkbookWorkerStateSnapshot) => void;
  readonly onSelectedCell: (cell: CellSnapshot) => void;
  readonly onBridgeState: (bridgeState: ZeroWorkbookBridgeState | null) => void;
  readonly onSelection: (selection: WorkerRuntimeSelection) => void;
  readonly onError: (message: string) => void;
}

export type ZeroClient = ConstructorParameters<typeof ZeroWorkbookBridge>[0];

export interface CreateWorkerRuntimeSessionInput {
  readonly documentId: string;
  readonly replicaId: string;
  readonly baseUrl: string | null;
  readonly persistState: boolean;
  readonly zeroViewportBridge: boolean;
  readonly zero?: ZeroClient | null;
  readonly initialSelection: WorkerRuntimeSelection;
  readonly pollIntervalMs?: number;
}

export interface WorkerRuntimeSessionController {
  readonly handle: WorkerHandle;
  readonly runtimeState: WorkbookWorkerStateSnapshot;
  readonly selectedCell: CellSnapshot;
  readonly bridgeState: ZeroWorkbookBridgeState | null;
  readonly selection: WorkerRuntimeSelection;
  readonly setSelection: (selection: WorkerRuntimeSelection) => Promise<void>;
  readonly subscribeViewport: (
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
  ) => () => void;
  readonly dispose: () => void;
}

function noop(): void {}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
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

function isRuntimeStateSnapshot(value: unknown): value is WorkbookWorkerStateSnapshot {
  return (
    isRecord(value) &&
    typeof value["workbookName"] === "string" &&
    Array.isArray(value["sheetNames"]) &&
    isRecord(value["metrics"]) &&
    typeof value["syncState"] === "string"
  );
}

function createWorkerPort(worker: WorkerLike): MessagePortLike {
  type PortListener = Parameters<NonNullable<MessagePortLike["addEventListener"]>>[1];
  const listenerMap = new Map<PortListener, EventListener>();
  return {
    postMessage(message: unknown) {
      worker.postMessage(message, []);
    },
    addEventListener(type: "message", listener: PortListener) {
      const wrapped: EventListener = (event) => {
        if (event instanceof MessageEvent) {
          listener(event);
        }
      };
      listenerMap.set(listener, wrapped);
      worker.addEventListener(type, wrapped);
    },
    removeEventListener(type: "message", listener: PortListener) {
      const wrapped = listenerMap.get(listener);
      if (!wrapped) {
        return;
      }
      listenerMap.delete(listener);
      worker.removeEventListener(type, wrapped);
    },
  };
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

function toErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
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

async function loadSelectedCell(
  client: WorkerEngineClient,
  cache: WorkerViewportCache,
  selection: WorkerRuntimeSelection,
): Promise<CellSnapshot> {
  const cached = cache.peekCell(selection.sheetName, selection.address);
  if (cached) {
    return cached;
  }
  const response = await client.invoke("getCell", selection.sheetName, selection.address);
  if (!isCellSnapshot(response)) {
    throw new Error("Worker returned an invalid cell snapshot");
  }
  return response;
}

export async function createWorkerRuntimeSessionController(
  input: CreateWorkerRuntimeSessionInput,
  callbacks: WorkerRuntimeSessionCallbacks,
): Promise<WorkerRuntimeSessionController> {
  const worker = new Worker(new URL("./workbook.worker.ts", import.meta.url), { type: "module" });
  const client = createWorkerEngineClient({ port: createWorkerPort(worker) });
  const cache = new WorkerViewportCache(client);
  const handle: WorkerHandle = { worker, client, cache };
  let currentSelection = input.initialSelection;
  let currentBridgeState: ZeroWorkbookBridgeState | null = null;
  let bridge: ZeroWorkbookBridge | null = null;
  let disposed = false;
  let unsubscribeEvents: () => void = noop;
  let unsubscribeCache: () => void = noop;
  let unsubscribeBridgeWorkbook: () => void = noop;
  let unsubscribeBridgeSelection: () => void = noop;
  let interval = 0;

  const reportError = (error: unknown) => {
    if (disposed) {
      return;
    }
    callbacks.onError(toErrorMessage(error));
  };

  const applySelection = async (selection: WorkerRuntimeSelection): Promise<CellSnapshot> => {
    currentSelection = selection;
    callbacks.onSelection(selection);
    if (bridge) {
      bridge.setSelection(selection.sheetName, selection.address);
      const next =
        cache.peekCell(selection.sheetName, selection.address) ?? emptyCellSnapshot(selection);
      callbacks.onSelectedCell(next);
      return next;
    }
    const next = await loadSelectedCell(client, cache, selection);
    callbacks.onSelectedCell(next);
    return next;
  };

  const refreshRuntimeState = async (): Promise<WorkbookWorkerStateSnapshot> => {
    const response = await client.invoke("getRuntimeState");
    if (!isRuntimeStateSnapshot(response)) {
      throw new Error("Worker returned an invalid runtime state payload");
    }
    cache.setKnownSheets(response.sheetNames);
    callbacks.onRuntimeState(response);
    if (!bridge) {
      const reconciledSelection = reconcileSelection(currentSelection, response.sheetNames);
      if (
        reconciledSelection.sheetName !== currentSelection.sheetName ||
        reconciledSelection.address !== currentSelection.address
      ) {
        await applySelection(reconciledSelection);
      }
    }
    return response;
  };

  try {
    const response = await client.invoke("bootstrap", {
      documentId: input.documentId,
      replicaId: input.replicaId,
      baseUrl: input.baseUrl,
      persistState: input.persistState,
    } satisfies WorkbookWorkerBootstrapOptions);
    if (!isRuntimeStateSnapshot(response)) {
      throw new Error("Worker returned an invalid bootstrap payload");
    }
    cache.setKnownSheets(response.sheetNames);
    currentSelection = reconcileSelection(currentSelection, response.sheetNames);
    callbacks.onRuntimeState(response);
    callbacks.onSelection(currentSelection);
    const selectedCell = await loadSelectedCell(client, cache, currentSelection);
    callbacks.onSelectedCell(selectedCell);

    unsubscribeEvents = client.subscribe(() => {
      void refreshRuntimeState().catch(reportError);
    });
    unsubscribeCache = cache.subscribe(() => {
      if (disposed) {
        return;
      }
      const next = cache.peekCell(currentSelection.sheetName, currentSelection.address);
      if (next) {
        callbacks.onSelectedCell(next);
      }
    });
    interval = window.setInterval(() => {
      void refreshRuntimeState().catch(reportError);
    }, input.pollIntervalMs ?? 250);

    if (!input.baseUrl && input.zeroViewportBridge && input.zero) {
      bridge = new ZeroWorkbookBridge(input.zero, input.documentId, cache, reportError);
      unsubscribeBridgeWorkbook = bridge.subscribeWorkbookState((state) => {
        currentBridgeState = state;
        callbacks.onBridgeState(state);
        cache.setKnownSheets(state.sheetNames);
        const reconciledSelection = reconcileSelection(currentSelection, state.sheetNames);
        if (
          reconciledSelection.sheetName !== currentSelection.sheetName ||
          reconciledSelection.address !== currentSelection.address
        ) {
          currentSelection = reconciledSelection;
          callbacks.onSelection(currentSelection);
          bridge?.setSelection(currentSelection.sheetName, currentSelection.address);
        }
      });
      unsubscribeBridgeSelection = bridge.subscribeSelectedCell((cell) => {
        if (cell) {
          callbacks.onSelectedCell(cell);
        }
      });
    }

    return {
      handle,
      runtimeState: response,
      selectedCell,
      bridgeState: currentBridgeState,
      selection: currentSelection,
      async setSelection(selection) {
        await applySelection(selection);
      },
      subscribeViewport(sheetName, viewport, listener) {
        const disposers = [cache.subscribeViewport(sheetName, viewport, listener)];
        if (bridge) {
          disposers.push(bridge.subscribeViewport(sheetName, viewport, listener));
        }
        return () => {
          disposers.forEach((dispose) => dispose());
        };
      },
      dispose() {
        if (disposed) {
          return;
        }
        disposed = true;
        unsubscribeBridgeWorkbook();
        unsubscribeBridgeSelection();
        bridge?.dispose();
        bridge = null;
        unsubscribeEvents();
        unsubscribeCache();
        if (interval) {
          window.clearInterval(interval);
        }
        client.dispose();
        worker.terminate();
      },
    };
  } catch (error) {
    client.dispose();
    worker.terminate();
    throw error;
  }
}
