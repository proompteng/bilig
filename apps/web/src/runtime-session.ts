import type { CellSnapshot, RecalcMetrics, Viewport } from "@bilig/protocol";
import { ValueTag } from "@bilig/protocol";
import type { WorkbookWorkerStateSnapshot } from "./worker-runtime.js";
import { WorkerViewportCache } from "./viewport-cache.js";
import { ZeroWorkbookBridge, type ZeroWorkbookBridgeState } from "./zero/ZeroWorkbookBridge.js";

export interface WorkerHandle {
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
  readonly persistState: boolean;
  readonly zero: ZeroClient;
  readonly initialSelection: WorkerRuntimeSelection;
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
    sheetViewId?: string,
  ) => () => void;
  readonly dispose: () => void;
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

function noop(): void {}

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

function createRuntimeState(
  documentId: string,
  bridgeState: ZeroWorkbookBridgeState | null,
): WorkbookWorkerStateSnapshot {
  return {
    workbookName: bridgeState?.workbookName ?? documentId,
    sheetNames: bridgeState?.sheetNames ? [...bridgeState.sheetNames] : ["Sheet1"],
    metrics: EMPTY_METRICS,
    syncState: bridgeState ? "live" : "syncing",
  };
}

export async function createWorkerRuntimeSessionController(
  input: CreateWorkerRuntimeSessionInput,
  callbacks: WorkerRuntimeSessionCallbacks,
): Promise<WorkerRuntimeSessionController> {
  const cache = new WorkerViewportCache();
  cache.setKnownSheets([input.initialSelection.sheetName]);
  const handle: WorkerHandle = { cache };
  let currentSelection = input.initialSelection;
  let currentBridgeState: ZeroWorkbookBridgeState | null = null;
  let currentRuntimeState = createRuntimeState(input.documentId, null);
  let disposed = false;
  let unsubscribeBridgeWorkbook: () => void = noop;
  let unsubscribeBridgeSelection: () => void = noop;

  const reportError = (error: unknown) => {
    if (disposed) {
      return;
    }
    callbacks.onError(toErrorMessage(error));
  };

  const zeroBridge = new ZeroWorkbookBridge(input.zero, input.documentId, cache, reportError);

  const applySelection = async (selection: WorkerRuntimeSelection): Promise<CellSnapshot> => {
    currentSelection = selection;
    callbacks.onSelection(selection);
    zeroBridge.setSelection(selection.sheetName, selection.address);
    const next =
      cache.peekCell(selection.sheetName, selection.address) ?? emptyCellSnapshot(selection);
    callbacks.onSelectedCell(next);
    return next;
  };

  callbacks.onRuntimeState(currentRuntimeState);
  callbacks.onSelection(currentSelection);
  callbacks.onSelectedCell(emptyCellSnapshot(currentSelection));

  unsubscribeBridgeWorkbook = zeroBridge.subscribeWorkbookState((state) => {
    currentBridgeState = state;
    currentRuntimeState = createRuntimeState(input.documentId, state);
    callbacks.onBridgeState(state);
    callbacks.onRuntimeState(currentRuntimeState);
    cache.setKnownSheets(state.sheetNames);
    const reconciledSelection = reconcileSelection(currentSelection, state.sheetNames);
    if (
      reconciledSelection.sheetName !== currentSelection.sheetName ||
      reconciledSelection.address !== currentSelection.address
    ) {
      currentSelection = reconciledSelection;
      callbacks.onSelection(currentSelection);
      zeroBridge.setSelection(currentSelection.sheetName, currentSelection.address);
    }
  });

  unsubscribeBridgeSelection = zeroBridge.subscribeSelectedCell((cell) => {
    callbacks.onSelectedCell(cell ?? emptyCellSnapshot(currentSelection));
  });

  return {
    handle,
    runtimeState: currentRuntimeState,
    selectedCell:
      cache.peekCell(currentSelection.sheetName, currentSelection.address) ??
      emptyCellSnapshot(currentSelection),
    bridgeState: currentBridgeState,
    selection: currentSelection,
    async setSelection(selection) {
      await applySelection(selection);
    },
    subscribeViewport(sheetName, viewport, listener, sheetViewId) {
      return zeroBridge.subscribeViewport(sheetName, viewport, listener, sheetViewId);
    },
    dispose() {
      if (disposed) {
        return;
      }
      disposed = true;
      unsubscribeBridgeWorkbook();
      unsubscribeBridgeSelection();
      zeroBridge.dispose();
    },
  };
}
