import type { Zero } from "@rocicorp/zero";
import { createWorkerEngineClient, type MessagePortLike } from "@bilig/worker-transport";
import { parseCellAddress } from "@bilig/formula";
import {
  isAuthoritativeWorkbookEventBatch,
  type AuthoritativeWorkbookEventBatch,
} from "@bilig/zero-sync";
import {
  isCellSnapshot,
  isWorkbookSnapshot,
  ValueTag,
  type CellSnapshot,
  type RecalcMetrics,
  type Viewport,
  type WorkbookSnapshot,
} from "@bilig/protocol";
import type {
  InstallAuthoritativeSnapshotInput,
  WorkbookWorkerBootstrapResult,
  WorkbookWorkerStateSnapshot,
} from "./worker-runtime.js";
import { ProjectedViewportStore } from "./projected-viewport-store.js";
import {
  ZeroWorkbookRevisionSync,
  type WorkbookRevisionState,
} from "./runtime-zero-revision-sync.js";

export interface WorkerHandle {
  readonly viewportStore: ProjectedViewportStore;
}

export interface WorkerRuntimeSelection {
  readonly sheetName: string;
  readonly address: string;
}

export type WorkerRuntimeSessionPhase =
  | "hydratingLocal"
  | "syncing"
  | "reconciling"
  | "recovering"
  | "steady";

export interface WorkerRuntimeSessionCallbacks {
  readonly onRuntimeState: (runtimeState: WorkbookWorkerStateSnapshot) => void;
  readonly onSelection: (selection: WorkerRuntimeSelection) => void;
  readonly onError: (message: string) => void;
  readonly onPhase?: (phase: WorkerRuntimeSessionPhase) => void;
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
    definedNames: [],
    metrics: EMPTY_METRICS,
    syncState: "syncing",
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookWorkerStateSnapshot(value: unknown): value is WorkbookWorkerStateSnapshot {
  if (!isRecord(value)) {
    return false;
  }
  return (
    typeof value["workbookName"] === "string" &&
    Array.isArray(value["sheetNames"]) &&
    value["sheetNames"].every((sheetName) => typeof sheetName === "string") &&
    Array.isArray(value["definedNames"]) &&
    typeof value["metrics"] === "object" &&
    value["metrics"] !== null &&
    typeof value["syncState"] === "string"
  );
}

function isWorkbookWorkerBootstrapResult(value: unknown): value is WorkbookWorkerBootstrapResult {
  if (!isRecord(value)) {
    return false;
  }
  return (
    typeof value["restoredFromPersistence"] === "boolean" &&
    typeof value["requiresAuthoritativeHydrate"] === "boolean" &&
    isWorkbookWorkerStateSnapshot(value["runtimeState"])
  );
}

function isNumber(value: unknown): value is number {
  return typeof value === "number";
}

function isInstallAuthoritativeSnapshotInput(
  value: unknown,
): value is InstallAuthoritativeSnapshotInput {
  if (!isRecord(value)) {
    return false;
  }
  return (
    isWorkbookSnapshot(value["snapshot"]) &&
    typeof value["authoritativeRevision"] === "number" &&
    (value["mode"] === "bootstrap" || value["mode"] === "reconcile")
  );
}

async function loadAuthoritativeEventBatch(
  documentId: string,
  afterRevision: number,
  fetchImpl: typeof fetch,
): Promise<AuthoritativeWorkbookEventBatch> {
  const response = await fetchImpl(
    `/v2/documents/${encodeURIComponent(documentId)}/events?afterRevision=${String(afterRevision)}`,
    {
      headers: {
        accept: "application/json",
      },
      cache: "no-store",
    },
  );
  if (!response.ok) {
    throw new Error(`Failed to load authoritative events (${response.status})`);
  }
  const parsed: unknown = JSON.parse(await response.text());
  if (!isAuthoritativeWorkbookEventBatch(parsed)) {
    throw new Error("Authoritative event payload does not match the expected schema");
  }
  return parsed;
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
  if (response.status === 204 || response.status === 404) {
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
  const viewportStore = new ProjectedViewportStore(client);
  const handle: WorkerHandle = { viewportStore };
  const fetchImpl = input.fetchImpl ?? fetch;
  let currentSelection = input.initialSelection;
  let currentRuntimeState = createInitialRuntimeState(input.documentId);
  let disposed = false;
  let bootstrapped = false;
  let currentAuthoritativeRevision = 0;
  let requestedAuthoritativeRevision = 0;
  let pendingRevisionState: WorkbookRevisionState | null = null;
  let rebaseQueue = Promise.resolve();
  let selectionViewportCleanup = EMPTY_UNSUBSCRIBE;
  let currentPhase: WorkerRuntimeSessionPhase = "hydratingLocal";
  const liveSync = input.zero
    ? new ZeroWorkbookRevisionSync({
        zero: input.zero,
        documentId: input.documentId,
        onRevisionState(revisionState) {
          pendingRevisionState = revisionState;
          if (bootstrapped) {
            queueAuthoritativeRebase(revisionState);
          }
        },
      })
    : null;

  const publishPhase = (phase: WorkerRuntimeSessionPhase) => {
    if (currentPhase === phase) {
      return;
    }
    currentPhase = phase;
    callbacks.onPhase?.(phase);
  };

  const publishRuntimeState = (runtimeState: WorkbookWorkerStateSnapshot) => {
    currentRuntimeState = runtimeState;
    viewportStore.setKnownSheets(runtimeState.sheetNames);
    callbacks.onRuntimeState(runtimeState);
  };

  const subscribeProjectedViewport = (
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
  ): (() => void) => {
    return viewportStore.subscribeViewport(sheetName, viewport, listener);
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
    const snapshot = await loadSelectionCellSnapshot(selection);
    viewportStore.setCellSnapshot(snapshot ?? emptyCellSnapshot(selection));
    updateSelectionViewport(selection);
    return snapshot ?? emptyCellSnapshot(selection);
  };

  const loadSelectionCellSnapshot = async (
    selection: WorkerRuntimeSelection,
  ): Promise<CellSnapshot | null> => {
    return await invokeWorkerMethod(
      client,
      "getCell",
      isCellSnapshot,
      selection.sheetName,
      selection.address,
    );
  };

  const refreshSelectedCellSnapshot = async (
    selection: WorkerRuntimeSelection = currentSelection,
  ): Promise<void> => {
    const snapshot = await loadSelectionCellSnapshot(selection);
    viewportStore.setCellSnapshot(snapshot ?? emptyCellSnapshot(selection));
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

  const syncSelectionAfterRuntimeState = async (
    runtimeState: WorkbookWorkerStateSnapshot,
  ): Promise<void> => {
    const reconciledSelection = reconcileSelection(currentSelection, runtimeState.sheetNames);
    if (
      reconciledSelection.sheetName !== currentSelection.sheetName ||
      reconciledSelection.address !== currentSelection.address
    ) {
      await applySelection(reconciledSelection);
      return;
    }
    updateSelectionViewport(reconciledSelection);
  };

  const runAuthoritativeRebase = async (): Promise<void> => {
    if (disposed) {
      return;
    }
    const targetRevision = requestedAuthoritativeRevision;
    if (targetRevision <= currentAuthoritativeRevision) {
      return;
    }
    const eventBatch = await loadAuthoritativeEventBatch(
      input.documentId,
      currentAuthoritativeRevision,
      fetchImpl,
    );
    if (
      eventBatch.events.length > 0 &&
      eventBatch.headRevision >= targetRevision &&
      eventBatch.headRevision > currentAuthoritativeRevision
    ) {
      const runtimeState = await invokeWorkerMethod(
        client,
        "applyAuthoritativeEvents",
        isWorkbookWorkerStateSnapshot,
        eventBatch.events,
        eventBatch.headRevision,
      );
      currentAuthoritativeRevision = eventBatch.headRevision;
      publishRuntimeState(runtimeState);
      await syncSelectionAfterRuntimeState(runtimeState);
      return runAuthoritativeRebase();
    }
    publishPhase("recovering");
    const snapshot = await loadLatestWorkbookSnapshot(input.documentId, fetchImpl);
    if (!snapshot) {
      throw new Error("Authoritative workbook snapshot was not available for rebase");
    }
    const runtimeState = await invokeWorkerMethod(
      client,
      "installAuthoritativeSnapshot",
      isWorkbookWorkerStateSnapshot,
      {
        snapshot,
        authoritativeRevision: targetRevision,
        mode: "reconcile",
      } satisfies InstallAuthoritativeSnapshotInput,
    );
    currentAuthoritativeRevision = targetRevision;
    publishRuntimeState(runtimeState);
    await syncSelectionAfterRuntimeState(runtimeState);
    return runAuthoritativeRebase();
  };

  const queueAuthoritativeRebase = (revisionState: WorkbookRevisionState | null): void => {
    if (
      revisionState === null ||
      revisionState.calculatedRevision < revisionState.headRevision ||
      revisionState.headRevision <= currentAuthoritativeRevision
    ) {
      return;
    }
    requestedAuthoritativeRevision = Math.max(
      requestedAuthoritativeRevision,
      revisionState.headRevision,
    );
    const previousRebaseQueue = rebaseQueue;
    rebaseQueue = (async () => {
      await previousRebaseQueue.catch(() => undefined);
      publishPhase("reconciling");
      try {
        await runAuthoritativeRebase();
      } catch (error) {
        if (!disposed) {
          callbacks.onError(toErrorMessage(error));
        }
      } finally {
        if (!disposed) {
          publishPhase("steady");
        }
      }
    })();
  };

  callbacks.onRuntimeState(currentRuntimeState);
  callbacks.onSelection(currentSelection);
  callbacks.onPhase?.(currentPhase);

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
    currentAuthoritativeRevision = await invokeWorkerMethod(
      client,
      "getAuthoritativeRevision",
      isNumber,
    );
    requestedAuthoritativeRevision = currentAuthoritativeRevision;

    if (!bootstrap.restoredFromPersistence || bootstrap.requiresAuthoritativeHydrate) {
      publishPhase("syncing");
      const snapshot = await loadLatestWorkbookSnapshot(input.documentId, fetchImpl);
      if (snapshot) {
        const hydratedState = await invokeWorkerMethod(
          client,
          "installAuthoritativeSnapshot",
          isWorkbookWorkerStateSnapshot,
          {
            snapshot,
            authoritativeRevision: currentAuthoritativeRevision,
            mode: "bootstrap",
          } satisfies InstallAuthoritativeSnapshotInput,
        );
        publishRuntimeState(hydratedState);
      }
    }

    bootstrapped = true;
    publishPhase("steady");
    if (pendingRevisionState) {
      queueAuthoritativeRebase(pendingRevisionState);
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
        if (
          method === "installAuthoritativeSnapshot" &&
          !isInstallAuthoritativeSnapshotInput(args[0])
        ) {
          throw new Error(
            "installAuthoritativeSnapshot requires a valid authoritative snapshot input",
          );
        }
        const result = await client.invoke(method, ...args);
        if (method === "enqueuePendingMutation") {
          await refreshSelectedCellSnapshot();
        }
        if (method === "renderCommit" || method === "installAuthoritativeSnapshot") {
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
