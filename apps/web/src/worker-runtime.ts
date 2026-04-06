import type { CommitOp, EngineReplicaSnapshot } from "@bilig/core";
import { SpreadsheetEngine } from "@bilig/core";
import {
  buildWorkbookAgentPreview,
  isWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewSummary,
} from "@bilig/agent-api";
import type { EngineOpBatch } from "@bilig/workbook-domain";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import {
  createOpfsWorkbookLocalStoreFactory,
  type WorkbookLocalStore,
  type WorkbookLocalStoreFactory,
  WorkbookLocalStoreLockedError,
  type WorkbookStoredState,
} from "@bilig/storage-browser";
import {
  applyWorkbookEvent,
  isAuthoritativeWorkbookEventRecord,
  type AuthoritativeWorkbookEventRecord,
} from "@bilig/zero-sync";
import {
  type CellRangeRef,
  type CellNumberFormatInput,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
  ValueTag,
  type CellSnapshot,
  type EngineEvent,
  type RecalcMetrics,
  type SyncState,
  type WorkbookSnapshot,
  formatCellDisplayValue,
} from "@bilig/protocol";
import {
  encodeViewportPatch,
  type ViewportPatch,
  type ViewportPatchSubscription,
} from "@bilig/worker-transport";
import {
  isPendingWorkbookMutation,
  isPendingWorkbookMutationInput,
  type PendingWorkbookMutation,
  type PendingWorkbookMutationInput,
} from "./workbook-sync.js";
import { applyPendingWorkbookMutationToEngine } from "./worker-runtime-mutation-replay.js";
import {
  collectSheetViewportImpacts,
  normalizeViewport,
  viewportPatchMayBeImpacted,
  type SheetViewportImpact,
  type ViewportSubscriptionState,
  type WorkerEngine,
} from "./worker-runtime-support.js";
import {
  AUTOFIT_CHAR_WIDTH,
  AUTOFIT_PADDING,
  DEFAULT_STYLE_ID,
  MAX_COLUMN_WIDTH,
  MIN_COLUMN_WIDTH,
  buildViewportPatchFromEngine,
  buildViewportPatchFromLocalBase,
} from "./worker-runtime-viewport.js";
import { buildWorkbookLocalAuthoritativeBase } from "./worker-local-base.js";
import { buildWorkbookLocalAuthoritativeDelta } from "./worker-local-authoritative-delta.js";
import {
  buildWorkbookLocalProjectionOverlay,
  collectProjectionOverlayScopeFromEngineEvents,
  createEmptyProjectionOverlayScope,
  mergeProjectionOverlayScopes,
  type ProjectionOverlayScope,
} from "./worker-local-overlay.js";
import { WorkerViewportTileStore } from "./worker-viewport-tile-store.js";

export interface WorkbookWorkerBootstrapOptions {
  documentId: string;
  replicaId: string;
  persistState: boolean;
}

export interface WorkbookWorkerStateSnapshot {
  workbookName: string;
  sheetNames: string[];
  metrics: RecalcMetrics;
  syncState: SyncState;
}

export interface WorkbookWorkerBootstrapResult {
  runtimeState: WorkbookWorkerStateSnapshot;
  restoredFromPersistence: boolean;
  requiresAuthoritativeHydrate: boolean;
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
const DEFERRED_PROJECTION_ENGINE_MIN_CELL_COUNT = 100_000;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isWorkbookSnapshotValue(value: unknown): value is WorkbookSnapshot {
  return (
    isRecord(value) &&
    value["version"] === 1 &&
    isRecord(value["workbook"]) &&
    typeof value["workbook"]["name"] === "string" &&
    Array.isArray(value["sheets"])
  );
}

function isEngineReplicaSnapshotValue(value: unknown): value is EngineReplicaSnapshot {
  return (
    isRecord(value) &&
    isRecord(value["replica"]) &&
    Array.isArray(value["entityVersions"]) &&
    Array.isArray(value["sheetDeleteVersions"])
  );
}

export class WorkbookWorkerRuntime {
  [method: string]: unknown;
  private readonly localStoreFactory: WorkbookLocalStoreFactory;
  private localStore: WorkbookLocalStore | null = null;
  private engine: (SpreadsheetEngine & WorkerEngine) | null = null;
  private projectionEnginePromise: Promise<SpreadsheetEngine & WorkerEngine> | null = null;
  private projectionBuildVersion = 0;
  private authoritativeEngine: SpreadsheetEngine | null = null;
  private authoritativeStateSource: "none" | "memory" | "localStore" = "none";
  private bootstrapOptions: WorkbookWorkerBootstrapOptions | null = null;
  private engineSubscription: (() => void) | null = null;
  private externalSyncState: SyncState | null = null;
  private readonly viewportSubscriptions = new Set<ViewportSubscriptionState>();
  private readonly viewportSubscriptionsBySheet = new Map<string, Set<ViewportSubscriptionState>>();
  private readonly formatIds = new Map<string, number>([["", 0]]);
  private readonly styles = new Map<string, CellStyleRecord>([
    [DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }],
  ]);
  private nextFormatId = 1;
  private snapshotCache: WorkbookSnapshot | null = null;
  private snapshotDirty = true;
  private runtimeStateCache: WorkbookWorkerStateSnapshot | null = null;
  private authoritativeSnapshotCache: WorkbookSnapshot | null = null;
  private authoritativeSnapshotDirty = true;
  private authoritativeReplicaCache: EngineReplicaSnapshot | null = null;
  private authoritativeReplicaDirty = true;
  private persistInFlight: Promise<void> | null = null;
  private persistQueued = false;
  private pendingMutations: PendingWorkbookMutation[] = [];
  private nextPendingMutationSeq = 1;
  private appliedPendingLocalSeq = 0;
  private authoritativeRevision = 0;
  private projectionMatchesLocalStore = false;
  private projectionOverlayScope: ProjectionOverlayScope | null = null;
  private readonly viewportTileStore = new WorkerViewportTileStore();

  constructor(
    options: {
      localStoreFactory?: WorkbookLocalStoreFactory;
    } = {},
  ) {
    this.localStoreFactory = options.localStoreFactory ?? createOpfsWorkbookLocalStoreFactory();
  }

  async ready(): Promise<void> {
    await (await this.getProjectionEngine()).ready();
  }

  dispose(): void {
    this.cleanup();
  }

  async bootstrap(options: WorkbookWorkerBootstrapOptions): Promise<WorkbookWorkerBootstrapResult> {
    this.cleanup();
    this.bootstrapOptions = options;
    this.externalSyncState = null;
    this.runtimeStateCache = null;
    this.snapshotCache = null;
    this.snapshotDirty = true;
    this.pendingMutations = [];
    this.nextPendingMutationSeq = 1;
    this.appliedPendingLocalSeq = 0;
    this.projectionMatchesLocalStore = false;
    this.projectionOverlayScope = null;
    let restoredFromPersistence = false;
    let requiresAuthoritativeHydrate = false;
    let restoredBootstrapState: {
      workbookName: string;
      sheetNames: readonly string[];
      materializedCellCount: number;
      authoritativeRevision: number;
      appliedPendingLocalSeq: number;
    } | null = null;
    let restoredState: WorkbookStoredState | null = null;
    if (options.persistState) {
      try {
        this.localStore = await this.localStoreFactory.open(options.documentId);
        restoredBootstrapState = await this.localStore.loadBootstrapState();
        if (restoredBootstrapState) {
          restoredFromPersistence = true;
          this.authoritativeRevision = restoredBootstrapState.authoritativeRevision;
          this.appliedPendingLocalSeq = restoredBootstrapState.appliedPendingLocalSeq;
        }
      } catch (error) {
        if (!(error instanceof WorkbookLocalStoreLockedError)) {
          throw error;
        }
        this.localStore = null;
      }
      const persistedPendingMutations = this.localStore
        ? await this.localStore.listPendingMutations()
        : [];
      if (persistedPendingMutations.length > 0) {
        this.pendingMutations = persistedPendingMutations.flatMap((mutation) =>
          isPendingWorkbookMutation(mutation) ? [mutation] : [],
        );
      }
      this.nextPendingMutationSeq =
        Math.max(
          this.appliedPendingLocalSeq,
          persistedPendingMutations.reduce((max, mutation) => Math.max(max, mutation.localSeq), 0),
        ) + 1;
    }

    const highestPendingLocalSeq = this.pendingMutations.at(-1)?.localSeq ?? 0;
    if (restoredBootstrapState === null && this.pendingMutations.length > 0) {
      requiresAuthoritativeHydrate = true;
    }

    const projectionMatchesRestoredLocalStore =
      Boolean(this.localStore) &&
      restoredBootstrapState !== null &&
      restoredBootstrapState.appliedPendingLocalSeq === highestPendingLocalSeq;
    this.projectionMatchesLocalStore = projectionMatchesRestoredLocalStore;

    const shouldDeferProjectionEngineBootstrap =
      projectionMatchesRestoredLocalStore &&
      restoredBootstrapState !== null &&
      restoredBootstrapState.materializedCellCount >= DEFERRED_PROJECTION_ENGINE_MIN_CELL_COUNT;

    if (shouldDeferProjectionEngineBootstrap && restoredBootstrapState !== null) {
      this.authoritativeStateSource = "localStore";
      this.runtimeStateCache = this.buildRuntimeStateFromBootstrapState(restoredBootstrapState);
    } else {
      restoredState = this.localStore ? await this.localStore.loadState() : null;
      const parsedRestoredSnapshot = isWorkbookSnapshotValue(restoredState?.snapshot)
        ? restoredState.snapshot
        : null;
      const parsedRestoredReplica = isEngineReplicaSnapshotValue(restoredState?.replica)
        ? restoredState.replica
        : null;
      if (parsedRestoredSnapshot || parsedRestoredReplica) {
        this.installRestoredAuthoritativeState(parsedRestoredSnapshot, parsedRestoredReplica);
      }
      const { engine, overlayScope } = await this.createProjectionEngineFromState({
        snapshot: parsedRestoredSnapshot,
        replica: parsedRestoredReplica,
      });
      this.projectionOverlayScope = overlayScope;
      this.installEngine(engine);
      if (!requiresAuthoritativeHydrate && !projectionMatchesRestoredLocalStore) {
        await this.persistStateNow();
      }
    }

    return {
      runtimeState: this.getRuntimeState(),
      restoredFromPersistence,
      requiresAuthoritativeHydrate,
    };
  }

  getRuntimeState(): WorkbookWorkerStateSnapshot {
    const cachedState = this.runtimeStateCache;
    if (cachedState) {
      return {
        workbookName: cachedState.workbookName,
        sheetNames: [...cachedState.sheetNames],
        metrics: { ...cachedState.metrics },
        syncState: this.externalSyncState ?? cachedState.syncState,
      };
    }
    const engine = this.requireEngine();
    return this.storeRuntimeState({
      workbookName: engine.workbook.workbookName,
      sheetNames: this.listSheetNames(),
      metrics: { ...engine.getLastMetrics() },
      syncState: engine.getSyncState(),
    });
  }

  getAuthoritativeRevision(): number {
    return this.authoritativeRevision;
  }

  async replaceSnapshot(
    snapshot: WorkbookSnapshot,
    authoritativeRevision = this.authoritativeRevision,
  ): Promise<WorkbookWorkerStateSnapshot> {
    this.projectionBuildVersion += 1;
    this.projectionEnginePromise = null;
    const authoritativeEngine = await this.createEngineFromState({
      snapshot,
      replica: null,
    });
    this.authoritativeRevision = Math.max(this.authoritativeRevision, authoritativeRevision);
    this.installAuthoritativeEngine(authoritativeEngine, snapshot, null);
    const { engine, overlayScope } = await this.rebuildProjectionEngine();
    this.projectionOverlayScope = overlayScope;
    this.installEngine(engine);
    this.projectionMatchesLocalStore = false;
    this.viewportTileStore.reset();
    this.invalidateSnapshotCache();
    await this.persistStateNow();
    this.broadcastViewportPatches(null, engine.getLastMetrics());
    return this.getRuntimeState();
  }

  async rebaseToSnapshot(
    snapshot: WorkbookSnapshot,
    authoritativeRevision: number,
  ): Promise<WorkbookWorkerStateSnapshot> {
    this.projectionBuildVersion += 1;
    this.projectionEnginePromise = null;
    const authoritativeEngine = await this.createEngineFromState({
      snapshot,
      replica: null,
    });
    this.authoritativeRevision = authoritativeRevision;
    this.installAuthoritativeEngine(authoritativeEngine, snapshot, null);
    const { engine, overlayScope } = await this.rebuildProjectionEngine();
    this.projectionOverlayScope = overlayScope;
    this.installEngine(engine);
    this.projectionMatchesLocalStore = false;
    this.viewportTileStore.reset();
    this.invalidateSnapshotCache();
    await this.persistStateNow();
    this.broadcastViewportPatches(null, engine.getLastMetrics());
    return this.getRuntimeState();
  }

  async applyAuthoritativeEvents(
    events: readonly AuthoritativeWorkbookEventRecord[],
    authoritativeRevision: number,
  ): Promise<WorkbookWorkerStateSnapshot> {
    this.projectionBuildVersion += 1;
    this.projectionEnginePromise = null;
    if (!events.every((event) => isAuthoritativeWorkbookEventRecord(event))) {
      throw new Error("Invalid authoritative workbook event batch");
    }
    const authoritativeEngine = await this.getAuthoritativeEngine();
    const previousSheets = [...authoritativeEngine.workbook.sheetsByName.values()].map((sheet) => ({
      sheetId: sheet.id,
      name: sheet.name,
    }));
    const authoritativeEngineEvents: EngineEvent[] = [];
    const unsubscribe = authoritativeEngine.subscribe((event) => {
      authoritativeEngineEvents.push(event);
    });
    const absorbedMutationIds = new Set(
      events.flatMap((event) =>
        typeof event.clientMutationId === "string" ? [event.clientMutationId] : [],
      ),
    );
    try {
      events.forEach((event) => {
        applyWorkbookEvent(authoritativeEngine, event.payload);
      });
    } finally {
      unsubscribe();
    }
    if (absorbedMutationIds.size > 0) {
      this.pendingMutations = this.pendingMutations.filter((mutation) => {
        return !absorbedMutationIds.has(mutation.id);
      });
    }
    this.authoritativeRevision = Math.max(this.authoritativeRevision, authoritativeRevision);
    this.invalidateAuthoritativeStateCache();
    const { engine, overlayScope } = await this.rebuildProjectionEngine();
    this.projectionOverlayScope = overlayScope;
    this.installEngine(engine);
    this.invalidateSnapshotCache();
    const localStore = this.bootstrapOptions?.persistState ? this.localStore : null;
    if (localStore) {
      const authoritativeDelta = buildWorkbookLocalAuthoritativeDelta({
        engine: authoritativeEngine,
        payloads: events.map((event) => event.payload),
        engineEvents: authoritativeEngineEvents,
        previousSheets,
      });
      const persisted: WorkbookStoredState = {
        snapshot: this.getAuthoritativeSnapshot(),
        replica: this.getAuthoritativeReplica(),
        authoritativeRevision: this.authoritativeRevision,
        appliedPendingLocalSeq: this.pendingMutations.at(-1)?.localSeq ?? 0,
      };
      await localStore.ingestAuthoritativeDelta({
        state: persisted,
        authoritativeDelta,
        projectionOverlay: buildWorkbookLocalProjectionOverlay({
          authoritativeEngine,
          projectionEngine: engine,
          scope: this.getProjectionOverlayScopeForPersist(),
        }),
        removePendingMutationIds: [...absorbedMutationIds],
      });
      this.projectionMatchesLocalStore = true;
      if (authoritativeDelta.replaceAll) {
        this.viewportTileStore.reset();
      } else {
        authoritativeDelta.replacedSheetIds.forEach((sheetId) => {
          this.viewportTileStore.invalidateSheet(sheetId);
        });
      }
    } else {
      this.projectionMatchesLocalStore = false;
      this.viewportTileStore.reset();
    }
    this.broadcastViewportPatches(null, engine.getLastMetrics());
    return this.getRuntimeState();
  }

  setExternalSyncState(syncState: SyncState | null): WorkbookWorkerStateSnapshot {
    this.externalSyncState = syncState;
    return this.getRuntimeState();
  }

  exportSnapshot(): WorkbookSnapshot {
    if (this.engine) {
      return this.getCachedSnapshot();
    }
    if (
      this.pendingMutations.length === 0 &&
      this.authoritativeSnapshotCache &&
      !this.authoritativeSnapshotDirty
    ) {
      return this.authoritativeSnapshotCache;
    }
    throw new Error("Workbook worker runtime projection snapshot is not ready");
  }

  async previewAgentCommandBundle(
    bundle: WorkbookAgentCommandBundle,
  ): Promise<WorkbookAgentPreviewSummary> {
    if (!isWorkbookAgentCommandBundle(bundle)) {
      throw new Error("Invalid workbook agent preview bundle");
    }
    return await buildWorkbookAgentPreview({
      snapshot: this.engine
        ? this.exportSnapshot()
        : (await this.getProjectionEngine()).exportSnapshot(),
      replicaId: this.requireBootstrapOptions().replicaId,
      bundle,
    });
  }

  listPendingMutations(): PendingWorkbookMutation[] {
    return this.pendingMutations.map((mutation) => ({
      ...mutation,
      args: [...mutation.args],
    }));
  }

  async enqueuePendingMutation(
    input: PendingWorkbookMutationInput,
  ): Promise<PendingWorkbookMutation> {
    if (!isPendingWorkbookMutationInput(input)) {
      throw new Error("Invalid pending workbook mutation");
    }
    const documentId = this.requireBootstrapOptions().documentId;
    const localSeq = this.nextPendingMutationSeq++;
    const nextMutation: PendingWorkbookMutation = {
      id: `${documentId}:pending:${localSeq}`,
      localSeq,
      baseRevision: this.authoritativeRevision,
      method: input.method,
      args: [...input.args],
      enqueuedAtUnixMs: Date.now(),
      submittedAtUnixMs: null,
      status: "pending",
    };
    this.pendingMutations.push(nextMutation);
    this.appliedPendingLocalSeq = localSeq;
    this.projectionMatchesLocalStore = false;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.appendPendingMutation({
        ...nextMutation,
        args: [...nextMutation.args],
      });
    }
    applyPendingWorkbookMutationToEngine(await this.getProjectionEngine(), nextMutation);
    await this.persistStateNow();
    return {
      ...nextMutation,
      args: [...nextMutation.args],
    };
  }

  async markPendingMutationSubmitted(id: string): Promise<void> {
    const pendingMutation = this.pendingMutations.find((mutation) => mutation.id === id);
    if (!pendingMutation || pendingMutation.status === "submitted") {
      return;
    }
    const submittedMutation: PendingWorkbookMutation = {
      ...pendingMutation,
      args: [...pendingMutation.args],
      submittedAtUnixMs: Date.now(),
      status: "submitted",
    };
    this.pendingMutations = this.pendingMutations.map((mutation) =>
      mutation.id === id ? submittedMutation : mutation,
    );
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(submittedMutation);
    }
  }

  async ackPendingMutation(id: string): Promise<void> {
    const nextMutations = this.pendingMutations.filter((mutation) => mutation.id !== id);
    if (nextMutations.length === this.pendingMutations.length) {
      return;
    }
    this.pendingMutations = nextMutations;
    this.projectionMatchesLocalStore = false;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.removePendingMutation(id);
      await this.persistStateNow();
    }
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    const engine = this.engine;
    if (!engine) {
      const localCell = this.readProjectedCellFromLocalStore(sheetName, address);
      if (localCell) {
        this.scheduleProjectionEngineMaterialization();
        return localCell;
      }
      return this.emptyCellSnapshot(sheetName, address);
    }
    if (!engine.workbook.getSheet(sheetName)) {
      return this.emptyCellSnapshot(sheetName, address);
    }
    return engine.getCell(sheetName, address);
  }

  async setCellValue(
    sheetName: string,
    address: string,
    value: CellSnapshot["input"],
  ): Promise<CellSnapshot> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).setCellValue(sheetName, address, value ?? null);
    return this.getCell(sheetName, address);
  }

  async setCellFormula(sheetName: string, address: string, formula: string): Promise<CellSnapshot> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).setCellFormula(sheetName, address, formula);
    return this.getCell(sheetName, address);
  }

  async setRangeStyle(range: CellRangeRef, patch: CellStylePatch): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).setRangeStyle(range, patch);
  }

  async clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).clearRangeStyle(range, fields);
  }

  async setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).setRangeNumberFormat(range, format);
  }

  async clearRangeNumberFormat(range: CellRangeRef): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).clearRangeNumberFormat(range);
  }

  async clearRange(range: CellRangeRef): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).clearRange(range);
  }

  async clearCell(sheetName: string, address: string): Promise<CellSnapshot> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).clearCell(sheetName, address);
    return this.getCell(sheetName, address);
  }

  async renderCommit(ops: CommitOp[]): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).renderCommit(ops);
  }

  async fillRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).fillRange(source, target);
  }

  async copyRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).copyRange(source, target);
  }

  async moveRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).moveRange(source, target);
  }

  async updateColumnWidth(sheetName: string, columnIndex: number, width: number): Promise<number> {
    this.markProjectionDivergedFromLocalStore();
    const clamped = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(width)));
    (await this.getProjectionEngine()).updateColumnMetadata(
      sheetName,
      columnIndex,
      1,
      clamped,
      null,
    );
    return clamped;
  }

  async autofitColumn(sheetName: string, columnIndex: number): Promise<number> {
    const engine = await this.getProjectionEngine();
    const sheet = engine.workbook.getSheet(sheetName);
    let widest = indexToColumn(columnIndex).length * AUTOFIT_CHAR_WIDTH;

    sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
      if (col !== columnIndex) {
        return;
      }
      const display = this.toDisplayText(engine.getCell(sheetName, formatAddress(row, col)));
      widest = Math.max(widest, display.length * AUTOFIT_CHAR_WIDTH);
    });

    return await this.updateColumnWidth(sheetName, columnIndex, widest + AUTOFIT_PADDING);
  }

  subscribe(listener: (event: EngineEvent) => void): () => void {
    return this.requireEngine().subscribe(listener);
  }

  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void {
    return this.requireEngine().subscribeBatches(listener);
  }

  subscribeViewportPatches(
    subscription: ViewportPatchSubscription,
    listener: (patch: Uint8Array) => void,
  ): () => void {
    const state: ViewportSubscriptionState = {
      subscription: normalizeViewport(subscription),
      listener,
      nextVersion: 1,
      knownStyleIds: new Set(),
      lastStyleSignatures: new Map<string, string>(),
      lastCellSignatures: new Map<string, string>(),
      lastColumnSignatures: new Map<number, string>(),
      lastRowSignatures: new Map<number, string>(),
    };
    listener(encodeViewportPatch(this.buildViewportPatch(state, null, this.getCurrentMetrics())));
    if (!this.engine && this.canReadLocalProjectionForViewport()) {
      this.scheduleProjectionEngineMaterialization();
    }
    this.viewportSubscriptions.add(state);
    this.addViewportSubscription(state);
    return () => {
      this.viewportSubscriptions.delete(state);
      this.removeViewportSubscription(state);
    };
  }

  private cleanup(): void {
    this.projectionBuildVersion += 1;
    this.engineSubscription?.();
    this.engineSubscription = null;
    this.projectionEnginePromise = null;
    this.persistQueued = false;
    this.viewportSubscriptions.clear();
    this.viewportSubscriptionsBySheet.clear();
    this.styles.clear();
    this.styles.set(DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID });
    this.snapshotCache = null;
    this.snapshotDirty = true;
    this.runtimeStateCache = null;
    this.authoritativeSnapshotCache = null;
    this.authoritativeSnapshotDirty = true;
    this.authoritativeReplicaCache = null;
    this.authoritativeReplicaDirty = true;
    this.authoritativeStateSource = "none";
    this.pendingMutations = [];
    this.nextPendingMutationSeq = 1;
    this.appliedPendingLocalSeq = 0;
    this.authoritativeRevision = 0;
    this.projectionMatchesLocalStore = false;
    this.projectionOverlayScope = null;
    this.viewportTileStore.reset();
    this.localStore?.close();
    this.localStore = null;
    this.authoritativeEngine = null;
    this.engine = null;
  }

  private requireBootstrapOptions(): WorkbookWorkerBootstrapOptions {
    if (!this.bootstrapOptions) {
      throw new Error("Workbook worker runtime has not been bootstrapped");
    }
    return this.bootstrapOptions;
  }

  private requireEngine(): SpreadsheetEngine & WorkerEngine {
    if (!this.engine) {
      throw new Error("Workbook worker runtime has not been bootstrapped");
    }
    return this.engine;
  }

  private requireAuthoritativeEngine(): SpreadsheetEngine {
    if (!this.authoritativeEngine) {
      throw new Error("Workbook worker runtime has no authoritative base state");
    }
    return this.authoritativeEngine;
  }

  private canPersistState(): boolean {
    return Boolean(this.bootstrapOptions?.persistState && this.localStore);
  }

  private canReadLocalProjectionForViewport(): boolean {
    return Boolean(this.localStore && this.projectionMatchesLocalStore);
  }

  private markProjectionDivergedFromLocalStore(): void {
    this.projectionBuildVersion += 1;
    this.projectionEnginePromise = null;
    this.projectionMatchesLocalStore = false;
    this.viewportTileStore.reset();
  }

  private listSheetNames(): string[] {
    if (this.engine) {
      return [...this.engine.workbook.sheetsByName.values()]
        .toSorted((left, right) => left.order - right.order)
        .map((sheet) => sheet.name);
    }
    return [...(this.runtimeStateCache?.sheetNames ?? ["Sheet1"])];
  }

  private storeRuntimeState(state: WorkbookWorkerStateSnapshot): WorkbookWorkerStateSnapshot {
    this.runtimeStateCache = {
      workbookName: state.workbookName,
      sheetNames: [...state.sheetNames],
      metrics: { ...state.metrics },
      syncState: state.syncState,
    };
    return {
      workbookName: state.workbookName,
      sheetNames: [...state.sheetNames],
      metrics: { ...state.metrics },
      syncState: this.externalSyncState ?? state.syncState,
    };
  }

  private buildRuntimeStateFromBootstrapState(state: {
    workbookName: string;
    sheetNames: readonly string[];
    materializedCellCount: number;
    authoritativeRevision: number;
    appliedPendingLocalSeq: number;
  }): WorkbookWorkerStateSnapshot {
    return this.storeRuntimeState({
      workbookName: state.workbookName,
      sheetNames: [...state.sheetNames],
      metrics: { ...EMPTY_METRICS },
      syncState: "syncing",
    });
  }

  private updateRuntimeStateFromEngine(
    engine: SpreadsheetEngine & WorkerEngine = this.requireEngine(),
  ): WorkbookWorkerStateSnapshot {
    return this.storeRuntimeState({
      workbookName: engine.workbook.workbookName,
      sheetNames: [...engine.workbook.sheetsByName.values()]
        .toSorted((left, right) => left.order - right.order)
        .map((sheet) => sheet.name),
      metrics: { ...engine.getLastMetrics() },
      syncState: engine.getSyncState(),
    });
  }

  private getCurrentMetrics(): RecalcMetrics {
    if (this.engine) {
      return { ...this.engine.getLastMetrics() };
    }
    return { ...(this.runtimeStateCache?.metrics ?? EMPTY_METRICS) };
  }

  private async getAuthoritativeStateInput(): Promise<{
    snapshot: WorkbookSnapshot | null;
    replica: EngineReplicaSnapshot | null;
  }> {
    if (this.authoritativeStateSource === "localStore") {
      const restoredState = this.localStore ? await this.localStore.loadState() : null;
      const restoredSnapshot = isWorkbookSnapshotValue(restoredState?.snapshot)
        ? restoredState.snapshot
        : null;
      const restoredReplica = isEngineReplicaSnapshotValue(restoredState?.replica)
        ? restoredState.replica
        : null;
      this.installRestoredAuthoritativeState(restoredSnapshot, restoredReplica);
    }
    const snapshot =
      this.authoritativeSnapshotDirty && this.authoritativeEngine
        ? this.storeCachedAuthoritativeSnapshot(this.authoritativeEngine.exportSnapshot())
        : this.authoritativeSnapshotCache;
    const replica =
      this.authoritativeReplicaDirty && this.authoritativeEngine
        ? this.storeCachedAuthoritativeReplica(this.authoritativeEngine.exportReplicaSnapshot())
        : this.authoritativeReplicaCache;
    return {
      snapshot,
      replica,
    };
  }

  private readProjectedCellFromLocalStore(sheetName: string, address: string): CellSnapshot | null {
    if (!this.canReadLocalProjectionForViewport() || !this.localStore) {
      return null;
    }
    const parsed = parseCellAddress(address, sheetName);
    const localBase = this.viewportTileStore.readViewport({
      localStore: this.localStore,
      sheetName,
      viewport: {
        rowStart: parsed.row,
        rowEnd: parsed.row,
        colStart: parsed.col,
        colEnd: parsed.col,
      },
    });
    const cell = localBase?.cells.find((entry) => entry.snapshot.address === address);
    return cell ? structuredClone(cell.snapshot) : null;
  }

  private async createEngineFromState(input: {
    snapshot: WorkbookSnapshot | null;
    replica: EngineReplicaSnapshot | null;
  }): Promise<SpreadsheetEngine> {
    const options = this.requireBootstrapOptions();
    const engine = new SpreadsheetEngine({
      workbookName: options.documentId,
      replicaId: options.replicaId,
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

  private async createProjectionEngineFromState(input: {
    snapshot: WorkbookSnapshot | null;
    replica: EngineReplicaSnapshot | null;
  }): Promise<{
    engine: SpreadsheetEngine;
    overlayScope: ProjectionOverlayScope | null;
  }> {
    const engine = await this.createEngineFromState(input);
    if (this.pendingMutations.length === 0) {
      return { engine, overlayScope: null };
    }
    const replayEvents: EngineEvent[] = [];
    const unsubscribe = engine.subscribe((event) => {
      replayEvents.push(event);
    });
    try {
      this.pendingMutations.forEach((mutation) => {
        applyPendingWorkbookMutationToEngine(engine, mutation);
      });
    } finally {
      unsubscribe();
    }
    return {
      engine,
      overlayScope: collectProjectionOverlayScopeFromEngineEvents(engine, replayEvents),
    };
  }

  private installAuthoritativeEngine(
    engine: SpreadsheetEngine,
    snapshot: WorkbookSnapshot | null,
    replica: EngineReplicaSnapshot | null,
  ): void {
    this.authoritativeStateSource = "memory";
    this.authoritativeEngine = engine;
    this.storeCachedAuthoritativeSnapshot(snapshot ?? engine.exportSnapshot());
    this.storeCachedAuthoritativeReplica(replica ?? engine.exportReplicaSnapshot());
  }

  private installRestoredAuthoritativeState(
    snapshot: WorkbookSnapshot | null,
    replica: EngineReplicaSnapshot | null,
  ): void {
    this.authoritativeStateSource = "memory";
    this.authoritativeEngine = null;
    this.authoritativeSnapshotCache = snapshot;
    this.authoritativeSnapshotDirty = snapshot === null;
    this.authoritativeReplicaCache = replica;
    this.authoritativeReplicaDirty = replica === null;
  }

  private async getAuthoritativeEngine(): Promise<SpreadsheetEngine> {
    if (this.authoritativeEngine) {
      return this.authoritativeEngine;
    }
    const authoritativeState = await this.getAuthoritativeStateInput();
    const engine = await this.createEngineFromState({
      snapshot: authoritativeState.snapshot,
      replica: authoritativeState.replica,
    });
    this.authoritativeEngine = engine;
    if (authoritativeState.snapshot) {
      this.authoritativeSnapshotDirty = false;
    } else {
      this.storeCachedAuthoritativeSnapshot(engine.exportSnapshot());
    }
    if (authoritativeState.replica) {
      this.authoritativeReplicaDirty = false;
    } else {
      this.storeCachedAuthoritativeReplica(engine.exportReplicaSnapshot());
    }
    return engine;
  }

  private async rebuildProjectionEngine(): Promise<{
    engine: SpreadsheetEngine;
    overlayScope: ProjectionOverlayScope | null;
  }> {
    const authoritativeState = await this.getAuthoritativeStateInput();
    return await this.createProjectionEngineFromState(authoritativeState);
  }

  private async getProjectionEngine(): Promise<SpreadsheetEngine & WorkerEngine> {
    if (this.engine) {
      return this.engine;
    }
    if (this.projectionEnginePromise) {
      return await this.projectionEnginePromise;
    }
    const buildVersion = this.projectionBuildVersion;
    const buildPromise = (async () => {
      const { engine, overlayScope } = await this.rebuildProjectionEngine();
      if (buildVersion !== this.projectionBuildVersion) {
        return this.engine ?? engine;
      }
      this.projectionOverlayScope = overlayScope;
      this.installEngine(engine);
      return this.requireEngine();
    })();
    this.projectionEnginePromise = buildPromise;
    try {
      return await buildPromise;
    } finally {
      if (this.projectionEnginePromise === buildPromise) {
        this.projectionEnginePromise = null;
      }
    }
  }

  private scheduleProjectionEngineMaterialization(): void {
    if (this.engine || this.projectionEnginePromise || !this.bootstrapOptions) {
      return;
    }
    const scheduledVersion = this.projectionBuildVersion;
    setTimeout(() => {
      if (scheduledVersion !== this.projectionBuildVersion || this.engine) {
        return;
      }
      void this.getProjectionEngine().catch(() => undefined);
    }, 0);
  }

  private installEngine(engine: SpreadsheetEngine & WorkerEngine): void {
    this.engineSubscription?.();
    this.engine = engine;
    this.updateRuntimeStateFromEngine(engine);
    this.engineSubscription = engine.subscribe((event) => {
      this.invalidateSnapshotCache();
      if (this.pendingMutations.length > 0) {
        this.projectionOverlayScope = mergeProjectionOverlayScopes(
          this.projectionOverlayScope,
          collectProjectionOverlayScopeFromEngineEvents(engine, [event]),
        );
      }
      this.updateRuntimeStateFromEngine(engine);
      this.broadcastViewportPatches(event);
    });
  }

  private getProjectionOverlayScopeForPersist(): ProjectionOverlayScope | null {
    if (this.projectionOverlayScope) {
      return this.projectionOverlayScope;
    }
    return this.pendingMutations.length === 0 ? createEmptyProjectionOverlayScope() : null;
  }

  private async persistStateNow(): Promise<void> {
    if (!this.canPersistState()) {
      return;
    }
    this.persistQueued = true;
    await this.flushPersistState();
  }

  private async flushPersistState(): Promise<void> {
    if (!this.persistQueued || !this.canPersistState()) {
      return;
    }
    if (this.persistInFlight) {
      await this.persistInFlight;
      if (this.persistQueued) {
        await this.flushPersistState();
      }
      return;
    }

    const authoritativeEngine = await this.getAuthoritativeEngine();
    const projectionEngine = await this.getProjectionEngine();
    const persisted: WorkbookStoredState = {
      snapshot: this.getAuthoritativeSnapshot(),
      replica: this.getAuthoritativeReplica(),
      authoritativeRevision: this.authoritativeRevision,
      appliedPendingLocalSeq: this.pendingMutations.at(-1)?.localSeq ?? 0,
    };
    this.persistQueued = false;
    const localStore = this.localStore;
    if (!localStore) {
      return;
    }
    const savePromise = localStore.persistProjectionState({
      state: persisted,
      authoritativeBase: buildWorkbookLocalAuthoritativeBase(authoritativeEngine),
      projectionOverlay: buildWorkbookLocalProjectionOverlay({
        authoritativeEngine,
        projectionEngine,
        scope: this.getProjectionOverlayScopeForPersist(),
      }),
    });
    this.persistInFlight = savePromise;
    try {
      await savePromise;
      this.projectionMatchesLocalStore = true;
    } finally {
      if (this.persistInFlight === savePromise) {
        this.persistInFlight = null;
      }
    }
    if (this.persistQueued) {
      await this.flushPersistState();
    }
  }

  private broadcastViewportPatches(
    event: EngineEvent | null,
    metrics: RecalcMetrics = this.getCurrentMetrics(),
  ): void {
    const impactsBySheet =
      event === null ? null : collectSheetViewportImpacts(this.requireEngine(), event);
    const impactedSheets = impactsBySheet === null ? null : new Set(impactsBySheet.keys());
    for (const subscription of this.getViewportSubscriptionsForEvent(event, impactedSheets)) {
      const sheetImpact = impactsBySheet?.get(subscription.subscription.sheetName) ?? null;
      if (
        event !== null &&
        !viewportPatchMayBeImpacted(subscription.subscription, event, sheetImpact, impactedSheets)
      ) {
        continue;
      }
      const patch = this.buildViewportPatch(subscription, event, metrics, sheetImpact);
      if (patch.cells.length === 0 && patch.columns.length === 0 && patch.rows.length === 0) {
        continue;
      }
      subscription.listener(encodeViewportPatch(patch));
    }
  }

  private addViewportSubscription(state: ViewportSubscriptionState): void {
    const subscriptions =
      this.viewportSubscriptionsBySheet.get(state.subscription.sheetName) ?? new Set();
    subscriptions.add(state);
    this.viewportSubscriptionsBySheet.set(state.subscription.sheetName, subscriptions);
  }

  private removeViewportSubscription(state: ViewportSubscriptionState): void {
    const subscriptions = this.viewportSubscriptionsBySheet.get(state.subscription.sheetName);
    if (!subscriptions) {
      return;
    }
    subscriptions.delete(state);
    if (subscriptions.size === 0) {
      this.viewportSubscriptionsBySheet.delete(state.subscription.sheetName);
    }
  }

  private getViewportSubscriptionsForEvent(
    event: EngineEvent | null,
    impactedSheets: ReadonlySet<string> | null,
  ): Iterable<ViewportSubscriptionState> {
    if (event === null || event.invalidation === "full" || impactedSheets === null) {
      return this.viewportSubscriptions;
    }

    const subscriptions = new Set<ViewportSubscriptionState>();
    impactedSheets.forEach((sheetName) => {
      this.viewportSubscriptionsBySheet.get(sheetName)?.forEach((subscription) => {
        subscriptions.add(subscription);
      });
    });
    return subscriptions;
  }

  private buildViewportPatch(
    state: ViewportSubscriptionState,
    event: EngineEvent | null,
    metrics: RecalcMetrics = this.getCurrentMetrics(),
    sheetImpact: SheetViewportImpact | null = null,
  ): ViewportPatch {
    if (
      (event === null || event.invalidation === "full") &&
      this.canReadLocalProjectionForViewport()
    ) {
      const localBase =
        this.localStore === null
          ? null
          : this.viewportTileStore.readViewport({
              localStore: this.localStore,
              sheetName: state.subscription.sheetName,
              viewport: state.subscription,
            });
      if (localBase) {
        return buildViewportPatchFromLocalBase({
          state,
          metrics,
          base: localBase,
          getFormatId: (format) => this.getFormatId(format),
        });
      }
    }

    return buildViewportPatchFromEngine({
      state,
      event,
      metrics,
      sheetImpact,
      engine: this.requireEngine(),
      emptyCellSnapshot: (sheetName, address) => this.emptyCellSnapshot(sheetName, address),
      getStyleRecord: (styleId) => this.getStyleRecord(styleId),
      getFormatId: (format) => this.getFormatId(format),
    });
  }

  private getStyleRecord(styleId: string): CellStyleRecord {
    const existing = this.styles.get(styleId);
    if (existing) {
      return existing;
    }
    const resolved = this.requireEngine().getCellStyle(styleId) ?? { id: DEFAULT_STYLE_ID };
    this.styles.set(resolved.id, resolved);
    return resolved;
  }

  private getFormatId(format: string | undefined): number {
    const key = format ?? "";
    const existing = this.formatIds.get(key);
    if (existing !== undefined) {
      return existing;
    }
    const nextId = this.nextFormatId++;
    this.formatIds.set(key, nextId);
    return nextId;
  }

  private toDisplayText(snapshot: CellSnapshot): string {
    return formatCellDisplayValue(snapshot.value, snapshot.format);
  }

  private emptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
    return {
      sheetName,
      address,
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
    };
  }

  private invalidateSnapshotCache(): void {
    this.snapshotDirty = true;
  }

  private storeCachedSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
    this.snapshotCache = snapshot;
    this.snapshotDirty = false;
    return snapshot;
  }

  private getCachedSnapshot(): WorkbookSnapshot {
    if (this.snapshotCache && !this.snapshotDirty) {
      return this.snapshotCache;
    }
    return this.storeCachedSnapshot(this.requireEngine().exportSnapshot());
  }

  private invalidateAuthoritativeStateCache(): void {
    this.authoritativeSnapshotDirty = true;
    this.authoritativeReplicaDirty = true;
  }

  private storeCachedAuthoritativeSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
    this.authoritativeSnapshotCache = snapshot;
    this.authoritativeSnapshotDirty = false;
    return snapshot;
  }

  private getAuthoritativeSnapshot(): WorkbookSnapshot {
    if (this.authoritativeSnapshotCache && !this.authoritativeSnapshotDirty) {
      return this.authoritativeSnapshotCache;
    }
    if (!this.authoritativeEngine && this.pendingMutations.length === 0 && this.engine) {
      return this.storeCachedAuthoritativeSnapshot(this.engine.exportSnapshot());
    }
    return this.storeCachedAuthoritativeSnapshot(
      this.requireAuthoritativeEngine().exportSnapshot(),
    );
  }

  private storeCachedAuthoritativeReplica(replica: EngineReplicaSnapshot): EngineReplicaSnapshot {
    this.authoritativeReplicaCache = replica;
    this.authoritativeReplicaDirty = false;
    return replica;
  }

  private getAuthoritativeReplica(): EngineReplicaSnapshot {
    if (this.authoritativeReplicaCache && !this.authoritativeReplicaDirty) {
      return this.authoritativeReplicaCache;
    }
    if (!this.authoritativeEngine && this.pendingMutations.length === 0 && this.engine) {
      return this.storeCachedAuthoritativeReplica(this.engine.exportReplicaSnapshot());
    }
    return this.storeCachedAuthoritativeReplica(
      this.requireAuthoritativeEngine().exportReplicaSnapshot(),
    );
  }
}
