import type { CommitOp, EngineReplicaSnapshot } from "@bilig/core";
import { isEngineReplicaSnapshot, SpreadsheetEngine } from "@bilig/core";
import { Effect } from "effect";
import {
  buildWorkbookAgentPreview,
  isWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewSummary,
} from "@bilig/agent-api";
import type { EngineOpBatch } from "@bilig/workbook-domain";
import { formatAddress, indexToColumn } from "@bilig/formula";
import {
  createOpfsWorkbookLocalStoreFactory,
  type WorkbookLocalStore,
  type WorkbookLocalStoreFactory,
  type WorkbookStoredState,
} from "@bilig/storage-browser";
import {
  applyWorkbookEvent,
  isAuthoritativeWorkbookEventRecord,
  type AuthoritativeWorkbookEventRecord,
} from "@bilig/zero-sync";
import {
  isWorkbookSnapshot,
  type CellRangeRef,
  type CellNumberFormatInput,
  type CellStyleField,
  type CellStylePatch,
  type CellSnapshot,
  type EngineEvent,
  type RecalcMetrics,
  type SyncState,
  type WorkbookDefinedNameSnapshot,
  type WorkbookSnapshot,
  formatCellDisplayValue,
} from "@bilig/protocol";
import type { ViewportPatch, ViewportPatchSubscription } from "@bilig/worker-transport";
import {
  isPendingWorkbookMutationInput,
  type PendingWorkbookMutation,
  type PendingWorkbookMutationInput,
} from "./workbook-sync.js";
import {
  clonePendingWorkbookMutation,
  markPendingWorkbookMutationAcked,
  markPendingWorkbookMutationFailed,
  markPendingWorkbookMutationSubmitted,
  queuePendingWorkbookMutationRetry,
  recordPendingWorkbookMutationAttempt,
} from "./workbook-mutation-journal.js";
import {
  buildPendingMutationSummary,
  markJournalMutationsRebased,
  replaceJournalMutation,
  syncPendingMutationsFromJournal,
} from "./worker-runtime-pending-mutations.js";
import { applyPendingWorkbookMutationToEngine } from "./worker-runtime-mutation-replay.js";
import {
  ensureAuthoritativeEngine,
  installAuthoritativeEngineState,
  installRestoredAuthoritativeState,
  readProjectedCellFromLocalStore,
  rebuildProjectionEngine,
  resolveAuthoritativeStateInput,
} from "./worker-runtime-engine-access.js";
import { restoreBootstrapPersistence } from "./worker-runtime-bootstrap-persistence.js";
import {
  createProjectionEngineFromState,
  createWorkbookEngineFromState,
} from "./worker-runtime-engine-state.js";
import {
  EMPTY_RUNTIME_METRICS,
  buildWorkerRuntimeStateFromBootstrap,
  buildWorkerRuntimeStateFromEngine,
  cloneRuntimeMetrics,
  cloneWorkerRuntimeState,
  withExternalSyncState,
} from "./worker-runtime-state.js";
import { WorkerRuntimeSnapshotCaches } from "./worker-runtime-snapshot-caches.js";
import {
  collectSheetViewportImpacts,
  type SheetViewportImpact,
  type ViewportSubscriptionState,
  type WorkerEngine,
} from "./worker-runtime-support.js";
import {
  AUTOFIT_CHAR_WIDTH,
  AUTOFIT_PADDING,
  MAX_COLUMN_WIDTH,
  MIN_COLUMN_WIDTH,
} from "./worker-runtime-viewport.js";
import {
  WorkerViewportPatchPublisher,
  createEmptyCellSnapshot,
} from "./worker-runtime-viewport-publisher.js";
import { buildWorkbookLocalAuthoritativeDelta } from "./worker-local-authoritative-delta.js";
import {
  collectProjectionOverlayScopeFromEngineEvents,
  createEmptyProjectionOverlayScope,
  mergeProjectionOverlayScopes,
  type ProjectionOverlayScope,
} from "./worker-local-overlay.js";
import {
  buildPersistedWorkerState,
  ingestAuthoritativeDeltaToLocalStore,
  persistProjectionStateToLocalStore,
} from "./worker-runtime-local-persistence.js";
import { WorkerViewportTileStore } from "./worker-viewport-tile-store.js";

export interface WorkbookWorkerBootstrapOptions {
  documentId: string;
  replicaId: string;
  persistState: boolean;
}

export interface WorkbookWorkerStateSnapshot {
  workbookName: string;
  sheetNames: string[];
  definedNames: WorkbookDefinedNameSnapshot[];
  metrics: RecalcMetrics;
  syncState: SyncState;
  pendingMutationSummary?: WorkbookPendingMutationSummarySnapshot;
}

export interface WorkbookFailedPendingMutationSnapshot {
  readonly id: string;
  readonly method: string;
  readonly failureMessage: string;
  readonly attemptCount: number;
}

export interface WorkbookPendingMutationSummarySnapshot {
  readonly activeCount: number;
  readonly failedCount: number;
  readonly firstFailed: WorkbookFailedPendingMutationSnapshot | null;
}

export interface WorkbookWorkerBootstrapResult {
  runtimeState: WorkbookWorkerStateSnapshot;
  restoredFromPersistence: boolean;
  requiresAuthoritativeHydrate: boolean;
}

export interface InstallAuthoritativeSnapshotInput {
  readonly snapshot: WorkbookSnapshot;
  readonly authoritativeRevision: number;
  readonly mode: "bootstrap" | "reconcile";
}

const DEFERRED_PROJECTION_ENGINE_MIN_CELL_COUNT = 100_000;

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
  private runtimeStateCache: WorkbookWorkerStateSnapshot | null = null;
  private persistInFlight: Promise<void> | null = null;
  private persistQueued = false;
  private mutationJournalEntries: PendingWorkbookMutation[] = [];
  private pendingMutations: PendingWorkbookMutation[] = [];
  private nextPendingMutationSeq = 1;
  private authoritativeRevision = 0;
  private projectionMatchesLocalStore = false;
  private projectionOverlayScope: ProjectionOverlayScope | null = null;
  private readonly snapshotCaches = new WorkerRuntimeSnapshotCaches();
  private readonly viewportTileStore = new WorkerViewportTileStore();
  private readonly viewportPatchPublisher = new WorkerViewportPatchPublisher({
    buildPatch: (state, event, metrics, sheetImpact) =>
      this.buildViewportPatch(state, event, metrics, sheetImpact),
    canReadLocalProjectionForViewport: () => this.canReadLocalProjectionForViewport(),
    getCurrentMetrics: () => this.getCurrentMetrics(),
    getProjectionEngine: () => this.requireEngine(),
    hasProjectionEngine: () => this.engine !== null,
    readLocalViewport: (sheetName, viewport) => {
      if (!this.localStore) {
        return null;
      }
      return this.viewportTileStore.readViewport({
        localStore: this.localStore,
        sheetName,
        viewport,
      });
    },
    scheduleProjectionEngineMaterialization: () => this.scheduleProjectionEngineMaterialization(),
  });

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
    this.snapshotCaches.reset();
    this.mutationJournalEntries = [];
    this.pendingMutations = [];
    this.nextPendingMutationSeq = 1;
    this.projectionMatchesLocalStore = false;
    this.projectionOverlayScope = null;
    let requiresAuthoritativeHydrate = false;
    const restoredPersistence = await restoreBootstrapPersistence({
      persistState: options.persistState,
      documentId: options.documentId,
      localStoreFactory: this.localStoreFactory,
    });
    this.localStore = restoredPersistence.localStore;
    this.authoritativeRevision = restoredPersistence.authoritativeRevision;
    this.mutationJournalEntries = restoredPersistence.mutationJournalEntries;
    this.pendingMutations = restoredPersistence.pendingMutations;
    this.nextPendingMutationSeq = restoredPersistence.nextPendingMutationSeq;
    const restoredFromPersistence = restoredPersistence.restoredFromPersistence;
    const restoredBootstrapState = restoredPersistence.restoredBootstrapState;
    let restoredState: WorkbookStoredState | null = null;

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
      this.runtimeStateCache = buildWorkerRuntimeStateFromBootstrap(restoredBootstrapState);
    } else {
      restoredState = this.localStore ? await this.localStore.loadState() : null;
      const parsedRestoredSnapshot = isWorkbookSnapshot(restoredState?.snapshot)
        ? restoredState.snapshot
        : null;
      const parsedRestoredReplica = isEngineReplicaSnapshot(restoredState?.replica)
        ? restoredState.replica
        : null;
      if (parsedRestoredSnapshot || parsedRestoredReplica) {
        this.installRestoredAuthoritativeState(parsedRestoredSnapshot, parsedRestoredReplica);
      }
      const { engine, overlayScope } = await createProjectionEngineFromState({
        workbookName: options.documentId,
        replicaId: options.replicaId,
        snapshot: parsedRestoredSnapshot,
        replica: parsedRestoredReplica,
        pendingMutations: this.pendingMutations,
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
      return withExternalSyncState(
        {
          workbookName: cachedState.workbookName,
          sheetNames: cachedState.sheetNames,
          definedNames: cachedState.definedNames,
          metrics: cachedState.metrics,
          syncState: cachedState.syncState,
          pendingMutationSummary: this.buildPendingMutationSummary(),
        },
        this.externalSyncState,
      );
    }
    const engine = this.requireEngine();
    return this.storeRuntimeState({
      ...buildWorkerRuntimeStateFromEngine(engine),
    });
  }

  getAuthoritativeRevision(): number {
    return this.authoritativeRevision;
  }

  async installAuthoritativeSnapshot(
    input: InstallAuthoritativeSnapshotInput,
  ): Promise<WorkbookWorkerStateSnapshot> {
    const { snapshot, authoritativeRevision, mode } = input;
    if (mode !== "bootstrap" && mode !== "reconcile") {
      throw new Error("Invalid authoritative snapshot install mode");
    }
    return this.installAuthoritativeSnapshotInternal(snapshot, authoritativeRevision, mode);
  }

  private async installAuthoritativeSnapshotInternal(
    snapshot: WorkbookSnapshot,
    authoritativeRevision: number,
    mode: InstallAuthoritativeSnapshotInput["mode"],
  ): Promise<WorkbookWorkerStateSnapshot> {
    this.projectionBuildVersion += 1;
    this.projectionEnginePromise = null;
    const options = this.requireBootstrapOptions();
    const authoritativeEngine = await createWorkbookEngineFromState({
      workbookName: options.documentId,
      replicaId: options.replicaId,
      snapshot,
      replica: null,
    });
    this.authoritativeRevision =
      mode === "bootstrap"
        ? Math.max(this.authoritativeRevision, authoritativeRevision)
        : authoritativeRevision;
    this.installAuthoritativeEngine(authoritativeEngine, snapshot, null);
    const { engine, overlayScope } = await this.rebuildProjectionEngine();
    if (mode === "reconcile" && this.pendingMutations.length > 0) {
      await this.markRemainingJournalMutationsRebased();
    }
    this.projectionOverlayScope = overlayScope;
    this.installEngine(engine);
    this.projectionMatchesLocalStore = false;
    this.viewportTileStore.reset();
    this.snapshotCaches.invalidateProjectionSnapshot();
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
      const ackedAtUnixMs = Date.now();
      this.mutationJournalEntries = this.mutationJournalEntries.map((mutation) => {
        if (!absorbedMutationIds.has(mutation.id)) {
          return mutation;
        }
        return Effect.runSync(markPendingWorkbookMutationAcked(mutation, ackedAtUnixMs));
      });
      this.syncPendingMutationsFromJournal();
    }
    this.authoritativeRevision = Math.max(this.authoritativeRevision, authoritativeRevision);
    this.snapshotCaches.invalidateAuthoritativeState();
    const { engine, overlayScope } = await this.rebuildProjectionEngine();
    if (events.length > 0 && this.pendingMutations.length > 0) {
      await this.markRemainingJournalMutationsRebased();
    }
    this.projectionOverlayScope = overlayScope;
    this.installEngine(engine);
    this.snapshotCaches.invalidateProjectionSnapshot();
    const localStore = this.bootstrapOptions?.persistState ? this.localStore : null;
    if (localStore) {
      const authoritativeDelta = buildWorkbookLocalAuthoritativeDelta({
        engine: authoritativeEngine,
        payloads: events.map((event) => event.payload),
        engineEvents: authoritativeEngineEvents,
        previousSheets,
      });
      const persisted = buildPersistedWorkerState({
        snapshotCaches: this.snapshotCaches,
        authoritativeEngine,
        projectionEngine: engine,
        hasDedicatedAuthoritativeEngine: this.authoritativeEngine !== null,
        authoritativeRevision: this.authoritativeRevision,
        appliedPendingLocalSeq: this.pendingMutations.at(-1)?.localSeq ?? 0,
      });
      await ingestAuthoritativeDeltaToLocalStore({
        localStore,
        state: persisted,
        authoritativeDelta,
        authoritativeEngine,
        projectionEngine: engine,
        projectionOverlayScope: this.getProjectionOverlayScopeForPersist(),
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
      return this.snapshotCaches.getProjectionSnapshot(() => this.requireEngine().exportSnapshot());
    }
    const readyAuthoritativeSnapshot =
      this.pendingMutations.length === 0
        ? this.snapshotCaches.getReadyAuthoritativeSnapshot()
        : null;
    if (readyAuthoritativeSnapshot) {
      return readyAuthoritativeSnapshot;
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
    return this.pendingMutations.map(clonePendingWorkbookMutation);
  }

  listMutationJournalEntries(): PendingWorkbookMutation[] {
    return this.mutationJournalEntries.map(clonePendingWorkbookMutation);
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
      lastAttemptedAtUnixMs: null,
      ackedAtUnixMs: null,
      rebasedAtUnixMs: null,
      failedAtUnixMs: null,
      attemptCount: 0,
      failureMessage: null,
      status: "local",
    };
    this.mutationJournalEntries.push(nextMutation);
    this.syncPendingMutationsFromJournal();
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
    const pendingMutation = this.mutationJournalEntries.find((mutation) => mutation.id === id);
    if (!pendingMutation || pendingMutation.status === "submitted") {
      return;
    }
    const submittedMutation = Effect.runSync(
      markPendingWorkbookMutationSubmitted(pendingMutation, Date.now()),
    );
    this.replaceJournalMutation(submittedMutation);
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(submittedMutation);
    }
  }

  async ackPendingMutation(id: string): Promise<void> {
    const pendingMutation = this.mutationJournalEntries.find((mutation) => mutation.id === id);
    if (!pendingMutation) {
      return;
    }
    const ackedMutation = Effect.runSync(
      markPendingWorkbookMutationAcked(pendingMutation, Date.now()),
    );
    this.replaceJournalMutation(ackedMutation);
    this.projectionMatchesLocalStore = false;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(ackedMutation);
      await this.persistStateNow();
    }
  }

  async recordPendingMutationAttempt(id: string): Promise<void> {
    const pendingMutation = this.mutationJournalEntries.find((mutation) => mutation.id === id);
    if (!pendingMutation) {
      return;
    }
    const attemptedMutation = Effect.runSync(
      recordPendingWorkbookMutationAttempt(pendingMutation, Date.now()),
    );
    this.replaceJournalMutation(attemptedMutation);
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(attemptedMutation);
    }
  }

  async markPendingMutationFailed(id: string, failureMessage: string): Promise<void> {
    const pendingMutation = this.mutationJournalEntries.find((mutation) => mutation.id === id);
    if (!pendingMutation) {
      return;
    }
    const failedMutation = Effect.runSync(
      markPendingWorkbookMutationFailed(pendingMutation, failureMessage, Date.now()),
    );
    this.replaceJournalMutation(failedMutation);
    this.projectionMatchesLocalStore = false;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(failedMutation);
      await this.persistStateNow();
    }
  }

  async retryPendingMutation(id: string): Promise<void> {
    const pendingMutation = this.mutationJournalEntries.find((mutation) => mutation.id === id);
    if (!pendingMutation) {
      return;
    }
    const retriedMutation = Effect.runSync(queuePendingWorkbookMutationRetry(pendingMutation));
    this.replaceJournalMutation(retriedMutation);
    this.projectionMatchesLocalStore = false;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(retriedMutation);
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
      return createEmptyCellSnapshot(sheetName, address);
    }
    if (!engine.workbook.getSheet(sheetName)) {
      return createEmptyCellSnapshot(sheetName, address);
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

  async updateRowMetadata(
    sheetName: string,
    startRow: number,
    count: number,
    height: number | null,
    hidden: boolean | null,
  ): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    const normalizedHeight = height === null ? null : Math.max(1, Math.round(height));
    (await this.getProjectionEngine()).updateRowMetadata(
      sheetName,
      startRow,
      count,
      normalizedHeight,
      hidden,
    );
  }

  async updateColumnMetadata(
    sheetName: string,
    startCol: number,
    count: number,
    width: number | null,
    hidden: boolean | null,
  ): Promise<number | null> {
    this.markProjectionDivergedFromLocalStore();
    const normalizedWidth =
      width === null
        ? null
        : Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(width)));
    (await this.getProjectionEngine()).updateColumnMetadata(
      sheetName,
      startCol,
      count,
      normalizedWidth,
      hidden,
    );
    return normalizedWidth;
  }

  async updateColumnWidth(sheetName: string, columnIndex: number, width: number): Promise<number> {
    const normalizedWidth = await this.updateColumnMetadata(sheetName, columnIndex, 1, width, null);
    return normalizedWidth ?? width;
  }

  async setFreezePane(sheetName: string, rows: number, cols: number): Promise<void> {
    this.markProjectionDivergedFromLocalStore();
    (await this.getProjectionEngine()).setFreezePane(
      sheetName,
      Math.max(0, Math.round(rows)),
      Math.max(0, Math.round(cols)),
    );
  }

  async autofitColumn(sheetName: string, columnIndex: number): Promise<number> {
    const engine = await this.getProjectionEngine();
    const sheet = engine.workbook.getSheet(sheetName);
    let widest = indexToColumn(columnIndex).length * AUTOFIT_CHAR_WIDTH;

    sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
      if (col !== columnIndex) {
        return;
      }
      const snapshot = engine.getCell(sheetName, formatAddress(row, col));
      const display = formatCellDisplayValue(snapshot.value, snapshot.format);
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
    return this.viewportPatchPublisher.subscribe(subscription, listener);
  }

  private cleanup(): void {
    this.projectionBuildVersion += 1;
    this.engineSubscription?.();
    this.engineSubscription = null;
    this.projectionEnginePromise = null;
    this.persistQueued = false;
    this.viewportPatchPublisher.reset();
    this.runtimeStateCache = null;
    this.snapshotCaches.reset();
    this.authoritativeStateSource = "none";
    this.mutationJournalEntries = [];
    this.pendingMutations = [];
    this.nextPendingMutationSeq = 1;
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

  private syncPendingMutationsFromJournal(): void {
    this.pendingMutations = syncPendingMutationsFromJournal(this.mutationJournalEntries);
  }

  private replaceJournalMutation(nextMutation: PendingWorkbookMutation): void {
    const nextState = replaceJournalMutation(this.mutationJournalEntries, nextMutation);
    this.mutationJournalEntries = nextState.mutationJournalEntries;
    this.pendingMutations = nextState.pendingMutations;
  }

  private async markRemainingJournalMutationsRebased(rebasedAtUnixMs = Date.now()): Promise<void> {
    const nextState = markJournalMutationsRebased(this.mutationJournalEntries, rebasedAtUnixMs);
    this.mutationJournalEntries = nextState.mutationJournalEntries;
    this.pendingMutations = nextState.pendingMutations;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await Promise.all(
        nextState.updatedMutations.map((mutation) =>
          this.localStore!.updatePendingMutation(mutation),
        ),
      );
    }
  }

  private storeRuntimeState(state: WorkbookWorkerStateSnapshot): WorkbookWorkerStateSnapshot {
    this.runtimeStateCache = cloneWorkerRuntimeState({
      ...state,
      pendingMutationSummary: this.buildPendingMutationSummary(),
    });
    const cachedState = this.runtimeStateCache;
    return withExternalSyncState(cachedState, this.externalSyncState);
  }

  private updateRuntimeStateFromEngine(
    engine: SpreadsheetEngine & WorkerEngine = this.requireEngine(),
  ): WorkbookWorkerStateSnapshot {
    return this.storeRuntimeState(buildWorkerRuntimeStateFromEngine(engine));
  }

  private getCurrentMetrics(): RecalcMetrics {
    if (this.engine) {
      return cloneRuntimeMetrics(this.engine.getLastMetrics());
    }
    return cloneRuntimeMetrics(this.runtimeStateCache?.metrics ?? EMPTY_RUNTIME_METRICS);
  }

  private buildPendingMutationSummary(): WorkbookPendingMutationSummarySnapshot {
    return buildPendingMutationSummary(this.mutationJournalEntries, this.pendingMutations);
  }

  private async getAuthoritativeStateInput(): Promise<{
    snapshot: WorkbookSnapshot | null;
    replica: EngineReplicaSnapshot | null;
  }> {
    return await resolveAuthoritativeStateInput({
      authoritativeStateSource: this.authoritativeStateSource,
      localStore: this.localStore,
      snapshotCaches: this.snapshotCaches,
      authoritativeEngine: this.authoritativeEngine,
      installRestoredAuthoritativeState: (snapshot, replica) => {
        this.installRestoredAuthoritativeState(snapshot, replica);
      },
    });
  }

  private readProjectedCellFromLocalStore(sheetName: string, address: string): CellSnapshot | null {
    return readProjectedCellFromLocalStore({
      canReadLocalProjectionForViewport: this.canReadLocalProjectionForViewport(),
      localStore: this.localStore,
      viewportTileStore: this.viewportTileStore,
      sheetName,
      address,
    });
  }

  private installAuthoritativeEngine(
    engine: SpreadsheetEngine,
    snapshot: WorkbookSnapshot | null,
    replica: EngineReplicaSnapshot | null,
  ): void {
    this.authoritativeStateSource = "memory";
    this.authoritativeEngine = engine;
    installAuthoritativeEngineState(this.snapshotCaches, engine, snapshot, replica);
  }

  private installRestoredAuthoritativeState(
    snapshot: WorkbookSnapshot | null,
    replica: EngineReplicaSnapshot | null,
  ): void {
    this.authoritativeStateSource = "memory";
    this.authoritativeEngine = null;
    installRestoredAuthoritativeState(this.snapshotCaches, snapshot, replica);
  }

  private async getAuthoritativeEngine(): Promise<SpreadsheetEngine> {
    const options = this.requireBootstrapOptions();
    const engine = await ensureAuthoritativeEngine({
      authoritativeEngine: this.authoritativeEngine,
      documentId: options.documentId,
      replicaId: options.replicaId,
      snapshotCaches: this.snapshotCaches,
      resolveAuthoritativeStateInput: () => this.getAuthoritativeStateInput(),
    });
    this.authoritativeEngine = engine;
    return engine;
  }

  private async rebuildProjectionEngine(): Promise<{
    engine: SpreadsheetEngine;
    overlayScope: ProjectionOverlayScope | null;
  }> {
    const options = this.requireBootstrapOptions();
    return await rebuildProjectionEngine({
      documentId: options.documentId,
      replicaId: options.replicaId,
      pendingMutations: this.pendingMutations,
      resolveAuthoritativeStateInput: () => this.getAuthoritativeStateInput(),
    });
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
      this.snapshotCaches.invalidateProjectionSnapshot();
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
    const persisted = buildPersistedWorkerState({
      snapshotCaches: this.snapshotCaches,
      authoritativeEngine,
      projectionEngine,
      hasDedicatedAuthoritativeEngine: this.authoritativeEngine !== null,
      authoritativeRevision: this.authoritativeRevision,
      appliedPendingLocalSeq: this.pendingMutations.at(-1)?.localSeq ?? 0,
    });
    this.persistQueued = false;
    const localStore = this.localStore;
    if (!localStore) {
      return;
    }
    const savePromise = persistProjectionStateToLocalStore({
      localStore,
      state: persisted,
      authoritativeEngine,
      projectionEngine,
      projectionOverlayScope: this.getProjectionOverlayScopeForPersist(),
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
    this.viewportPatchPublisher.broadcast({
      event,
      metrics,
      impactsBySheet:
        event === null ? null : collectSheetViewportImpacts(this.requireEngine(), event),
    });
  }

  private buildViewportPatch(
    state: ViewportSubscriptionState,
    event: EngineEvent | null,
    metrics: RecalcMetrics = this.getCurrentMetrics(),
    sheetImpact: SheetViewportImpact | null = null,
  ): ViewportPatch {
    return this.viewportPatchPublisher.buildPatch(state, event, metrics, sheetImpact);
  }
}
