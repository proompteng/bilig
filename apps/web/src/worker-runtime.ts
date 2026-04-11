import type { CommitOp, EngineReplicaSnapshot } from "@bilig/core";
import { isEngineReplicaSnapshot, SpreadsheetEngine } from "@bilig/core";
import {
  buildWorkbookAgentPreview,
  isWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentPreviewSummary,
} from "@bilig/agent-api";
import type { EngineOpBatch } from "@bilig/workbook-domain";
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
import { clonePendingWorkbookMutation } from "./workbook-mutation-journal.js";
import {
  ackAbsorbedMutations,
  ackMutationInJournal,
  createRuntimePendingMutation,
  failMutationInJournal,
  markMutationSubmittedInJournal,
  recordMutationAttemptInJournal,
  retryMutationInJournal,
} from "./worker-runtime-mutation-actions.js";
import {
  buildPendingMutationSummary,
  markJournalMutationsRebased,
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
  resolveProjectionOverlayScopeForPersist,
  WorkerRuntimePersistCoordinator,
} from "./worker-runtime-persist-coordinator.js";
import {
  acquireProjectionEngine,
  scheduleProjectionEngineMaterialization,
} from "./worker-runtime-projection-engine.js";
import { WorkerRuntimeProjectionCommands } from "./worker-runtime-projection-commands.js";
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
  private readonly persistCoordinator = new WorkerRuntimePersistCoordinator({
    canPersistState: () => this.canPersistState(),
    getLocalStore: () => this.localStore,
    getAuthoritativeEngine: () => this.getAuthoritativeEngine(),
    getProjectionEngine: () => this.getProjectionEngine(),
    buildPersistedState: ({ authoritativeEngine, projectionEngine }) =>
      buildPersistedWorkerState({
        snapshotCaches: this.snapshotCaches,
        authoritativeEngine,
        projectionEngine,
        hasDedicatedAuthoritativeEngine: this.authoritativeEngine !== null,
        authoritativeRevision: this.authoritativeRevision,
        appliedPendingLocalSeq: this.pendingMutations.at(-1)?.localSeq ?? 0,
      }),
    getProjectionOverlayScope: () =>
      resolveProjectionOverlayScopeForPersist({
        projectionOverlayScope: this.projectionOverlayScope,
        pendingMutationCount: this.pendingMutations.length,
      }),
    saveState: (input) => persistProjectionStateToLocalStore(input),
    markProjectionMatchesLocalStore: () => {
      this.projectionMatchesLocalStore = true;
    },
  });
  private readonly projectionCommands = new WorkerRuntimeProjectionCommands({
    markProjectionDivergedFromLocalStore: () => this.markProjectionDivergedFromLocalStore(),
    getProjectionEngine: () => this.getProjectionEngine(),
    getCell: (sheetName, address) => this.getCell(sheetName, address),
    minColumnWidth: MIN_COLUMN_WIDTH,
    maxColumnWidth: MAX_COLUMN_WIDTH,
    autofitCharWidth: AUTOFIT_CHAR_WIDTH,
    autofitPadding: AUTOFIT_PADDING,
    formatCellDisplayValue,
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
        await this.persistCoordinator.queuePersist();
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
    await this.persistCoordinator.queuePersist();
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
      const result = ackAbsorbedMutations({
        mutationJournalEntries: this.mutationJournalEntries,
        absorbedMutationIds,
        ackedAtUnixMs: Date.now(),
      });
      this.mutationJournalEntries = result.mutationJournalEntries;
      this.pendingMutations = result.pendingMutations;
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
        projectionOverlayScope: resolveProjectionOverlayScopeForPersist({
          projectionOverlayScope: this.projectionOverlayScope,
          pendingMutationCount: this.pendingMutations.length,
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
    const nextMutation = createRuntimePendingMutation({
      documentId,
      localSeq,
      authoritativeRevision: this.authoritativeRevision,
      input,
      enqueuedAtUnixMs: Date.now(),
    });
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
    await this.persistCoordinator.queuePersist();
    return {
      ...nextMutation,
      args: [...nextMutation.args],
    };
  }

  async markPendingMutationSubmitted(id: string): Promise<void> {
    const result = markMutationSubmittedInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
      submittedAtUnixMs: Date.now(),
    });
    if (!result) {
      return;
    }
    this.mutationJournalEntries = result.mutationJournalEntries;
    this.pendingMutations = result.pendingMutations;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(result.updatedMutation);
    }
  }

  async ackPendingMutation(id: string): Promise<void> {
    const result = ackMutationInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
      ackedAtUnixMs: Date.now(),
    });
    if (!result) {
      return;
    }
    this.mutationJournalEntries = result.mutationJournalEntries;
    this.pendingMutations = result.pendingMutations;
    this.projectionMatchesLocalStore = false;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(result.updatedMutation);
      await this.persistCoordinator.queuePersist();
    }
  }

  async recordPendingMutationAttempt(id: string): Promise<void> {
    const result = recordMutationAttemptInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
      attemptedAtUnixMs: Date.now(),
    });
    if (!result) {
      return;
    }
    this.mutationJournalEntries = result.mutationJournalEntries;
    this.pendingMutations = result.pendingMutations;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(result.updatedMutation);
    }
  }

  async markPendingMutationFailed(id: string, failureMessage: string): Promise<void> {
    const result = failMutationInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
      failureMessage,
      failedAtUnixMs: Date.now(),
    });
    if (!result) {
      return;
    }
    this.mutationJournalEntries = result.mutationJournalEntries;
    this.pendingMutations = result.pendingMutations;
    this.projectionMatchesLocalStore = false;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(result.updatedMutation);
      await this.persistCoordinator.queuePersist();
    }
  }

  async retryPendingMutation(id: string): Promise<void> {
    const result = retryMutationInJournal({
      mutationJournalEntries: this.mutationJournalEntries,
      id,
    });
    if (!result) {
      return;
    }
    this.mutationJournalEntries = result.mutationJournalEntries;
    this.pendingMutations = result.pendingMutations;
    this.projectionMatchesLocalStore = false;
    if (this.bootstrapOptions?.persistState && this.localStore) {
      await this.localStore.updatePendingMutation(result.updatedMutation);
      await this.persistCoordinator.queuePersist();
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
    return await this.projectionCommands.setCellValue(sheetName, address, value);
  }

  async setCellFormula(sheetName: string, address: string, formula: string): Promise<CellSnapshot> {
    return await this.projectionCommands.setCellFormula(sheetName, address, formula);
  }

  async setRangeStyle(range: CellRangeRef, patch: CellStylePatch): Promise<void> {
    await this.projectionCommands.setRangeStyle(range, patch);
  }

  async clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): Promise<void> {
    await this.projectionCommands.clearRangeStyle(range, fields);
  }

  async setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): Promise<void> {
    await this.projectionCommands.setRangeNumberFormat(range, format);
  }

  async clearRangeNumberFormat(range: CellRangeRef): Promise<void> {
    await this.projectionCommands.clearRangeNumberFormat(range);
  }

  async clearRange(range: CellRangeRef): Promise<void> {
    await this.projectionCommands.clearRange(range);
  }

  async clearCell(sheetName: string, address: string): Promise<CellSnapshot> {
    return await this.projectionCommands.clearCell(sheetName, address);
  }

  async renderCommit(ops: CommitOp[]): Promise<void> {
    await this.projectionCommands.renderCommit(ops);
  }

  async fillRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    await this.projectionCommands.fillRange(source, target);
  }

  async copyRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    await this.projectionCommands.copyRange(source, target);
  }

  async moveRange(source: CellRangeRef, target: CellRangeRef): Promise<void> {
    await this.projectionCommands.moveRange(source, target);
  }

  async updateRowMetadata(
    sheetName: string,
    startRow: number,
    count: number,
    height: number | null,
    hidden: boolean | null,
  ): Promise<void> {
    await this.projectionCommands.updateRowMetadata(sheetName, startRow, count, height, hidden);
  }

  async updateColumnMetadata(
    sheetName: string,
    startCol: number,
    count: number,
    width: number | null,
    hidden: boolean | null,
  ): Promise<number | null> {
    return await this.projectionCommands.updateColumnMetadata(
      sheetName,
      startCol,
      count,
      width,
      hidden,
    );
  }

  async updateColumnWidth(sheetName: string, columnIndex: number, width: number): Promise<number> {
    return await this.projectionCommands.updateColumnWidth(sheetName, columnIndex, width);
  }

  async setFreezePane(sheetName: string, rows: number, cols: number): Promise<void> {
    await this.projectionCommands.setFreezePane(sheetName, rows, cols);
  }

  async autofitColumn(sheetName: string, columnIndex: number): Promise<number> {
    return await this.projectionCommands.autofitColumn(sheetName, columnIndex);
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
    this.persistCoordinator.reset();
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
    return await acquireProjectionEngine({
      getInstalledEngine: () => this.engine,
      getProjectionEnginePromise: () => this.projectionEnginePromise,
      getProjectionBuildVersion: () => this.projectionBuildVersion,
      rebuildProjectionEngine: () => this.rebuildProjectionEngine(),
      setProjectionOverlayScope: (overlayScope) => {
        this.projectionOverlayScope = overlayScope;
      },
      installEngine: (engine) => {
        this.installEngine(engine);
      },
      setProjectionEnginePromise: (promise) => {
        this.projectionEnginePromise = promise;
      },
      requireInstalledEngine: () => this.requireEngine(),
    });
  }

  private scheduleProjectionEngineMaterialization(): void {
    scheduleProjectionEngineMaterialization({
      hasInstalledEngine: () => this.engine !== null,
      hasProjectionEnginePromise: () => this.projectionEnginePromise !== null,
      hasBootstrapOptions: () => this.bootstrapOptions !== null,
      getProjectionBuildVersion: () => this.projectionBuildVersion,
      getProjectionEngine: () => this.getProjectionEngine(),
      schedule: (callback) => {
        setTimeout(() => {
          callback();
        }, 0);
      },
    });
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
