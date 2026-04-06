import type { CommitOp, EngineReplicaSnapshot } from "@bilig/core";
import { SpreadsheetEngine } from "@bilig/core";
import type { EngineOpBatch } from "@bilig/workbook-domain";
import { formatAddress, indexToColumn } from "@bilig/formula";
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
import { buildWorkbookLocalProjectionOverlay } from "./worker-local-overlay.js";
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
  private authoritativeEngine: SpreadsheetEngine | null = null;
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
  private readonly viewportTileStore = new WorkerViewportTileStore();

  constructor(
    options: {
      localStoreFactory?: WorkbookLocalStoreFactory;
    } = {},
  ) {
    this.localStoreFactory = options.localStoreFactory ?? createOpfsWorkbookLocalStoreFactory();
  }

  async ready(): Promise<void> {
    await this.engine?.ready();
  }

  dispose(): void {
    this.cleanup();
  }

  async bootstrap(options: WorkbookWorkerBootstrapOptions): Promise<WorkbookWorkerBootstrapResult> {
    this.cleanup();
    this.bootstrapOptions = options;
    this.externalSyncState = null;
    this.snapshotCache = null;
    this.snapshotDirty = true;
    this.pendingMutations = [];
    this.nextPendingMutationSeq = 1;
    this.appliedPendingLocalSeq = 0;
    this.projectionMatchesLocalStore = false;
    let restoredFromPersistence = false;
    let requiresAuthoritativeHydrate = false;
    let restoredState: WorkbookStoredState | null = null;
    if (options.persistState) {
      try {
        this.localStore = await this.localStoreFactory.open(options.documentId);
        restoredState = await this.localStore.loadState();
        if (restoredState) {
          restoredFromPersistence = true;
          this.authoritativeRevision = restoredState.authoritativeRevision;
          this.appliedPendingLocalSeq = restoredState.appliedPendingLocalSeq;
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

    const restoredSnapshot = isWorkbookSnapshotValue(restoredState?.snapshot)
      ? restoredState.snapshot
      : null;
    const restoredReplica = isEngineReplicaSnapshotValue(restoredState?.replica)
      ? restoredState.replica
      : null;
    const highestPendingLocalSeq = this.pendingMutations.at(-1)?.localSeq ?? 0;
    if (restoredSnapshot === null && this.pendingMutations.length > 0) {
      requiresAuthoritativeHydrate = true;
    }

    if (restoredSnapshot || restoredReplica) {
      this.installRestoredAuthoritativeState(restoredSnapshot, restoredReplica);
    }
    const engine = await this.createEngineFromState({
      snapshot: restoredSnapshot,
      replica: restoredReplica,
      pendingMutationsToReplay: this.pendingMutations,
    });
    this.installEngine(engine);
    const projectionMatchesRestoredLocalStore =
      Boolean(this.localStore) &&
      restoredState !== null &&
      restoredState.appliedPendingLocalSeq === highestPendingLocalSeq;
    this.projectionMatchesLocalStore = projectionMatchesRestoredLocalStore;
    if (!requiresAuthoritativeHydrate && !projectionMatchesRestoredLocalStore) {
      await this.persistStateNow();
    }

    return {
      runtimeState: this.getRuntimeState(),
      restoredFromPersistence,
      requiresAuthoritativeHydrate,
    };
  }

  getRuntimeState(): WorkbookWorkerStateSnapshot {
    const engine = this.requireEngine();
    return {
      workbookName: engine.workbook.workbookName,
      sheetNames: this.listSheetNames(),
      metrics: { ...engine.getLastMetrics() },
      syncState: this.externalSyncState ?? engine.getSyncState(),
    };
  }

  getAuthoritativeRevision(): number {
    return this.authoritativeRevision;
  }

  async replaceSnapshot(
    snapshot: WorkbookSnapshot,
    authoritativeRevision = this.authoritativeRevision,
  ): Promise<WorkbookWorkerStateSnapshot> {
    const authoritativeEngine = await this.createEngineFromState({
      snapshot,
      replica: null,
      pendingMutationsToReplay: [],
    });
    this.authoritativeRevision = Math.max(this.authoritativeRevision, authoritativeRevision);
    this.installAuthoritativeEngine(authoritativeEngine, snapshot, null);
    const engine = await this.rebuildProjectionEngine();
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
    const authoritativeEngine = await this.createEngineFromState({
      snapshot,
      replica: null,
      pendingMutationsToReplay: [],
    });
    this.authoritativeRevision = authoritativeRevision;
    this.installAuthoritativeEngine(authoritativeEngine, snapshot, null);
    const engine = await this.rebuildProjectionEngine();
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
    const engine = await this.rebuildProjectionEngine();
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
    return this.getCachedSnapshot();
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
    applyPendingWorkbookMutationToEngine(this.requireEngine(), nextMutation);
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
    const engine = this.requireEngine();
    if (!engine.workbook.getSheet(sheetName)) {
      return this.emptyCellSnapshot(sheetName, address);
    }
    return engine.getCell(sheetName, address);
  }

  setCellValue(sheetName: string, address: string, value: CellSnapshot["input"]): CellSnapshot {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().setCellValue(sheetName, address, value ?? null);
    return this.getCell(sheetName, address);
  }

  setCellFormula(sheetName: string, address: string, formula: string): CellSnapshot {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().setCellFormula(sheetName, address, formula);
    return this.getCell(sheetName, address);
  }

  setRangeStyle(range: CellRangeRef, patch: CellStylePatch): void {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().setRangeStyle(range, patch);
  }

  clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): void {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().clearRangeStyle(range, fields);
  }

  setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): void {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().setRangeNumberFormat(range, format);
  }

  clearRangeNumberFormat(range: CellRangeRef): void {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().clearRangeNumberFormat(range);
  }

  clearRange(range: CellRangeRef): void {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().clearRange(range);
  }

  clearCell(sheetName: string, address: string): CellSnapshot {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().clearCell(sheetName, address);
    return this.getCell(sheetName, address);
  }

  renderCommit(ops: CommitOp[]): void {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().renderCommit(ops);
  }

  fillRange(source: CellRangeRef, target: CellRangeRef): void {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().fillRange(source, target);
  }

  copyRange(source: CellRangeRef, target: CellRangeRef): void {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().copyRange(source, target);
  }

  moveRange(source: CellRangeRef, target: CellRangeRef): void {
    this.markProjectionDivergedFromLocalStore();
    this.requireEngine().moveRange(source, target);
  }

  updateColumnWidth(sheetName: string, columnIndex: number, width: number): number {
    this.markProjectionDivergedFromLocalStore();
    const clamped = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(width)));
    this.requireEngine().updateColumnMetadata(sheetName, columnIndex, 1, clamped, null);
    return clamped;
  }

  autofitColumn(sheetName: string, columnIndex: number): number {
    const engine = this.requireEngine();
    const sheet = engine.workbook.getSheet(sheetName);
    let widest = indexToColumn(columnIndex).length * AUTOFIT_CHAR_WIDTH;

    sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
      if (col !== columnIndex) {
        return;
      }
      const display = this.toDisplayText(engine.getCell(sheetName, formatAddress(row, col)));
      widest = Math.max(widest, display.length * AUTOFIT_CHAR_WIDTH);
    });

    return this.updateColumnWidth(sheetName, columnIndex, widest + AUTOFIT_PADDING);
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
    listener(encodeViewportPatch(this.buildViewportPatch(state, null)));
    this.viewportSubscriptions.add(state);
    this.addViewportSubscription(state);
    return () => {
      this.viewportSubscriptions.delete(state);
      this.removeViewportSubscription(state);
    };
  }

  private cleanup(): void {
    this.engineSubscription?.();
    this.engineSubscription = null;
    this.persistQueued = false;
    this.viewportSubscriptions.clear();
    this.viewportSubscriptionsBySheet.clear();
    this.styles.clear();
    this.styles.set(DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID });
    this.snapshotCache = null;
    this.snapshotDirty = true;
    this.authoritativeSnapshotCache = null;
    this.authoritativeSnapshotDirty = true;
    this.authoritativeReplicaCache = null;
    this.authoritativeReplicaDirty = true;
    this.pendingMutations = [];
    this.nextPendingMutationSeq = 1;
    this.appliedPendingLocalSeq = 0;
    this.authoritativeRevision = 0;
    this.projectionMatchesLocalStore = false;
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
    this.projectionMatchesLocalStore = false;
    this.viewportTileStore.reset();
  }

  private listSheetNames(): string[] {
    return [...this.requireEngine().workbook.sheetsByName.values()]
      .toSorted((left, right) => left.order - right.order)
      .map((sheet) => sheet.name);
  }

  private async createEngineFromState(input: {
    snapshot: WorkbookSnapshot | null;
    replica: EngineReplicaSnapshot | null;
    pendingMutationsToReplay: readonly PendingWorkbookMutation[];
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
    for (const mutation of input.pendingMutationsToReplay) {
      applyPendingWorkbookMutationToEngine(engine, mutation);
    }
    return engine;
  }

  private installAuthoritativeEngine(
    engine: SpreadsheetEngine,
    snapshot: WorkbookSnapshot | null,
    replica: EngineReplicaSnapshot | null,
  ): void {
    this.authoritativeEngine = engine;
    this.storeCachedAuthoritativeSnapshot(snapshot ?? engine.exportSnapshot());
    this.storeCachedAuthoritativeReplica(replica ?? engine.exportReplicaSnapshot());
  }

  private installRestoredAuthoritativeState(
    snapshot: WorkbookSnapshot | null,
    replica: EngineReplicaSnapshot | null,
  ): void {
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
    const engine = await this.createEngineFromState({
      snapshot: this.authoritativeSnapshotCache,
      replica: this.authoritativeReplicaCache,
      pendingMutationsToReplay: [],
    });
    this.authoritativeEngine = engine;
    if (this.authoritativeSnapshotCache) {
      this.authoritativeSnapshotDirty = false;
    } else {
      this.storeCachedAuthoritativeSnapshot(engine.exportSnapshot());
    }
    if (this.authoritativeReplicaCache) {
      this.authoritativeReplicaDirty = false;
    } else {
      this.storeCachedAuthoritativeReplica(engine.exportReplicaSnapshot());
    }
    return engine;
  }

  private async rebuildProjectionEngine(): Promise<SpreadsheetEngine> {
    return await this.createEngineFromState({
      snapshot: this.getAuthoritativeSnapshot(),
      replica: this.getAuthoritativeReplica(),
      pendingMutationsToReplay: this.pendingMutations,
    });
  }

  private installEngine(engine: SpreadsheetEngine & WorkerEngine): void {
    this.engineSubscription?.();
    this.engine = engine;
    this.engineSubscription = engine.subscribe((event) => {
      this.invalidateSnapshotCache();
      this.broadcastViewportPatches(event);
    });
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
        projectionEngine: this.requireEngine(),
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
    metrics: RecalcMetrics = this.requireEngine().getLastMetrics(),
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
    metrics: RecalcMetrics = this.requireEngine().getLastMetrics(),
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
