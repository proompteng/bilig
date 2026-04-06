import type { CommitOp } from "@bilig/core";
import { SpreadsheetEngine } from "@bilig/core";
import type { EngineOpBatch } from "@bilig/workbook-domain";
import { formatAddress, indexToColumn } from "@bilig/formula";
import { createBrowserPersistence, type BrowserPersistence } from "@bilig/storage-browser";
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
  type ViewportPatchedCell,
  type ViewportPatchSubscription,
} from "@bilig/worker-transport";
import {
  isPendingWorkbookMutationInput,
  type PendingWorkbookMutation,
  type PendingWorkbookMutationInput,
} from "./workbook-sync.js";
import {
  buildAxisPatches,
  collectSheetViewportImpacts,
  collectViewportCells,
  indexAxisEntries,
  normalizeViewport,
  parsePersistedPendingMutationState,
  parsePersistedWorkbookState,
  styleSignature,
  viewportPatchMayBeImpacted,
  type PersistedPendingMutationState,
  type PersistedWorkbookState,
  type SheetViewportImpact,
  type ViewportSubscriptionState,
  type WorkerEngine,
} from "./worker-runtime-support.js";

const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_ROW_HEIGHT = 22;
const MIN_COLUMN_WIDTH = 44;
const MAX_COLUMN_WIDTH = 480;
const AUTOFIT_PADDING = 28;
const AUTOFIT_CHAR_WIDTH = 8;
const DEFAULT_STYLE_ID = "style-0";
const PERSIST_DEBOUNCE_MS = 120;

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
}

export class WorkbookWorkerRuntime {
  [method: string]: unknown;
  private readonly persistence: BrowserPersistence;
  private engine: WorkerEngine | null = null;
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
  private persistTimer: ReturnType<typeof setTimeout> | null = null;
  private persistInFlight: Promise<void> | null = null;
  private persistQueued = false;
  private pendingMutations: PendingWorkbookMutation[] = [];
  private nextPendingMutationSeq = 1;

  constructor(
    options: {
      persistence?: BrowserPersistence;
    } = {},
  ) {
    this.persistence = options.persistence ?? createBrowserPersistence();
  }

  async ready(): Promise<void> {
    await this.engine?.ready();
  }

  async bootstrap(options: WorkbookWorkerBootstrapOptions): Promise<WorkbookWorkerBootstrapResult> {
    this.cleanup();
    this.bootstrapOptions = options;
    this.externalSyncState = null;
    this.snapshotCache = null;
    this.snapshotDirty = true;
    this.pendingMutations = [];
    this.nextPendingMutationSeq = 1;
    const engine = new SpreadsheetEngine({
      workbookName: options.documentId,
      replicaId: options.replicaId,
    });
    this.engine = engine;
    await engine.ready();

    let restoredFromPersistence = false;
    if (options.persistState) {
      const persisted = await this.persistence.loadJson(
        this.persistenceKey(options.documentId),
        parsePersistedWorkbookState,
      );
      if (persisted) {
        engine.importSnapshot(persisted.snapshot);
        engine.importReplicaSnapshot(persisted.replica);
        restoredFromPersistence = true;
      }
      const persistedPendingMutations = await this.persistence.loadJson(
        this.pendingMutationsKey(options.documentId),
        parsePersistedPendingMutationState,
      );
      if (persistedPendingMutations) {
        this.pendingMutations = persistedPendingMutations.pendingMutations;
        this.nextPendingMutationSeq =
          persistedPendingMutations.pendingMutations.reduce((max, mutation) => {
            const suffix = Number(mutation.id.split(":").at(-1));
            return Number.isFinite(suffix) ? Math.max(max, suffix) : max;
          }, 0) + 1;
      }
    }
    if (engine.workbook.sheetsByName.size === 0) {
      engine.createSheet("Sheet1");
    }

    this.engineSubscription = engine.subscribe((event) => {
      this.invalidateSnapshotCache();
      this.schedulePersistState();
      this.broadcastViewportPatches(event);
    });
    await this.persistStateNow();

    return {
      runtimeState: this.getRuntimeState(),
      restoredFromPersistence,
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

  async replaceSnapshot(snapshot: WorkbookSnapshot): Promise<WorkbookWorkerStateSnapshot> {
    const engine = this.requireEngine();
    engine.importSnapshot(snapshot);
    this.storeCachedSnapshot(snapshot);
    await this.persistStateNow();
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
    const nextMutation: PendingWorkbookMutation = {
      id: `${documentId}:pending:${this.nextPendingMutationSeq++}`,
      method: input.method,
      args: [...input.args],
      enqueuedAtUnixMs: Date.now(),
    };
    this.pendingMutations.push(nextMutation);
    await this.persistStateNow();
    await this.persistPendingMutationsNow();
    return {
      ...nextMutation,
      args: [...nextMutation.args],
    };
  }

  async ackPendingMutation(id: string): Promise<void> {
    const nextMutations = this.pendingMutations.filter((mutation) => mutation.id !== id);
    if (nextMutations.length === this.pendingMutations.length) {
      return;
    }
    this.pendingMutations = nextMutations;
    await this.persistPendingMutationsNow();
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    const engine = this.requireEngine();
    if (!engine.workbook.getSheet(sheetName)) {
      return this.emptyCellSnapshot(sheetName, address);
    }
    return engine.getCell(sheetName, address);
  }

  setCellValue(sheetName: string, address: string, value: CellSnapshot["input"]): CellSnapshot {
    this.requireEngine().setCellValue(sheetName, address, value ?? null);
    return this.getCell(sheetName, address);
  }

  setCellFormula(sheetName: string, address: string, formula: string): CellSnapshot {
    this.requireEngine().setCellFormula(sheetName, address, formula);
    return this.getCell(sheetName, address);
  }

  setRangeStyle(range: CellRangeRef, patch: CellStylePatch): void {
    this.requireEngine().setRangeStyle(range, patch);
  }

  clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): void {
    this.requireEngine().clearRangeStyle(range, fields);
  }

  setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): void {
    this.requireEngine().setRangeNumberFormat(range, format);
  }

  clearRangeNumberFormat(range: CellRangeRef): void {
    this.requireEngine().clearRangeNumberFormat(range);
  }

  clearRange(range: CellRangeRef): void {
    this.requireEngine().clearRange(range);
  }

  clearCell(sheetName: string, address: string): CellSnapshot {
    this.requireEngine().clearCell(sheetName, address);
    return this.getCell(sheetName, address);
  }

  renderCommit(ops: CommitOp[]): void {
    this.requireEngine().renderCommit(ops);
  }

  fillRange(source: CellRangeRef, target: CellRangeRef): void {
    this.requireEngine().fillRange(source, target);
  }

  copyRange(source: CellRangeRef, target: CellRangeRef): void {
    this.requireEngine().copyRange(source, target);
  }

  moveRange(source: CellRangeRef, target: CellRangeRef): void {
    this.requireEngine().moveRange(source, target);
  }

  updateColumnWidth(sheetName: string, columnIndex: number, width: number): number {
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
    if (this.persistTimer !== null) {
      clearTimeout(this.persistTimer);
      this.persistTimer = null;
    }
    this.persistQueued = false;
    this.viewportSubscriptions.clear();
    this.viewportSubscriptionsBySheet.clear();
    this.styles.clear();
    this.styles.set(DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID });
    this.snapshotCache = null;
    this.snapshotDirty = true;
    this.pendingMutations = [];
    this.nextPendingMutationSeq = 1;
    this.engine = null;
  }

  private requireBootstrapOptions(): WorkbookWorkerBootstrapOptions {
    if (!this.bootstrapOptions) {
      throw new Error("Workbook worker runtime has not been bootstrapped");
    }
    return this.bootstrapOptions;
  }

  private requireEngine(): WorkerEngine {
    if (!this.engine) {
      throw new Error("Workbook worker runtime has not been bootstrapped");
    }
    return this.engine;
  }

  private listSheetNames(): string[] {
    return [...this.requireEngine().workbook.sheetsByName.values()]
      .toSorted((left, right) => left.order - right.order)
      .map((sheet) => sheet.name);
  }

  private persistenceKey(documentId: string): string {
    return `bilig:web:${documentId}:runtime`;
  }

  private pendingMutationsKey(documentId: string): string {
    return `bilig:web:${documentId}:pending-mutations`;
  }

  private schedulePersistState(): void {
    if (!this.bootstrapOptions?.persistState) {
      return;
    }
    this.persistQueued = true;
    if (this.persistTimer !== null) {
      return;
    }
    this.persistTimer = setTimeout(() => {
      this.persistTimer = null;
      void this.flushPersistState();
    }, PERSIST_DEBOUNCE_MS);
  }

  private async persistStateNow(): Promise<void> {
    if (this.persistTimer !== null) {
      clearTimeout(this.persistTimer);
      this.persistTimer = null;
    }
    this.persistQueued = true;
    await this.flushPersistState();
  }

  private async flushPersistState(): Promise<void> {
    if (!this.persistQueued || !this.bootstrapOptions?.persistState || !this.engine) {
      return;
    }
    if (this.persistInFlight) {
      await this.persistInFlight;
      if (this.persistQueued) {
        await this.flushPersistState();
      }
      return;
    }

    const documentId = this.bootstrapOptions.documentId;
    const persisted: PersistedWorkbookState = {
      snapshot: this.getCachedSnapshot(),
      replica: this.engine.exportReplicaSnapshot(),
    };
    this.persistQueued = false;
    const savePromise = this.persistence.saveJson(this.persistenceKey(documentId), persisted);
    this.persistInFlight = savePromise;
    try {
      await savePromise;
    } finally {
      if (this.persistInFlight === savePromise) {
        this.persistInFlight = null;
      }
    }
    if (this.persistQueued) {
      await this.flushPersistState();
    }
  }

  private async persistPendingMutationsNow(): Promise<void> {
    if (!this.bootstrapOptions?.persistState) {
      return;
    }
    const documentId = this.bootstrapOptions.documentId;
    if (this.pendingMutations.length === 0) {
      await this.persistence.remove(this.pendingMutationsKey(documentId));
      return;
    }
    await this.persistence.saveJson(this.pendingMutationsKey(documentId), {
      pendingMutations: this.pendingMutations,
    } satisfies PersistedPendingMutationState);
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
    const engine = this.requireEngine();
    const viewport = state.subscription;
    const hasSheet = engine.workbook.getSheet(viewport.sheetName) !== undefined;
    const styles: CellStyleRecord[] = [];
    const cells: ViewportPatchedCell[] = [];
    const full = event === null || event.invalidation === "full";
    const invalidatedRanges = sheetImpact?.invalidatedRanges ?? [];
    const invalidatedRows = sheetImpact?.invalidatedRows ?? [];
    const invalidatedColumns = sheetImpact?.invalidatedColumns ?? [];

    if (full) {
      state.lastCellSignatures.clear();
      state.lastStyleSignatures.clear();
      for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
        for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
          this.appendPatchedCell(
            state,
            styles,
            cells,
            viewport.sheetName,
            row,
            col,
            hasSheet,
            true,
          );
        }
      }
    } else {
      const targetCells = collectViewportCells(
        viewport,
        sheetImpact?.changedCells ?? null,
        invalidatedRanges,
      );
      for (const cell of targetCells) {
        this.appendPatchedCell(
          state,
          styles,
          cells,
          viewport.sheetName,
          cell.row,
          cell.col,
          hasSheet,
          false,
          cell.address,
        );
      }
    }

    const columnEntries = indexAxisEntries(engine.getColumnAxisEntries(viewport.sheetName));
    const rowEntries = indexAxisEntries(engine.getRowAxisEntries(viewport.sheetName));
    const { patches: columns, signatures: columnSignatures } = buildAxisPatches(
      viewport.colStart,
      viewport.colEnd,
      columnEntries,
      PRODUCT_COLUMN_WIDTH,
      state.lastColumnSignatures,
      full,
      invalidatedColumns,
    );
    const { patches: rows, signatures: rowSignatures } = buildAxisPatches(
      viewport.rowStart,
      viewport.rowEnd,
      rowEntries,
      PRODUCT_ROW_HEIGHT,
      state.lastRowSignatures,
      full,
      invalidatedRows,
    );
    state.lastColumnSignatures = columnSignatures;
    state.lastRowSignatures = rowSignatures;

    return {
      version: state.nextVersion++,
      full,
      viewport,
      metrics: { ...metrics },
      styles,
      cells,
      columns,
      rows,
    };
  }

  private appendPatchedCell(
    state: ViewportSubscriptionState,
    styles: CellStyleRecord[],
    cells: ViewportPatchedCell[],
    sheetName: string,
    row: number,
    col: number,
    hasSheet: boolean,
    force: boolean,
    address = formatAddress(row, col),
  ): void {
    const key = `${sheetName}!${address}`;
    const snapshot = hasSheet
      ? this.requireEngine().getCell(sheetName, address)
      : this.emptyCellSnapshot(sheetName, address);
    const formatId = this.getFormatId(snapshot.format);
    const style = this.getStyleRecord(snapshot.styleId ?? DEFAULT_STYLE_ID);
    const nextStyleSignature = styleSignature(style);
    const previousStyleSignature = state.lastStyleSignatures.get(style.id);
    if (
      force ||
      previousStyleSignature !== nextStyleSignature ||
      !state.knownStyleIds.has(style.id)
    ) {
      state.knownStyleIds.add(style.id);
      state.lastStyleSignatures.set(style.id, nextStyleSignature);
      styles.push(style);
    }
    const editorText = this.toEditorText(snapshot);
    const displayText = this.toDisplayText(snapshot);
    const copyText = snapshot.formula ? editorText : displayText;
    const signature = this.buildPatchedCellSignature(
      snapshot,
      displayText,
      copyText,
      editorText,
      formatId,
      style.id,
    );
    if (force || state.lastCellSignatures.get(key) !== signature) {
      cells.push({
        row,
        col,
        snapshot,
        displayText,
        copyText,
        editorText,
        formatId,
        styleId: style.id,
      });
    }
    state.lastCellSignatures.set(key, signature);
  }

  private buildPatchedCellSignature(
    snapshot: CellSnapshot,
    displayText: string,
    copyText: string,
    editorText: string,
    formatId: number,
    styleId: string,
  ): string {
    return [
      snapshot.version,
      snapshot.flags,
      snapshot.formula ?? "",
      snapshot.input ?? "",
      snapshot.format ?? "",
      snapshot.styleId ?? "",
      formatId,
      styleId,
      this.snapshotValueSignature(snapshot),
      displayText,
      copyText,
      editorText,
    ].join("|");
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

  private toEditorText(snapshot: CellSnapshot): string {
    if (snapshot.formula) {
      return `=${snapshot.formula}`;
    }
    if (snapshot.input === null || snapshot.input === undefined) {
      return this.toDisplayText(snapshot);
    }
    if (typeof snapshot.input === "boolean") {
      return snapshot.input ? "TRUE" : "FALSE";
    }
    return String(snapshot.input);
  }

  private toDisplayText(snapshot: CellSnapshot): string {
    return formatCellDisplayValue(snapshot.value, snapshot.format);
  }

  private snapshotValueSignature(snapshot: CellSnapshot): string {
    switch (snapshot.value.tag) {
      case ValueTag.Number:
        return `n:${snapshot.value.value}`;
      case ValueTag.Boolean:
        return `b:${snapshot.value.value ? 1 : 0}`;
      case ValueTag.String:
        return `s:${snapshot.value.stringId}:${snapshot.value.value}`;
      case ValueTag.Error:
        return `e:${snapshot.value.code}`;
      case ValueTag.Empty:
        return "empty";
    }
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
}
