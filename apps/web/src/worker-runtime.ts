import type { EngineSyncClient } from "@bilig/core";
import { SpreadsheetEngine } from "@bilig/core";
import type { EngineOpBatch } from "@bilig/crdt";
import { formatAddress, indexToColumn } from "@bilig/formula";
import { createBrowserPersistence, type BrowserPersistence } from "@bilig/storage-browser";
import {
  MAX_COLS,
  MAX_ROWS,
  ValueTag,
  formatErrorCode,
  type CellSnapshot,
  type EngineEvent,
  type RecalcMetrics,
  type SyncState,
  type WorkbookAxisEntrySnapshot,
} from "@bilig/protocol";
import {
  createWebSocketSyncClient,
  encodeViewportPatch,
  type ViewportAxisPatch,
  type ViewportPatch,
  type ViewportPatchedCell,
  type ViewportPatchSubscription,
  type WebSocketSyncClientOptions,
} from "@bilig/worker-transport";

type WorkbookEngine = InstanceType<typeof SpreadsheetEngine>;

const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_ROW_HEIGHT = 22;
const MIN_COLUMN_WIDTH = 44;
const MAX_COLUMN_WIDTH = 480;
const AUTOFIT_PADDING = 28;
const AUTOFIT_CHAR_WIDTH = 8;

export interface WorkbookWorkerBootstrapOptions {
  documentId: string;
  replicaId: string;
  baseUrl: string | null;
  persistState: boolean;
}

export interface WorkbookWorkerStateSnapshot {
  workbookName: string;
  sheetNames: string[];
  metrics: RecalcMetrics;
  syncState: SyncState;
}

interface PersistedWorkbookState {
  snapshot: ReturnType<WorkbookEngine["exportSnapshot"]>;
  replica: ReturnType<WorkbookEngine["exportReplicaSnapshot"]>;
}

interface ViewportSubscriptionState {
  subscription: ViewportPatchSubscription;
  listener: (patch: Uint8Array) => void;
  nextVersion: number;
  lastCellSignatures: Map<string, string>;
  lastColumnSignatures: Map<number, string>;
  lastRowSignatures: Map<number, string>;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isPersistedWorkbookState(value: unknown): value is PersistedWorkbookState {
  if (!isRecord(value) || !isRecord(value["snapshot"]) || !isRecord(value["replica"])) {
    return false;
  }
  const snapshot = value["snapshot"];
  const replica = value["replica"];
  return (
    Array.isArray(snapshot["sheets"]) &&
    isRecord(replica["replica"]) &&
    Array.isArray(replica["entityVersions"])
  );
}

function parsePersistedWorkbookState(value: unknown): PersistedWorkbookState | null {
  return isPersistedWorkbookState(value) ? value : null;
}

export class WorkbookWorkerRuntime {
  [method: string]: unknown;
  private readonly persistence: BrowserPersistence;
  private readonly createSyncClient: (options: WebSocketSyncClientOptions) => EngineSyncClient;
  private engine: WorkbookEngine | null = null;
  private bootstrapOptions: WorkbookWorkerBootstrapOptions | null = null;
  private engineSubscription: (() => void) | null = null;
  private readonly viewportSubscriptions = new Set<ViewportSubscriptionState>();
  private readonly formatIds = new Map<string, number>([["", 0]]);
  private nextFormatId = 1;

  constructor(
    options: {
      persistence?: BrowserPersistence;
      createSyncClient?: (options: WebSocketSyncClientOptions) => EngineSyncClient;
    } = {},
  ) {
    this.persistence = options.persistence ?? createBrowserPersistence();
    this.createSyncClient = options.createSyncClient ?? createWebSocketSyncClient;
  }

  async ready(): Promise<void> {
    await this.engine?.ready();
  }

  async bootstrap(options: WorkbookWorkerBootstrapOptions): Promise<WorkbookWorkerStateSnapshot> {
    this.cleanup();
    this.bootstrapOptions = options;
    this.engine = new SpreadsheetEngine({
      workbookName: options.documentId,
      replicaId: options.replicaId,
    });
    await this.engine.ready();

    if (options.persistState) {
      const persisted = await this.persistence.loadJson(
        this.persistenceKey(options.documentId),
        parsePersistedWorkbookState,
      );
      if (persisted) {
        this.engine.importSnapshot(persisted.snapshot);
        this.engine.importReplicaSnapshot(persisted.replica);
      }
    }
    if (this.engine.workbook.sheetsByName.size === 0) {
      this.engine.createSheet("Sheet1");
    }

    this.engineSubscription = this.engine.subscribe((event) => {
      void this.persistState();
      this.broadcastViewportPatches(event.metrics);
    });
    await this.persistState();

    if (options.baseUrl) {
      try {
        await this.engine.connectSyncClient(
          this.createSyncClient({
            documentId: options.documentId,
            replicaId: options.replicaId,
            baseUrl: options.baseUrl,
          }),
        );
      } catch {
        await this.engine.disconnectSyncClient();
      }
    }

    return this.getRuntimeState();
  }

  getRuntimeState(): WorkbookWorkerStateSnapshot {
    const engine = this.requireEngine();
    return {
      workbookName: engine.workbook.workbookName,
      sheetNames: this.listSheetNames(),
      metrics: { ...engine.getLastMetrics() },
      syncState: engine.getSyncState(),
    };
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

  clearCell(sheetName: string, address: string): CellSnapshot {
    this.requireEngine().clearCell(sheetName, address);
    return this.getCell(sheetName, address);
  }

  renderCommit(ops: Parameters<WorkbookEngine["renderCommit"]>[0]): void {
    this.requireEngine().renderCommit(ops);
  }

  fillRange(
    source: Parameters<WorkbookEngine["fillRange"]>[0],
    target: Parameters<WorkbookEngine["fillRange"]>[1],
  ): void {
    this.requireEngine().fillRange(source, target);
  }

  copyRange(
    source: Parameters<WorkbookEngine["copyRange"]>[0],
    target: Parameters<WorkbookEngine["copyRange"]>[1],
  ): void {
    this.requireEngine().copyRange(source, target);
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
      subscription: this.normalizeViewport(subscription),
      listener,
      nextVersion: 1,
      lastCellSignatures: new Map<string, string>(),
      lastColumnSignatures: new Map<number, string>(),
      lastRowSignatures: new Map<number, string>(),
    };
    listener(encodeViewportPatch(this.buildViewportPatch(state, true)));
    this.viewportSubscriptions.add(state);
    return () => {
      this.viewportSubscriptions.delete(state);
    };
  }

  private cleanup(): void {
    this.engineSubscription?.();
    this.engineSubscription = null;
    this.viewportSubscriptions.clear();
    this.engine = null;
  }

  private requireEngine(): WorkbookEngine {
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

  private async persistState(): Promise<void> {
    if (!this.bootstrapOptions) {
      return;
    }
    if (!this.bootstrapOptions.persistState) {
      return;
    }
    const engine = this.requireEngine();
    const persisted: PersistedWorkbookState = {
      snapshot: engine.exportSnapshot(),
      replica: engine.exportReplicaSnapshot(),
    };
    await this.persistence.saveJson(
      this.persistenceKey(this.bootstrapOptions.documentId),
      persisted,
    );
  }

  private broadcastViewportPatches(metrics: RecalcMetrics): void {
    for (const subscription of this.viewportSubscriptions) {
      const patch = this.buildViewportPatch(subscription, false, metrics);
      if (patch.cells.length === 0 && patch.columns.length === 0 && patch.rows.length === 0) {
        continue;
      }
      subscription.listener(encodeViewportPatch(patch));
    }
  }

  private buildViewportPatch(
    state: ViewportSubscriptionState,
    full: boolean,
    metrics: RecalcMetrics = this.requireEngine().getLastMetrics(),
  ): ViewportPatch {
    const engine = this.requireEngine();
    const viewport = state.subscription;
    const nextCellSignatures = new Map<string, string>();
    const cells: ViewportPatchedCell[] = [];

    for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
      for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
        const address = formatAddress(row, col);
        const key = `${viewport.sheetName}!${address}`;
        const snapshot = engine.workbook.getSheet(viewport.sheetName)
          ? engine.getCell(viewport.sheetName, address)
          : this.emptyCellSnapshot(viewport.sheetName, address);
        const formatId = this.getFormatId(snapshot.format);
        const patchedCell = this.buildPatchedCell(snapshot, row, col, formatId);
        const signature = JSON.stringify([
          patchedCell.snapshot.version,
          patchedCell.snapshot.formula ?? "",
          patchedCell.snapshot.input ?? null,
          patchedCell.snapshot.format ?? "",
          patchedCell.snapshot.value,
          patchedCell.displayText,
          patchedCell.copyText,
          patchedCell.editorText,
          patchedCell.formatId,
          patchedCell.styleId,
        ]);
        nextCellSignatures.set(key, signature);
        if (full || state.lastCellSignatures.get(key) !== signature) {
          cells.push(patchedCell);
        }
      }
    }
    state.lastCellSignatures = nextCellSignatures;

    const columnEntries = this.indexAxisEntries(engine.getColumnAxisEntries(viewport.sheetName));
    const rowEntries = this.indexAxisEntries(engine.getRowAxisEntries(viewport.sheetName));
    const { patches: columns, signatures: columnSignatures } = this.buildAxisPatches(
      viewport.colStart,
      viewport.colEnd,
      columnEntries,
      PRODUCT_COLUMN_WIDTH,
      state.lastColumnSignatures,
      full,
    );
    const { patches: rows, signatures: rowSignatures } = this.buildAxisPatches(
      viewport.rowStart,
      viewport.rowEnd,
      rowEntries,
      PRODUCT_ROW_HEIGHT,
      state.lastRowSignatures,
      full,
    );
    state.lastColumnSignatures = columnSignatures;
    state.lastRowSignatures = rowSignatures;

    return {
      version: state.nextVersion++,
      full,
      viewport,
      metrics: { ...metrics },
      cells,
      columns,
      rows,
    };
  }

  private buildPatchedCell(
    snapshot: CellSnapshot,
    row: number,
    col: number,
    formatId: number,
  ): ViewportPatchedCell {
    const editorText = this.toEditorText(snapshot);
    const displayText = this.toDisplayText(snapshot);
    return {
      row,
      col,
      snapshot,
      displayText,
      copyText: snapshot.formula ? editorText : displayText,
      editorText,
      formatId,
      styleId: formatId,
    };
  }

  private buildAxisPatches(
    start: number,
    end: number,
    entries: Map<number, WorkbookAxisEntrySnapshot>,
    defaultSize: number,
    previous: Map<number, string>,
    full: boolean,
  ): { patches: ViewportAxisPatch[]; signatures: Map<number, string> } {
    const signatures = new Map<number, string>();
    const patches: ViewportAxisPatch[] = [];
    for (let index = start; index <= end; index += 1) {
      const entry = entries.get(index);
      const size = entry?.size ?? defaultSize;
      const hidden = entry?.hidden ?? false;
      const signature = `${size}:${hidden ? 1 : 0}`;
      signatures.set(index, signature);
      if (full || previous.get(index) !== signature) {
        patches.push({ index, size, hidden });
      }
    }
    return { patches, signatures };
  }

  private indexAxisEntries(
    entries: readonly WorkbookAxisEntrySnapshot[],
  ): Map<number, WorkbookAxisEntrySnapshot> {
    return new Map(entries.map((entry) => [entry.index, entry]));
  }

  private normalizeViewport(subscription: ViewportPatchSubscription): ViewportPatchSubscription {
    const rowStart = Math.max(0, Math.min(MAX_ROWS - 1, subscription.rowStart));
    const rowEnd = Math.max(rowStart, Math.min(MAX_ROWS - 1, subscription.rowEnd));
    const colStart = Math.max(0, Math.min(MAX_COLS - 1, subscription.colStart));
    const colEnd = Math.max(colStart, Math.min(MAX_COLS - 1, subscription.colEnd));
    return {
      sheetName: subscription.sheetName,
      rowStart,
      rowEnd,
      colStart,
      colEnd,
    };
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
    switch (snapshot.value.tag) {
      case ValueTag.Number:
        return String(snapshot.value.value);
      case ValueTag.Boolean:
        return snapshot.value.value ? "TRUE" : "FALSE";
      case ValueTag.String:
        return snapshot.value.value;
      case ValueTag.Error:
        return formatErrorCode(snapshot.value.code);
      case ValueTag.Empty:
        return "";
    }
    const exhaustiveValue: never = snapshot.value;
    return String(exhaustiveValue);
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
}
