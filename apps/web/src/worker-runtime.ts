import type { CommitOp, EngineReplicaSnapshot } from "@bilig/core";
import { SpreadsheetEngine } from "@bilig/core";
import type { EngineOpBatch } from "@bilig/workbook-domain";
import { formatAddress, indexToColumn, parseCellAddress } from "@bilig/formula";
import { createBrowserPersistence, type BrowserPersistence } from "@bilig/storage-browser";
import {
  type CellRangeRef,
  type CellNumberFormatInput,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
  MAX_COLS,
  MAX_ROWS,
  ValueTag,
  type CellSnapshot,
  type EngineEvent,
  type LiteralInput,
  type RecalcMetrics,
  type SyncState,
  type WorkbookAxisEntrySnapshot,
  type WorkbookSnapshot,
  formatCellDisplayValue,
} from "@bilig/protocol";
import {
  encodeViewportPatch,
  type ViewportAxisPatch,
  type ViewportPatch,
  type ViewportPatchedCell,
  type ViewportPatchSubscription,
} from "@bilig/worker-transport";

interface WorkerSheet {
  name: string;
  order: number;
  grid: {
    forEachCellEntry(listener: (cellIndex: number, row: number, col: number) => void): void;
  };
}

interface WorkerWorkbook {
  workbookName: string;
  sheetsByName: Map<string, WorkerSheet>;
  getSheet(sheetName: string): WorkerSheet | undefined;
  getQualifiedAddress(cellIndex: number): string;
}

interface WorkerEngine {
  workbook: WorkerWorkbook;
  ready(): Promise<void>;
  createSheet(name: string): void;
  subscribe(listener: (event: EngineEvent) => void): () => void;
  subscribeBatches(listener: (batch: EngineOpBatch) => void): () => void;
  getLastMetrics(): RecalcMetrics;
  getSyncState(): SyncState;
  getCell(sheetName: string, address: string): CellSnapshot;
  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined;
  setRangeNumberFormat(range: CellRangeRef, format: CellNumberFormatInput): void;
  clearRangeNumberFormat(range: CellRangeRef): void;
  clearRange(range: CellRangeRef): void;
  setCellValue(sheetName: string, address: string, value: LiteralInput): unknown;
  setCellFormula(sheetName: string, address: string, formula: string): unknown;
  setRangeStyle(range: CellRangeRef, patch: CellStylePatch): void;
  clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): void;
  clearCell(sheetName: string, address: string): void;
  renderCommit(ops: CommitOp[]): void;
  fillRange(source: CellRangeRef, target: CellRangeRef): void;
  copyRange(source: CellRangeRef, target: CellRangeRef): void;
  updateColumnMetadata(
    sheetName: string,
    start: number,
    count: number,
    size: number | null,
    hidden: boolean | null,
  ): unknown;
  exportSnapshot(): WorkbookSnapshot;
  exportReplicaSnapshot(): EngineReplicaSnapshot;
  importSnapshot(snapshot: WorkbookSnapshot): void;
  importReplicaSnapshot(snapshot: EngineReplicaSnapshot): void;
  getColumnAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[];
  getRowAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[];
}

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

interface PersistedWorkbookState {
  snapshot: WorkbookSnapshot;
  replica: EngineReplicaSnapshot;
}

interface ViewportSubscriptionState {
  subscription: ViewportPatchSubscription;
  listener: (patch: Uint8Array) => void;
  nextVersion: number;
  knownStyleIds: Set<string>;
  lastStyleSignatures: Map<string, string>;
  lastCellSignatures: Map<string, string>;
  lastColumnSignatures: Map<number, string>;
  lastRowSignatures: Map<number, string>;
}

interface ViewportCellPosition {
  address: string;
  row: number;
  col: number;
}

interface ChangedSheetCells {
  addresses: Set<string>;
  positions: ViewportCellPosition[];
}

interface NormalizedRangeImpact {
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

interface SheetViewportImpact {
  changedCells: ChangedSheetCells | null;
  invalidatedRanges: NormalizedRangeImpact[];
  invalidatedRows: { sheetName: string; startIndex: number; endIndex: number }[];
  invalidatedColumns: { sheetName: string; startIndex: number; endIndex: number }[];
}

function styleSignature(style: CellStyleRecord): string {
  const fill = style.fill?.backgroundColor ?? "";
  const font = style.font;
  const alignment = style.alignment;
  const borders = style.borders;
  return [
    fill,
    font?.family ?? "",
    font?.size ?? "",
    font?.bold ? 1 : 0,
    font?.italic ? 1 : 0,
    font?.underline ? 1 : 0,
    font?.color ?? "",
    alignment?.horizontal ?? "",
    alignment?.vertical ?? "",
    alignment?.wrap ? 1 : 0,
    alignment?.indent ?? "",
    borders?.top ? `${borders.top.style}:${borders.top.weight}:${borders.top.color}` : "",
    borders?.right ? `${borders.right.style}:${borders.right.weight}:${borders.right.color}` : "",
    borders?.bottom
      ? `${borders.bottom.style}:${borders.bottom.weight}:${borders.bottom.color}`
      : "",
    borders?.left ? `${borders.left.style}:${borders.left.weight}:${borders.left.color}` : "",
  ].join("|");
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

  async bootstrap(options: WorkbookWorkerBootstrapOptions): Promise<WorkbookWorkerStateSnapshot> {
    this.cleanup();
    this.bootstrapOptions = options;
    this.externalSyncState = null;
    this.snapshotCache = null;
    this.snapshotDirty = true;
    const engine = new SpreadsheetEngine({
      workbookName: options.documentId,
      replicaId: options.replicaId,
    });
    this.engine = engine;
    await engine.ready();

    if (options.persistState) {
      const persisted = await this.persistence.loadJson(
        this.persistenceKey(options.documentId),
        parsePersistedWorkbookState,
      );
      if (persisted) {
        engine.importSnapshot(persisted.snapshot);
        engine.importReplicaSnapshot(persisted.replica);
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

    return this.getRuntimeState();
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
    this.engine = null;
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

  private broadcastViewportPatches(
    event: EngineEvent | null,
    metrics: RecalcMetrics = this.requireEngine().getLastMetrics(),
  ): void {
    const impactsBySheet = event === null ? null : this.collectSheetViewportImpacts(event);
    const impactedSheets = impactsBySheet === null ? null : new Set(impactsBySheet.keys());
    for (const subscription of this.getViewportSubscriptionsForEvent(event, impactedSheets)) {
      const sheetImpact = impactsBySheet?.get(subscription.subscription.sheetName) ?? null;
      if (
        event !== null &&
        !this.viewportPatchMayBeImpacted(
          subscription.subscription,
          event,
          sheetImpact,
          impactedSheets,
        )
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
      const targetCells = this.collectViewportCells(
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

    const columnEntries = this.indexAxisEntries(engine.getColumnAxisEntries(viewport.sheetName));
    const rowEntries = this.indexAxisEntries(engine.getRowAxisEntries(viewport.sheetName));
    const { patches: columns, signatures: columnSignatures } = this.buildAxisPatches(
      viewport.colStart,
      viewport.colEnd,
      columnEntries,
      PRODUCT_COLUMN_WIDTH,
      state.lastColumnSignatures,
      full,
      invalidatedColumns,
    );
    const { patches: rows, signatures: rowSignatures } = this.buildAxisPatches(
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

  private buildAxisPatches(
    start: number,
    end: number,
    entries: Map<number, WorkbookAxisEntrySnapshot>,
    defaultSize: number,
    previous: Map<number, string>,
    full: boolean,
    invalidatedAxes: readonly { startIndex: number; endIndex: number }[] = [],
  ): { patches: ViewportAxisPatch[]; signatures: Map<number, string> } {
    if (!full && invalidatedAxes.length === 0) {
      return { patches: [], signatures: previous };
    }
    const signatures = full ? new Map<number, string>() : new Map(previous);
    const patches: ViewportAxisPatch[] = [];
    const indices = full
      ? this.collectAxisIndices(start, end, null)
      : this.collectAxisIndices(start, end, invalidatedAxes);
    for (const index of indices) {
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

  private collectViewportCells(
    viewport: ViewportPatchSubscription,
    changedCells: ChangedSheetCells | null,
    invalidatedRanges: readonly NormalizedRangeImpact[],
  ): ViewportCellPosition[] {
    const positions: ViewportCellPosition[] = [];
    const seen = new Set<string>();

    changedCells?.positions.forEach((cell) => {
      if (
        cell.row < viewport.rowStart ||
        cell.row > viewport.rowEnd ||
        cell.col < viewport.colStart ||
        cell.col > viewport.colEnd
      ) {
        return;
      }
      const key = `${cell.row}:${cell.col}`;
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      positions.push(cell);
    });

    for (let index = 0; index < invalidatedRanges.length; index += 1) {
      const range = invalidatedRanges[index]!;
      const rowStart = Math.max(viewport.rowStart, range.rowStart);
      const rowEnd = Math.min(viewport.rowEnd, range.rowEnd);
      const colStart = Math.max(viewport.colStart, range.colStart);
      const colEnd = Math.min(viewport.colEnd, range.colEnd);
      if (rowStart > rowEnd || colStart > colEnd) {
        continue;
      }
      for (let row = rowStart; row <= rowEnd; row += 1) {
        for (let col = colStart; col <= colEnd; col += 1) {
          const key = `${row}:${col}`;
          if (seen.has(key)) {
            continue;
          }
          seen.add(key);
          positions.push({ address: formatAddress(row, col), row, col });
        }
      }
    }
    return positions;
  }

  private collectAxisIndices(
    start: number,
    end: number,
    invalidatedAxes: readonly { startIndex: number; endIndex: number }[] | null,
  ): number[] {
    if (invalidatedAxes === null) {
      const indices: number[] = [];
      for (let index = start; index <= end; index += 1) {
        indices.push(index);
      }
      return indices;
    }

    const indices = new Set<number>();
    for (let axisIndex = 0; axisIndex < invalidatedAxes.length; axisIndex += 1) {
      const axis = invalidatedAxes[axisIndex]!;
      const clampedStart = Math.max(start, axis.startIndex);
      const clampedEnd = Math.min(end, axis.endIndex);
      if (clampedStart > clampedEnd) {
        continue;
      }
      for (let index = clampedStart; index <= clampedEnd; index += 1) {
        indices.add(index);
      }
    }
    return Array.from(indices).toSorted((left, right) => left - right);
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
    return formatCellDisplayValue(snapshot.value, snapshot.format);
  }

  private collectChangedCellsBySheet(
    changedCellIndices: readonly number[] | Uint32Array,
  ): Map<string, ChangedSheetCells> {
    const changedBySheet = new Map<string, ChangedSheetCells>();
    const workbook = this.requireEngine().workbook;

    for (let index = 0; index < changedCellIndices.length; index += 1) {
      const qualifiedAddress = workbook.getQualifiedAddress(changedCellIndices[index]!);
      const separator = qualifiedAddress.indexOf("!");
      if (separator <= 0) {
        continue;
      }
      const sheetName = qualifiedAddress.slice(0, separator);
      const address = qualifiedAddress.slice(separator + 1);
      const parsed = parseCellAddress(address, sheetName);
      const sheetCells = changedBySheet.get(sheetName) ?? {
        addresses: new Set<string>(),
        positions: [],
      };
      if (!sheetCells.addresses.has(address)) {
        sheetCells.addresses.add(address);
        sheetCells.positions.push({ address, row: parsed.row, col: parsed.col });
      }
      changedBySheet.set(sheetName, sheetCells);
    }

    return changedBySheet;
  }

  private collectSheetViewportImpacts(event: EngineEvent): Map<string, SheetViewportImpact> | null {
    const changedCellsBySheet =
      event.invalidation !== "full"
        ? this.collectChangedCellsBySheet(event.changedCellIndices)
        : null;
    const impactsBySheet = new Map<string, SheetViewportImpact>();

    changedCellsBySheet?.forEach((changedCells, sheetName) => {
      impactsBySheet.set(sheetName, {
        changedCells,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
      });
    });

    event.invalidatedRanges.forEach((range) => {
      const start = parseCellAddress(range.startAddress, range.sheetName);
      const end = parseCellAddress(range.endAddress, range.sheetName);
      const impact = impactsBySheet.get(range.sheetName) ?? {
        changedCells: null,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
      };
      impact.invalidatedRanges.push({
        rowStart: Math.min(start.row, end.row),
        rowEnd: Math.max(start.row, end.row),
        colStart: Math.min(start.col, end.col),
        colEnd: Math.max(start.col, end.col),
      });
      impactsBySheet.set(range.sheetName, impact);
    });

    event.invalidatedRows.forEach((entry) => {
      const impact = impactsBySheet.get(entry.sheetName) ?? {
        changedCells: null,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
      };
      impact.invalidatedRows.push(entry);
      impactsBySheet.set(entry.sheetName, impact);
    });

    event.invalidatedColumns.forEach((entry) => {
      const impact = impactsBySheet.get(entry.sheetName) ?? {
        changedCells: null,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
      };
      impact.invalidatedColumns.push(entry);
      impactsBySheet.set(entry.sheetName, impact);
    });

    return impactsBySheet.size > 0 ? impactsBySheet : null;
  }

  private viewportPatchMayBeImpacted(
    viewport: ViewportPatchSubscription,
    event: EngineEvent,
    sheetImpact: SheetViewportImpact | null,
    impactedSheets: ReadonlySet<string> | null,
  ): boolean {
    if (event.invalidation === "full") {
      return impactedSheets === null || impactedSheets.has(viewport.sheetName);
    }

    if (impactedSheets !== null && !impactedSheets.has(viewport.sheetName)) {
      return false;
    }

    const changedCells = sheetImpact?.changedCells;
    if (changedCells) {
      for (const parsed of changedCells.positions) {
        if (
          parsed.row >= viewport.rowStart &&
          parsed.row <= viewport.rowEnd &&
          parsed.col >= viewport.colStart &&
          parsed.col <= viewport.colEnd
        ) {
          return true;
        }
      }
    }

    for (let index = 0; index < (sheetImpact?.invalidatedRanges.length ?? 0); index += 1) {
      const range = sheetImpact!.invalidatedRanges[index]!;
      if (
        range.rowStart <= viewport.rowEnd &&
        range.rowEnd >= viewport.rowStart &&
        range.colStart <= viewport.colEnd &&
        range.colEnd >= viewport.colStart
      ) {
        return true;
      }
    }

    for (let index = 0; index < (sheetImpact?.invalidatedRows.length ?? 0); index += 1) {
      const rowInvalidation = sheetImpact!.invalidatedRows[index]!;
      if (
        rowInvalidation.startIndex <= viewport.rowEnd &&
        rowInvalidation.endIndex >= viewport.rowStart
      ) {
        return true;
      }
    }

    for (let index = 0; index < (sheetImpact?.invalidatedColumns.length ?? 0); index += 1) {
      const columnInvalidation = sheetImpact!.invalidatedColumns[index]!;
      if (
        columnInvalidation.startIndex <= viewport.colEnd &&
        columnInvalidation.endIndex >= viewport.colStart
      ) {
        return true;
      }
    }

    return false;
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
