import type { CommitOp, EngineReplicaSnapshot, EngineSyncClient } from "@bilig/core";
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
  createWebSocketSyncClient,
  encodeViewportPatch,
  type ViewportAxisPatch,
  type ViewportPatch,
  type ViewportPatchedCell,
  type ViewportPatchSubscription,
  type WebSocketSyncClientOptions,
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
  connectSyncClient(client: EngineSyncClient): Promise<void>;
  disconnectSyncClient(): Promise<void>;
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

export interface WorkbookWorkerBootstrapOptions {
  documentId: string;
  replicaId: string;
  baseUrl: string | null;
  persistState: boolean;
}

interface ServerSnapshotSeed {
  cursor: number;
  snapshot: WorkbookSnapshot;
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

function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  return (
    isRecord(value) &&
    value["version"] === 1 &&
    isRecord(value["workbook"]) &&
    typeof value["workbook"]["name"] === "string" &&
    Array.isArray(value["sheets"]) &&
    value["sheets"].every((sheet) => {
      return (
        isRecord(sheet) &&
        typeof sheet["name"] === "string" &&
        typeof sheet["order"] === "number" &&
        Array.isArray(sheet["cells"]) &&
        sheet["cells"].every((cell) => {
          return (
            isRecord(cell) &&
            typeof cell["address"] === "string" &&
            (cell["value"] === undefined ||
              cell["value"] === null ||
              typeof cell["value"] === "string" ||
              typeof cell["value"] === "number" ||
              typeof cell["value"] === "boolean") &&
            (cell["formula"] === undefined || typeof cell["formula"] === "string") &&
            (cell["format"] === undefined || typeof cell["format"] === "string")
          );
        })
      );
    })
  );
}

interface ViewportSubscriptionState {
  subscription: ViewportPatchSubscription;
  listener: (patch: Uint8Array) => void;
  nextVersion: number;
  knownStyleIds: Set<string>;
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
  private engine: WorkerEngine | null = null;
  private bootstrapOptions: WorkbookWorkerBootstrapOptions | null = null;
  private engineSubscription: (() => void) | null = null;
  private externalSyncState: SyncState | null = null;
  private readonly viewportSubscriptions = new Set<ViewportSubscriptionState>();
  private readonly formatIds = new Map<string, number>([["", 0]]);
  private readonly styles = new Map<string, CellStyleRecord>([
    [DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }],
  ]);
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
    this.externalSyncState = null;
    const engine = new SpreadsheetEngine({
      workbookName: options.documentId,
      replicaId: options.replicaId,
    });
    this.engine = engine;
    await engine.ready();
    let initialServerSnapshot: ServerSnapshotSeed | null = null;

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
    if (options.baseUrl) {
      try {
        initialServerSnapshot = await this.fetchLatestServerSnapshot(
          options.baseUrl,
          options.documentId,
        );
        if (initialServerSnapshot) {
          engine.importSnapshot(initialServerSnapshot.snapshot);
        }
      } catch {
        initialServerSnapshot = null;
      }
    }
    if (engine.workbook.sheetsByName.size === 0) {
      engine.createSheet("Sheet1");
    }

    this.engineSubscription = engine.subscribe((event) => {
      void this.persistState();
      this.broadcastViewportPatches(event.metrics);
    });
    await this.persistState();

    if (options.baseUrl) {
      try {
        await engine.connectSyncClient(
          this.createSyncClient({
            documentId: options.documentId,
            replicaId: options.replicaId,
            baseUrl: options.baseUrl,
            ...(initialServerSnapshot ? { initialServerCursor: initialServerSnapshot.cursor } : {}),
          }),
        );
      } catch {
        await engine.disconnectSyncClient();
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
      syncState: this.externalSyncState ?? engine.getSyncState(),
    };
  }

  async replaceSnapshot(snapshot: WorkbookSnapshot): Promise<WorkbookWorkerStateSnapshot> {
    const engine = this.requireEngine();
    engine.importSnapshot(snapshot);
    await this.persistState();
    this.broadcastViewportPatches(engine.getLastMetrics());
    return this.getRuntimeState();
  }

  setExternalSyncState(syncState: SyncState | null): WorkbookWorkerStateSnapshot {
    this.externalSyncState = syncState;
    return this.getRuntimeState();
  }

  exportSnapshot(): WorkbookSnapshot {
    return this.requireEngine().exportSnapshot();
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
    this.styles.clear();
    this.styles.set(DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID });
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

  private async fetchLatestServerSnapshot(
    baseUrl: string,
    documentId: string,
  ): Promise<ServerSnapshotSeed | null> {
    const url = new URL(`/v1/documents/${encodeURIComponent(documentId)}/snapshot/latest`, baseUrl);
    const response = await fetch(url);
    if (response.status === 404) {
      return null;
    }
    if (!response.ok) {
      throw new Error(`Failed to fetch latest server snapshot (${response.status})`);
    }
    const cursor = Number.parseInt(response.headers.get("x-bilig-snapshot-cursor") ?? "0", 10);
    const parsed: unknown = JSON.parse(await response.text());
    if (!isWorkbookSnapshot(parsed)) {
      throw new Error("Invalid workbook snapshot payload");
    }
    return {
      cursor: Number.isFinite(cursor) ? cursor : 0,
      snapshot: parsed,
    };
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
    const styles: CellStyleRecord[] = [];
    const cells: ViewportPatchedCell[] = [];

    for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
      for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
        const address = formatAddress(row, col);
        const key = `${viewport.sheetName}!${address}`;
        const snapshot = engine.workbook.getSheet(viewport.sheetName)
          ? engine.getCell(viewport.sheetName, address)
          : this.emptyCellSnapshot(viewport.sheetName, address);
        const formatId = this.getFormatId(snapshot.format);
        const style = this.getStyleRecord(snapshot.styleId ?? DEFAULT_STYLE_ID);
        if (full || !state.knownStyleIds.has(style.id)) {
          state.knownStyleIds.add(style.id);
          styles.push(style);
        }
        const patchedCell = this.buildPatchedCell(snapshot, row, col, formatId, style.id);
        const signature = JSON.stringify([
          patchedCell.snapshot.version,
          patchedCell.snapshot.formula ?? "",
          patchedCell.snapshot.input ?? null,
          patchedCell.snapshot.format ?? "",
          patchedCell.snapshot.styleId ?? "",
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
      styles,
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
    styleId: string,
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
      styleId,
    };
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
}
