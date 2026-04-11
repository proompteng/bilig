import { parseCellAddress } from "@bilig/formula";
import type { GridEngineLike } from "@bilig/grid";
import { ValueTag, type CellSnapshot, type CellStyleRecord, type Viewport } from "@bilig/protocol";
import {
  decodeViewportPatch,
  type ViewportPatch,
  type WorkerEngineClient,
} from "@bilig/worker-transport";
import { selectProjectedViewportKeysToEvict } from "./projected-viewport-cache-pruning.js";
import {
  ackProjectedViewportLocalAxisSize,
  rollbackProjectedViewportLocalAxisHidden,
  rollbackProjectedViewportLocalAxisSize,
  setProjectedViewportLocalAxisHidden,
  setProjectedViewportLocalAxisSize,
  type ProjectedViewportLocalAxisResult,
  type ProjectedViewportLocalAxisState,
} from "./projected-viewport-local-axis-state.js";
import {
  applyProjectedViewportPatch,
  cellSnapshotSignature,
  shouldKeepCurrentSnapshot,
} from "./projected-viewport-patch-application.js";

const EMPTY_WIDTHS: Readonly<Record<number, number>> = Object.freeze({});
const EMPTY_HEIGHTS: Readonly<Record<number, number>> = Object.freeze({});
const EMPTY_HIDDEN_AXES: Readonly<Record<number, true>> = Object.freeze({});
const EMPTY_FREEZE = 0;
const DEFAULT_STYLE_ID = "style-0";
const MAX_CACHED_CELLS_PER_SHEET = 6000;
type CellItem = readonly [number, number];

interface CellSubscription {
  sheetName: string;
  addresses: Set<string>;
  listener: () => void;
}

export class ProjectedViewportStore implements GridEngineLike {
  readonly workbook = {
    getSheet: (sheetName: string) => {
      if (!this.knownSheets.has(sheetName)) {
        return undefined;
      }
      const sheetCellKeys = this.cellKeysBySheet.get(sheetName);
      return {
        grid: {
          forEachCellEntry: (listener: (cellIndex: number, row: number, col: number) => void) => {
            let index = 0;
            sheetCellKeys?.forEach((key) => {
              const snapshot = this.cellSnapshots.get(key);
              if (!snapshot) {
                return;
              }
              const parsed = parseCellAddress(snapshot.address, snapshot.sheetName);
              listener(index++, parsed.row, parsed.col);
            });
          },
        },
      };
    },
  };

  private readonly cellSnapshots = new Map<string, CellSnapshot>();
  private readonly cellKeysBySheet = new Map<string, Set<string>>();
  private readonly cellStyles = new Map<string, CellStyleRecord>([
    [DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }],
  ]);
  private readonly cellSubscriptions = new Set<CellSubscription>();
  private readonly listeners = new Set<() => void>();
  private readonly columnSizesBySheet = new Map<string, Record<number, number>>();
  private readonly columnWidthsBySheet = new Map<string, Record<number, number>>();
  private readonly pendingColumnWidthsBySheet = new Map<string, Record<number, number>>();
  private readonly rowSizesBySheet = new Map<string, Record<number, number>>();
  private readonly rowHeightsBySheet = new Map<string, Record<number, number>>();
  private readonly pendingRowHeightsBySheet = new Map<string, Record<number, number>>();
  private readonly hiddenColumnsBySheet = new Map<string, Record<number, true>>();
  private readonly hiddenRowsBySheet = new Map<string, Record<number, true>>();
  private readonly freezeRowsBySheet = new Map<string, number>();
  private readonly freezeColsBySheet = new Map<string, number>();
  private readonly knownSheets = new Set<string>();
  private readonly activeViewportKeysBySheet = new Map<string, Set<string>>();
  private readonly activeViewports = new Map<string, Viewport>();
  private readonly cellAccessTicks = new Map<string, number>();
  private nextCellAccessTick = 1;

  constructor(private readonly client?: WorkerEngineClient) {}

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  peekCell(sheetName: string, address: string): CellSnapshot | undefined {
    const key = `${sheetName}!${address}`;
    const snapshot = this.cellSnapshots.get(key);
    if (snapshot) {
      this.touchCellKey(key);
    }
    return snapshot;
  }

  getColumnWidths(sheetName: string): Readonly<Record<number, number>> {
    return this.columnWidthsBySheet.get(sheetName) ?? EMPTY_WIDTHS;
  }

  getColumnSizes(sheetName: string): Readonly<Record<number, number>> {
    return this.columnSizesBySheet.get(sheetName) ?? EMPTY_WIDTHS;
  }

  getRowHeights(sheetName: string): Readonly<Record<number, number>> {
    return this.rowHeightsBySheet.get(sheetName) ?? EMPTY_HEIGHTS;
  }

  getRowSizes(sheetName: string): Readonly<Record<number, number>> {
    return this.rowSizesBySheet.get(sheetName) ?? EMPTY_HEIGHTS;
  }

  getHiddenColumns(sheetName: string): Readonly<Record<number, true>> {
    return this.hiddenColumnsBySheet.get(sheetName) ?? EMPTY_HIDDEN_AXES;
  }

  getHiddenRows(sheetName: string): Readonly<Record<number, true>> {
    return this.hiddenRowsBySheet.get(sheetName) ?? EMPTY_HIDDEN_AXES;
  }

  getFreezeRows(sheetName: string): number {
    return this.freezeRowsBySheet.get(sheetName) ?? EMPTY_FREEZE;
  }

  getFreezeCols(sheetName: string): number {
    return this.freezeColsBySheet.get(sheetName) ?? EMPTY_FREEZE;
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    return this.peekCell(sheetName, address) ?? this.emptyCellSnapshot(sheetName, address);
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    if (!styleId) {
      return this.cellStyles.get(DEFAULT_STYLE_ID);
    }
    return this.cellStyles.get(styleId) ?? this.cellStyles.get(DEFAULT_STYLE_ID);
  }

  setCellSnapshot(snapshot: CellSnapshot): void {
    const key = `${snapshot.sheetName}!${snapshot.address}`;
    const current = this.cellSnapshots.get(key);
    if (current) {
      if (shouldKeepCurrentSnapshot(current, snapshot)) {
        return;
      }
      if (cellSnapshotSignature(current) === cellSnapshotSignature(snapshot)) {
        return;
      }
    }
    this.knownSheets.add(snapshot.sheetName);
    this.cellSnapshots.set(key, snapshot);
    this.touchCellKey(key);
    this.sheetCellKeys(snapshot.sheetName).add(key);
    this.notifyCellSubscriptions(new Set([key]));
    this.listeners.forEach((listener) => listener());
  }

  setColumnWidth(sheetName: string, columnIndex: number, width: number): void {
    const nextState = setProjectedViewportLocalAxisSize({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.columnSizesBySheet,
        renderedSizesBySheet: this.columnWidthsBySheet,
        pendingSizesBySheet: this.pendingColumnWidthsBySheet,
        hiddenAxesBySheet: this.hiddenColumnsBySheet,
      }),
      index: columnIndex,
      size: width,
    });
    if (!nextState.changed) {
      return;
    }
    this.knownSheets.add(sheetName);
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
    });
    this.listeners.forEach((listener) => listener());
  }

  ackColumnWidth(sheetName: string, columnIndex: number, width: number): void {
    const nextState = ackProjectedViewportLocalAxisSize({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.columnSizesBySheet,
        renderedSizesBySheet: this.columnWidthsBySheet,
        pendingSizesBySheet: this.pendingColumnWidthsBySheet,
        hiddenAxesBySheet: this.hiddenColumnsBySheet,
      }),
      index: columnIndex,
      size: width,
    });
    if (!nextState.changed) {
      return;
    }
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
    });
  }

  rollbackColumnWidth(sheetName: string, columnIndex: number, width: number | undefined): void {
    const nextState = rollbackProjectedViewportLocalAxisSize({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.columnSizesBySheet,
        renderedSizesBySheet: this.columnWidthsBySheet,
        pendingSizesBySheet: this.pendingColumnWidthsBySheet,
        hiddenAxesBySheet: this.hiddenColumnsBySheet,
      }),
      index: columnIndex,
      size: width,
    });
    if (!nextState.changed) {
      return;
    }
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
    });
    this.listeners.forEach((listener) => listener());
  }

  setColumnHidden(sheetName: string, columnIndex: number, hidden: boolean, size: number): void {
    const nextState = setProjectedViewportLocalAxisHidden({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.columnSizesBySheet,
        renderedSizesBySheet: this.columnWidthsBySheet,
        pendingSizesBySheet: this.pendingColumnWidthsBySheet,
        hiddenAxesBySheet: this.hiddenColumnsBySheet,
      }),
      index: columnIndex,
      hidden,
      size,
    });
    if (!nextState.changed) {
      return;
    }
    this.knownSheets.add(sheetName);
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
    });
    this.listeners.forEach((listener) => listener());
  }

  rollbackColumnHidden(
    sheetName: string,
    columnIndex: number,
    previous: { hidden: boolean; size: number | undefined },
  ): void {
    const nextState = rollbackProjectedViewportLocalAxisHidden({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.columnSizesBySheet,
        renderedSizesBySheet: this.columnWidthsBySheet,
        pendingSizesBySheet: this.pendingColumnWidthsBySheet,
        hiddenAxesBySheet: this.hiddenColumnsBySheet,
      }),
      index: columnIndex,
      previous,
    });
    if (!nextState.changed) {
      return;
    }
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
    });
    this.listeners.forEach((listener) => listener());
  }

  setRowHeight(sheetName: string, rowIndex: number, height: number): void {
    const nextState = setProjectedViewportLocalAxisSize({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.rowSizesBySheet,
        renderedSizesBySheet: this.rowHeightsBySheet,
        pendingSizesBySheet: this.pendingRowHeightsBySheet,
        hiddenAxesBySheet: this.hiddenRowsBySheet,
      }),
      index: rowIndex,
      size: height,
    });
    if (!nextState.changed) {
      return;
    }
    this.knownSheets.add(sheetName);
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
    });
    this.listeners.forEach((listener) => listener());
  }

  ackRowHeight(sheetName: string, rowIndex: number, height: number): void {
    const nextState = ackProjectedViewportLocalAxisSize({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.rowSizesBySheet,
        renderedSizesBySheet: this.rowHeightsBySheet,
        pendingSizesBySheet: this.pendingRowHeightsBySheet,
        hiddenAxesBySheet: this.hiddenRowsBySheet,
      }),
      index: rowIndex,
      size: height,
    });
    if (!nextState.changed) {
      return;
    }
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
    });
  }

  rollbackRowHeight(sheetName: string, rowIndex: number, height: number | undefined): void {
    const nextState = rollbackProjectedViewportLocalAxisSize({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.rowSizesBySheet,
        renderedSizesBySheet: this.rowHeightsBySheet,
        pendingSizesBySheet: this.pendingRowHeightsBySheet,
        hiddenAxesBySheet: this.hiddenRowsBySheet,
      }),
      index: rowIndex,
      size: height,
    });
    if (!nextState.changed) {
      return;
    }
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
    });
    this.listeners.forEach((listener) => listener());
  }

  setRowHidden(sheetName: string, rowIndex: number, hidden: boolean, size: number): void {
    const nextState = setProjectedViewportLocalAxisHidden({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.rowSizesBySheet,
        renderedSizesBySheet: this.rowHeightsBySheet,
        pendingSizesBySheet: this.pendingRowHeightsBySheet,
        hiddenAxesBySheet: this.hiddenRowsBySheet,
      }),
      index: rowIndex,
      hidden,
      size,
    });
    if (!nextState.changed) {
      return;
    }
    this.knownSheets.add(sheetName);
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
    });
    this.listeners.forEach((listener) => listener());
  }

  rollbackRowHidden(
    sheetName: string,
    rowIndex: number,
    previous: { hidden: boolean; size: number | undefined },
  ): void {
    const nextState = rollbackProjectedViewportLocalAxisHidden({
      state: this.readLocalAxisState({
        sheetName,
        sizesBySheet: this.rowSizesBySheet,
        renderedSizesBySheet: this.rowHeightsBySheet,
        pendingSizesBySheet: this.pendingRowHeightsBySheet,
        hiddenAxesBySheet: this.hiddenRowsBySheet,
      }),
      index: rowIndex,
      previous,
    });
    if (!nextState.changed) {
      return;
    }
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
    });
    this.listeners.forEach((listener) => listener());
  }

  setKnownSheets(sheetNames: readonly string[]): void {
    if (
      sheetNames.length === this.knownSheets.size &&
      sheetNames.every((sheetName) => this.knownSheets.has(sheetName))
    ) {
      return;
    }
    const removedSheets = [...this.knownSheets].filter(
      (sheetName) => !sheetNames.includes(sheetName),
    );
    this.knownSheets.clear();
    sheetNames.forEach((sheetName) => this.knownSheets.add(sheetName));
    removedSheets.forEach((sheetName) => this.dropSheetCache(sheetName));
    this.listeners.forEach((listener) => listener());
  }

  subscribeCells(
    sheetName: string,
    addresses: readonly string[],
    listener: () => void,
  ): () => void {
    const subscription: CellSubscription = {
      sheetName,
      addresses: new Set(addresses),
      listener,
    };
    this.cellSubscriptions.add(subscription);
    return () => {
      this.cellSubscriptions.delete(subscription);
    };
  }

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: CellItem }[]) => void,
  ): () => void {
    if (!this.client) {
      throw new Error("Worker viewport subscriptions require a worker engine client");
    }
    const viewportKey = `${sheetName}:${viewport.rowStart}:${viewport.rowEnd}:${viewport.colStart}:${viewport.colEnd}`;
    this.activeViewports.set(viewportKey, viewport);
    const sheetViewportKeys = this.activeViewportKeysBySheet.get(sheetName) ?? new Set<string>();
    sheetViewportKeys.add(viewportKey);
    this.activeViewportKeysBySheet.set(sheetName, sheetViewportKeys);
    const unsubscribe = this.client.subscribeViewportPatches(
      { sheetName, ...viewport },
      (bytes: Uint8Array) => {
        const damage = this.applyPatch(decodeViewportPatch(bytes));
        listener(damage);
      },
    );
    return () => {
      unsubscribe();
      this.activeViewports.delete(viewportKey);
      const nextSheetViewportKeys = this.activeViewportKeysBySheet.get(sheetName);
      nextSheetViewportKeys?.delete(viewportKey);
      if (nextSheetViewportKeys && nextSheetViewportKeys.size === 0) {
        this.activeViewportKeysBySheet.delete(sheetName);
      }
      this.pruneSheetCache(sheetName);
    };
  }

  applyViewportPatch(patch: ViewportPatch): readonly { cell: CellItem }[] {
    return this.applyPatch(patch);
  }

  private applyPatch(patch: ReturnType<typeof decodeViewportPatch>): readonly { cell: CellItem }[] {
    const result = applyProjectedViewportPatch({
      state: {
        cellSnapshots: this.cellSnapshots,
        cellKeysBySheet: this.cellKeysBySheet,
        cellStyles: this.cellStyles,
        columnSizesBySheet: this.columnSizesBySheet,
        columnWidthsBySheet: this.columnWidthsBySheet,
        pendingColumnWidthsBySheet: this.pendingColumnWidthsBySheet,
        rowSizesBySheet: this.rowSizesBySheet,
        rowHeightsBySheet: this.rowHeightsBySheet,
        pendingRowHeightsBySheet: this.pendingRowHeightsBySheet,
        hiddenColumnsBySheet: this.hiddenColumnsBySheet,
        hiddenRowsBySheet: this.hiddenRowsBySheet,
        freezeRowsBySheet: this.freezeRowsBySheet,
        freezeColsBySheet: this.freezeColsBySheet,
        knownSheets: this.knownSheets,
      },
      patch,
      touchCellKey: (key) => this.touchCellKey(key),
    });
    this.pruneSheetCache(patch.viewport.sheetName);
    this.notifyCellSubscriptions(result.changedKeys);
    if (result.damage.length > 0 || result.axisChanged || result.freezeChanged) {
      this.listeners.forEach((listener) => listener());
    }
    return result.damage;
  }

  private readLocalAxisState(args: {
    sheetName: string;
    sizesBySheet: Map<string, Record<number, number>>;
    renderedSizesBySheet: Map<string, Record<number, number>>;
    pendingSizesBySheet: Map<string, Record<number, number>>;
    hiddenAxesBySheet: Map<string, Record<number, true>>;
  }): ProjectedViewportLocalAxisState {
    return {
      sizes: args.sizesBySheet.get(args.sheetName) ?? {},
      renderedSizes: args.renderedSizesBySheet.get(args.sheetName) ?? {},
      pendingSizes: args.pendingSizesBySheet.get(args.sheetName) ?? {},
      hiddenAxes: args.hiddenAxesBySheet.get(args.sheetName) ?? {},
    };
  }

  private writeLocalAxisState(args: {
    sheetName: string;
    sizesBySheet: Map<string, Record<number, number>>;
    renderedSizesBySheet: Map<string, Record<number, number>>;
    pendingSizesBySheet: Map<string, Record<number, number>>;
    hiddenAxesBySheet: Map<string, Record<number, true>>;
    nextState: ProjectedViewportLocalAxisResult;
  }): void {
    args.sizesBySheet.set(args.sheetName, args.nextState.sizes);
    args.renderedSizesBySheet.set(args.sheetName, args.nextState.renderedSizes);
    if (Object.keys(args.nextState.pendingSizes).length === 0) {
      args.pendingSizesBySheet.delete(args.sheetName);
    } else {
      args.pendingSizesBySheet.set(args.sheetName, args.nextState.pendingSizes);
    }
    if (Object.keys(args.nextState.hiddenAxes).length === 0) {
      args.hiddenAxesBySheet.delete(args.sheetName);
    } else {
      args.hiddenAxesBySheet.set(args.sheetName, args.nextState.hiddenAxes);
    }
  }

  private sheetCellKeys(sheetName: string): Set<string> {
    const existing = this.cellKeysBySheet.get(sheetName);
    if (existing) {
      return existing;
    }
    const created = new Set<string>();
    this.cellKeysBySheet.set(sheetName, created);
    return created;
  }

  private pruneSheetCache(sheetName: string): void {
    const sheetCellKeys = this.cellKeysBySheet.get(sheetName);
    if (!sheetCellKeys || sheetCellKeys.size <= MAX_CACHED_CELLS_PER_SHEET) {
      return;
    }
    const activeViewportKeys = this.activeViewportKeysBySheet.get(sheetName);
    const activeViewports =
      activeViewportKeys && activeViewportKeys.size > 0
        ? [...activeViewportKeys]
            .map((key) => this.activeViewports.get(key))
            .filter((viewport): viewport is Viewport => viewport !== undefined)
        : [];
    const pinnedKeys = new Set<string>();
    this.cellSubscriptions.forEach((subscription) => {
      if (subscription.sheetName !== sheetName) {
        return;
      }
      subscription.addresses.forEach((address) => pinnedKeys.add(`${sheetName}!${address}`));
    });
    const keysToEvict = selectProjectedViewportKeysToEvict({
      sheetCellKeys: Array.from(sheetCellKeys),
      cellSnapshots: this.cellSnapshots,
      cellAccessTicks: this.cellAccessTicks,
      pinnedKeys,
      activeViewports,
      maxCachedCellsPerSheet: MAX_CACHED_CELLS_PER_SHEET,
    });
    keysToEvict.forEach((key) => {
      this.cellSnapshots.delete(key);
      this.cellAccessTicks.delete(key);
      sheetCellKeys.delete(key);
    });
  }

  private dropSheetCache(sheetName: string): void {
    this.cellKeysBySheet.get(sheetName)?.forEach((key) => {
      this.cellSnapshots.delete(key);
      this.cellAccessTicks.delete(key);
    });
    this.cellKeysBySheet.delete(sheetName);
    this.columnSizesBySheet.delete(sheetName);
    this.columnWidthsBySheet.delete(sheetName);
    this.pendingColumnWidthsBySheet.delete(sheetName);
    this.rowSizesBySheet.delete(sheetName);
    this.rowHeightsBySheet.delete(sheetName);
    this.pendingRowHeightsBySheet.delete(sheetName);
    this.hiddenColumnsBySheet.delete(sheetName);
    this.hiddenRowsBySheet.delete(sheetName);
    this.freezeRowsBySheet.delete(sheetName);
    this.freezeColsBySheet.delete(sheetName);
    const viewportKeys = this.activeViewportKeysBySheet.get(sheetName);
    viewportKeys?.forEach((key) => this.activeViewports.delete(key));
    this.activeViewportKeysBySheet.delete(sheetName);
  }

  private touchCellKey(key: string): void {
    this.cellAccessTicks.set(key, this.nextCellAccessTick++);
  }

  private notifyCellSubscriptions(changedKeys: ReadonlySet<string>): void {
    this.cellSubscriptions.forEach((subscription) => {
      for (const address of subscription.addresses) {
        if (changedKeys.has(`${subscription.sheetName}!${address}`)) {
          subscription.listener();
          return;
        }
      }
    });
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
