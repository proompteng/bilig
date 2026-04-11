import type { GridEngineLike } from "@bilig/grid";
import type { CellSnapshot, CellStyleRecord, Viewport } from "@bilig/protocol";
import {
  decodeViewportPatch,
  type ViewportPatch,
  type WorkerEngineClient,
} from "@bilig/worker-transport";
import {
  ackProjectedViewportLocalAxisSize,
  rollbackProjectedViewportLocalAxisHidden,
  rollbackProjectedViewportLocalAxisSize,
  setProjectedViewportLocalAxisHidden,
  setProjectedViewportLocalAxisSize,
  type ProjectedViewportLocalAxisResult,
  type ProjectedViewportLocalAxisState,
} from "./projected-viewport-local-axis-state.js";
import { applyProjectedViewportPatch } from "./projected-viewport-patch-application.js";
import { ProjectedViewportCellCache } from "./projected-viewport-cell-cache.js";

const EMPTY_WIDTHS: Readonly<Record<number, number>> = Object.freeze({});
const EMPTY_HEIGHTS: Readonly<Record<number, number>> = Object.freeze({});
const EMPTY_HIDDEN_AXES: Readonly<Record<number, true>> = Object.freeze({});
const EMPTY_FREEZE = 0;
const MAX_CACHED_CELLS_PER_SHEET = 6000;
type CellItem = readonly [number, number];

export class ProjectedViewportStore implements GridEngineLike {
  private readonly cellCache = new ProjectedViewportCellCache({
    maxCachedCellsPerSheet: MAX_CACHED_CELLS_PER_SHEET,
  });

  readonly workbook = {
    getSheet: (sheetName: string) => this.cellCache.getSheet(sheetName),
  };
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

  constructor(private readonly client?: WorkerEngineClient) {}

  subscribe(listener: () => void): () => void {
    return this.cellCache.subscribe(listener);
  }

  peekCell(sheetName: string, address: string): CellSnapshot | undefined {
    return this.cellCache.peekCell(sheetName, address);
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
    return this.cellCache.getCell(sheetName, address);
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    return this.cellCache.getCellStyle(styleId);
  }

  setCellSnapshot(snapshot: CellSnapshot): void {
    this.cellCache.setCellSnapshot(snapshot);
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
    this.cellCache.markSheetKnown(sheetName);
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
    });
    this.cellCache.notifyListeners();
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
    this.cellCache.notifyListeners();
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
    this.cellCache.markSheetKnown(sheetName);
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
    });
    this.cellCache.notifyListeners();
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
    this.cellCache.notifyListeners();
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
    this.cellCache.markSheetKnown(sheetName);
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
    });
    this.cellCache.notifyListeners();
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
    this.cellCache.notifyListeners();
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
    this.cellCache.markSheetKnown(sheetName);
    this.writeLocalAxisState({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
    });
    this.cellCache.notifyListeners();
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
    this.cellCache.notifyListeners();
  }

  setKnownSheets(sheetNames: readonly string[]): void {
    const removedSheets = this.cellCache.setKnownSheets(sheetNames);
    removedSheets.forEach((sheetName) => this.dropAxisState(sheetName));
  }

  subscribeCells(
    sheetName: string,
    addresses: readonly string[],
    listener: () => void,
  ): () => void {
    return this.cellCache.subscribeCells(sheetName, addresses, listener);
  }

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: CellItem }[]) => void,
  ): () => void {
    if (!this.client) {
      throw new Error("Worker viewport subscriptions require a worker engine client");
    }
    const stopTrackingViewport = this.cellCache.trackViewport(sheetName, viewport);
    const unsubscribe = this.client.subscribeViewportPatches(
      { sheetName, ...viewport },
      (bytes: Uint8Array) => {
        const damage = this.applyPatch(decodeViewportPatch(bytes));
        listener(damage);
      },
    );
    return () => {
      unsubscribe();
      stopTrackingViewport();
    };
  }

  applyViewportPatch(patch: ViewportPatch): readonly { cell: CellItem }[] {
    return this.applyPatch(patch);
  }

  private applyPatch(patch: ReturnType<typeof decodeViewportPatch>): readonly { cell: CellItem }[] {
    const result = applyProjectedViewportPatch({
      state: {
        ...this.cellCache.getPatchState(),
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
      },
      patch,
      touchCellKey: (key) => this.cellCache.touchCellKey(key),
    });
    return this.cellCache.applyPatchResult(patch.viewport.sheetName, result);
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

  private dropAxisState(sheetName: string): void {
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
  }
}
