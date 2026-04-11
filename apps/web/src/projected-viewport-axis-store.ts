import {
  ackProjectedViewportLocalAxisSize,
  rollbackProjectedViewportLocalAxisHidden,
  rollbackProjectedViewportLocalAxisSize,
  setProjectedViewportLocalAxisHidden,
  setProjectedViewportLocalAxisSize,
  type ProjectedViewportLocalAxisResult,
  type ProjectedViewportLocalAxisState,
} from "./projected-viewport-local-axis-state.js";

const EMPTY_WIDTHS: Readonly<Record<number, number>> = Object.freeze({});
const EMPTY_HEIGHTS: Readonly<Record<number, number>> = Object.freeze({});
const EMPTY_HIDDEN_AXES: Readonly<Record<number, true>> = Object.freeze({});
const EMPTY_FREEZE = 0;

export class ProjectedViewportAxisStore {
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

  constructor(
    private readonly options: {
      markSheetKnown?: (sheetName: string) => void;
      notifyListeners?: () => void;
    } = {},
  ) {}

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

  getPatchState(): {
    columnSizesBySheet: Map<string, Record<number, number>>;
    columnWidthsBySheet: Map<string, Record<number, number>>;
    pendingColumnWidthsBySheet: Map<string, Record<number, number>>;
    rowSizesBySheet: Map<string, Record<number, number>>;
    rowHeightsBySheet: Map<string, Record<number, number>>;
    pendingRowHeightsBySheet: Map<string, Record<number, number>>;
    hiddenColumnsBySheet: Map<string, Record<number, true>>;
    hiddenRowsBySheet: Map<string, Record<number, true>>;
    freezeRowsBySheet: Map<string, number>;
    freezeColsBySheet: Map<string, number>;
  } {
    return {
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
    };
  }

  dropSheets(sheetNames: readonly string[]): void {
    sheetNames.forEach((sheetName) => this.dropAxisState(sheetName));
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
    this.commitAxisChange({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
      markSheetKnown: true,
      notifyListeners: true,
    });
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
    this.commitAxisChange({
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
    this.commitAxisChange({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
      notifyListeners: true,
    });
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
    this.commitAxisChange({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
      markSheetKnown: true,
      notifyListeners: true,
    });
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
    this.commitAxisChange({
      sheetName,
      sizesBySheet: this.columnSizesBySheet,
      renderedSizesBySheet: this.columnWidthsBySheet,
      pendingSizesBySheet: this.pendingColumnWidthsBySheet,
      hiddenAxesBySheet: this.hiddenColumnsBySheet,
      nextState,
      notifyListeners: true,
    });
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
    this.commitAxisChange({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
      markSheetKnown: true,
      notifyListeners: true,
    });
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
    this.commitAxisChange({
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
    this.commitAxisChange({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
      notifyListeners: true,
    });
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
    this.commitAxisChange({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
      markSheetKnown: true,
      notifyListeners: true,
    });
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
    this.commitAxisChange({
      sheetName,
      sizesBySheet: this.rowSizesBySheet,
      renderedSizesBySheet: this.rowHeightsBySheet,
      pendingSizesBySheet: this.pendingRowHeightsBySheet,
      hiddenAxesBySheet: this.hiddenRowsBySheet,
      nextState,
      notifyListeners: true,
    });
  }

  private commitAxisChange(args: {
    sheetName: string;
    sizesBySheet: Map<string, Record<number, number>>;
    renderedSizesBySheet: Map<string, Record<number, number>>;
    pendingSizesBySheet: Map<string, Record<number, number>>;
    hiddenAxesBySheet: Map<string, Record<number, true>>;
    nextState: ProjectedViewportLocalAxisResult;
    markSheetKnown?: boolean;
    notifyListeners?: boolean;
  }): void {
    if (!args.nextState.changed) {
      return;
    }
    if (args.markSheetKnown) {
      this.options.markSheetKnown?.(args.sheetName);
    }
    this.writeLocalAxisState({
      sheetName: args.sheetName,
      sizesBySheet: args.sizesBySheet,
      renderedSizesBySheet: args.renderedSizesBySheet,
      pendingSizesBySheet: args.pendingSizesBySheet,
      hiddenAxesBySheet: args.hiddenAxesBySheet,
      nextState: args.nextState,
    });
    if (args.notifyListeners) {
      this.options.notifyListeners?.();
    }
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
