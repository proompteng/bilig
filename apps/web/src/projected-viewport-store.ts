import { parseCellAddress } from "@bilig/formula";
import type { GridEngineLike } from "@bilig/grid";
import { ValueTag, type CellSnapshot, type CellStyleRecord, type Viewport } from "@bilig/protocol";
import {
  decodeViewportPatch,
  type ViewportPatch,
  type WorkerEngineClient,
} from "@bilig/worker-transport";
import { applyProjectedViewportAxisPatches } from "./projected-viewport-axis-patches.js";
import { selectProjectedViewportKeysToEvict } from "./projected-viewport-cache-pruning.js";

const EMPTY_WIDTHS: Readonly<Record<number, number>> = Object.freeze({});
const EMPTY_HEIGHTS: Readonly<Record<number, number>> = Object.freeze({});
const EMPTY_HIDDEN_AXES: Readonly<Record<number, true>> = Object.freeze({});
const EMPTY_FREEZE = 0;
const DEFAULT_STYLE_ID = "style-0";
const MAX_CACHED_CELLS_PER_SHEET = 6000;
type CellItem = readonly [number, number];

function snapshotValueKey(snapshot: CellSnapshot): string {
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
  return "empty";
}

function cellSnapshotSignature(snapshot: CellSnapshot): string {
  return [
    snapshot.version,
    snapshot.flags,
    snapshot.formula ?? "",
    snapshot.format ?? "",
    snapshot.styleId ?? "",
    snapshot.numberFormatId ?? "",
    snapshot.input ?? "",
    snapshotValueKey(snapshot),
  ].join("|");
}

function shouldKeepCurrentSnapshot(current: CellSnapshot, incoming: CellSnapshot): boolean {
  if (
    current.formula !== undefined &&
    incoming.formula === undefined &&
    incoming.input === undefined
  ) {
    return true;
  }
  if (current.version > incoming.version) {
    return true;
  }
  if (current.version < incoming.version) {
    return false;
  }
  // Zero source/eval rows can briefly lag the worker patch and drop formula metadata while
  // keeping the same cell version. Preserve the local formula snapshot until source metadata
  // catches up.
  return current.formula !== undefined && incoming.formula === undefined;
}

function cellStyleSignature(style: CellStyleRecord): string {
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

function isCellInsideViewport(snapshot: CellSnapshot, viewport: Viewport): boolean {
  const parsed = parseCellAddress(snapshot.address, snapshot.sheetName);
  return (
    parsed.row >= viewport.rowStart &&
    parsed.row <= viewport.rowEnd &&
    parsed.col >= viewport.colStart &&
    parsed.col <= viewport.colEnd
  );
}

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
    const currentWidth = this.columnWidthsBySheet.get(sheetName)?.[columnIndex];
    if (currentWidth === width) {
      return;
    }
    this.knownSheets.add(sheetName);
    const sizes = { ...this.columnSizesBySheet.get(sheetName) };
    sizes[columnIndex] = width;
    this.columnSizesBySheet.set(sheetName, sizes);
    const widths = { ...this.columnWidthsBySheet.get(sheetName) };
    widths[columnIndex] = this.hiddenColumnsBySheet.get(sheetName)?.[columnIndex] ? 0 : width;
    this.columnWidthsBySheet.set(sheetName, widths);
    const pending = { ...this.pendingColumnWidthsBySheet.get(sheetName) };
    pending[columnIndex] = width;
    this.pendingColumnWidthsBySheet.set(sheetName, pending);
    this.listeners.forEach((listener) => listener());
  }

  ackColumnWidth(sheetName: string, columnIndex: number, width: number): void {
    const pendingWidths = this.pendingColumnWidthsBySheet.get(sheetName);
    if (!pendingWidths || pendingWidths[columnIndex] !== width) {
      return;
    }
    const nextPendingWidths = { ...pendingWidths };
    delete nextPendingWidths[columnIndex];
    if (Object.keys(nextPendingWidths).length === 0) {
      this.pendingColumnWidthsBySheet.delete(sheetName);
    } else {
      this.pendingColumnWidthsBySheet.set(sheetName, nextPendingWidths);
    }
  }

  rollbackColumnWidth(sheetName: string, columnIndex: number, width: number | undefined): void {
    const sizes = { ...this.columnSizesBySheet.get(sheetName) };
    if (width === undefined) {
      delete sizes[columnIndex];
    } else {
      sizes[columnIndex] = width;
    }
    if (Object.keys(sizes).length === 0) {
      this.columnSizesBySheet.delete(sheetName);
    } else {
      this.columnSizesBySheet.set(sheetName, sizes);
    }

    const widths = { ...this.columnWidthsBySheet.get(sheetName) };
    if (width === undefined) {
      delete widths[columnIndex];
    } else {
      widths[columnIndex] = this.hiddenColumnsBySheet.get(sheetName)?.[columnIndex] ? 0 : width;
    }
    if (Object.keys(widths).length === 0) {
      this.columnWidthsBySheet.delete(sheetName);
    } else {
      this.columnWidthsBySheet.set(sheetName, widths);
    }

    const pendingWidths = { ...this.pendingColumnWidthsBySheet.get(sheetName) };
    delete pendingWidths[columnIndex];
    if (Object.keys(pendingWidths).length === 0) {
      this.pendingColumnWidthsBySheet.delete(sheetName);
    } else {
      this.pendingColumnWidthsBySheet.set(sheetName, pendingWidths);
    }
    this.listeners.forEach((listener) => listener());
  }

  setColumnHidden(sheetName: string, columnIndex: number, hidden: boolean, size: number): void {
    this.knownSheets.add(sheetName);
    const sizes = { ...this.columnSizesBySheet.get(sheetName) };
    sizes[columnIndex] = size;
    this.columnSizesBySheet.set(sheetName, sizes);

    const hiddenColumns = { ...this.hiddenColumnsBySheet.get(sheetName) };
    if (hidden) {
      hiddenColumns[columnIndex] = true;
    } else {
      delete hiddenColumns[columnIndex];
    }
    if (Object.keys(hiddenColumns).length === 0) {
      this.hiddenColumnsBySheet.delete(sheetName);
    } else {
      this.hiddenColumnsBySheet.set(sheetName, hiddenColumns);
    }

    const widths = { ...this.columnWidthsBySheet.get(sheetName) };
    widths[columnIndex] = hidden ? 0 : size;
    this.columnWidthsBySheet.set(sheetName, widths);
    this.listeners.forEach((listener) => listener());
  }

  rollbackColumnHidden(
    sheetName: string,
    columnIndex: number,
    previous: { hidden: boolean; size: number | undefined },
  ): void {
    const sizes = { ...this.columnSizesBySheet.get(sheetName) };
    if (previous.size === undefined) {
      delete sizes[columnIndex];
    } else {
      sizes[columnIndex] = previous.size;
    }
    if (Object.keys(sizes).length === 0) {
      this.columnSizesBySheet.delete(sheetName);
    } else {
      this.columnSizesBySheet.set(sheetName, sizes);
    }

    const hiddenColumns = { ...this.hiddenColumnsBySheet.get(sheetName) };
    if (previous.hidden) {
      hiddenColumns[columnIndex] = true;
    } else {
      delete hiddenColumns[columnIndex];
    }
    if (Object.keys(hiddenColumns).length === 0) {
      this.hiddenColumnsBySheet.delete(sheetName);
    } else {
      this.hiddenColumnsBySheet.set(sheetName, hiddenColumns);
    }

    const widths = { ...this.columnWidthsBySheet.get(sheetName) };
    if (previous.size === undefined) {
      delete widths[columnIndex];
    } else {
      widths[columnIndex] = previous.hidden ? 0 : previous.size;
    }
    if (Object.keys(widths).length === 0) {
      this.columnWidthsBySheet.delete(sheetName);
    } else {
      this.columnWidthsBySheet.set(sheetName, widths);
    }
    this.listeners.forEach((listener) => listener());
  }

  setRowHeight(sheetName: string, rowIndex: number, height: number): void {
    const currentHeight = this.rowHeightsBySheet.get(sheetName)?.[rowIndex];
    if (currentHeight === height) {
      return;
    }
    this.knownSheets.add(sheetName);
    const sizes = { ...this.rowSizesBySheet.get(sheetName) };
    sizes[rowIndex] = height;
    this.rowSizesBySheet.set(sheetName, sizes);
    const heights = { ...this.rowHeightsBySheet.get(sheetName) };
    heights[rowIndex] = this.hiddenRowsBySheet.get(sheetName)?.[rowIndex] ? 0 : height;
    this.rowHeightsBySheet.set(sheetName, heights);
    const pending = { ...this.pendingRowHeightsBySheet.get(sheetName) };
    pending[rowIndex] = height;
    this.pendingRowHeightsBySheet.set(sheetName, pending);
    this.listeners.forEach((listener) => listener());
  }

  ackRowHeight(sheetName: string, rowIndex: number, height: number): void {
    const pendingHeights = this.pendingRowHeightsBySheet.get(sheetName);
    if (!pendingHeights || pendingHeights[rowIndex] !== height) {
      return;
    }
    const nextPendingHeights = { ...pendingHeights };
    delete nextPendingHeights[rowIndex];
    if (Object.keys(nextPendingHeights).length === 0) {
      this.pendingRowHeightsBySheet.delete(sheetName);
    } else {
      this.pendingRowHeightsBySheet.set(sheetName, nextPendingHeights);
    }
  }

  rollbackRowHeight(sheetName: string, rowIndex: number, height: number | undefined): void {
    const sizes = { ...this.rowSizesBySheet.get(sheetName) };
    if (height === undefined) {
      delete sizes[rowIndex];
    } else {
      sizes[rowIndex] = height;
    }
    if (Object.keys(sizes).length === 0) {
      this.rowSizesBySheet.delete(sheetName);
    } else {
      this.rowSizesBySheet.set(sheetName, sizes);
    }

    const heights = { ...this.rowHeightsBySheet.get(sheetName) };
    if (height === undefined) {
      delete heights[rowIndex];
    } else {
      heights[rowIndex] = this.hiddenRowsBySheet.get(sheetName)?.[rowIndex] ? 0 : height;
    }
    if (Object.keys(heights).length === 0) {
      this.rowHeightsBySheet.delete(sheetName);
    } else {
      this.rowHeightsBySheet.set(sheetName, heights);
    }

    const pendingHeights = { ...this.pendingRowHeightsBySheet.get(sheetName) };
    delete pendingHeights[rowIndex];
    if (Object.keys(pendingHeights).length === 0) {
      this.pendingRowHeightsBySheet.delete(sheetName);
    } else {
      this.pendingRowHeightsBySheet.set(sheetName, pendingHeights);
    }
    this.listeners.forEach((listener) => listener());
  }

  setRowHidden(sheetName: string, rowIndex: number, hidden: boolean, size: number): void {
    this.knownSheets.add(sheetName);
    const sizes = { ...this.rowSizesBySheet.get(sheetName) };
    sizes[rowIndex] = size;
    this.rowSizesBySheet.set(sheetName, sizes);

    const hiddenRows = { ...this.hiddenRowsBySheet.get(sheetName) };
    if (hidden) {
      hiddenRows[rowIndex] = true;
    } else {
      delete hiddenRows[rowIndex];
    }
    if (Object.keys(hiddenRows).length === 0) {
      this.hiddenRowsBySheet.delete(sheetName);
    } else {
      this.hiddenRowsBySheet.set(sheetName, hiddenRows);
    }

    const heights = { ...this.rowHeightsBySheet.get(sheetName) };
    heights[rowIndex] = hidden ? 0 : size;
    this.rowHeightsBySheet.set(sheetName, heights);
    this.listeners.forEach((listener) => listener());
  }

  rollbackRowHidden(
    sheetName: string,
    rowIndex: number,
    previous: { hidden: boolean; size: number | undefined },
  ): void {
    const sizes = { ...this.rowSizesBySheet.get(sheetName) };
    if (previous.size === undefined) {
      delete sizes[rowIndex];
    } else {
      sizes[rowIndex] = previous.size;
    }
    if (Object.keys(sizes).length === 0) {
      this.rowSizesBySheet.delete(sheetName);
    } else {
      this.rowSizesBySheet.set(sheetName, sizes);
    }

    const hiddenRows = { ...this.hiddenRowsBySheet.get(sheetName) };
    if (previous.hidden) {
      hiddenRows[rowIndex] = true;
    } else {
      delete hiddenRows[rowIndex];
    }
    if (Object.keys(hiddenRows).length === 0) {
      this.hiddenRowsBySheet.delete(sheetName);
    } else {
      this.hiddenRowsBySheet.set(sheetName, hiddenRows);
    }

    const heights = { ...this.rowHeightsBySheet.get(sheetName) };
    if (previous.size === undefined) {
      delete heights[rowIndex];
    } else {
      heights[rowIndex] = previous.hidden ? 0 : previous.size;
    }
    if (Object.keys(heights).length === 0) {
      this.rowHeightsBySheet.delete(sheetName);
    } else {
      this.rowHeightsBySheet.set(sheetName, heights);
    }
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
    this.knownSheets.add(patch.viewport.sheetName);
    const nextFreezeRows = patch.freezeRows ?? 0;
    const nextFreezeCols = patch.freezeCols ?? 0;
    const freezeChanged =
      this.freezeRowsBySheet.get(patch.viewport.sheetName) !== nextFreezeRows ||
      this.freezeColsBySheet.get(patch.viewport.sheetName) !== nextFreezeCols;
    this.freezeRowsBySheet.set(patch.viewport.sheetName, nextFreezeRows);
    this.freezeColsBySheet.set(patch.viewport.sheetName, nextFreezeCols);

    const changedKeys = new Set<string>();
    const changedStyleIds = new Set<string>();
    const damagedCellKeys = new Set<string>();
    const damage: { cell: CellItem }[] = [];
    patch.styles.forEach((style) => {
      const current = this.cellStyles.get(style.id);
      if (!current || cellStyleSignature(current) !== cellStyleSignature(style)) {
        changedStyleIds.add(style.id);
      }
      this.cellStyles.set(style.id, style);
    });
    if (patch.full) {
      const incomingKeys = new Set(
        patch.cells.map((cell) => `${patch.viewport.sheetName}!${cell.snapshot.address}`),
      );
      const sheetCellKeys = this.cellKeysBySheet.get(patch.viewport.sheetName);
      if (sheetCellKeys) {
        for (const key of sheetCellKeys) {
          if (incomingKeys.has(key)) {
            continue;
          }
          const snapshot = this.cellSnapshots.get(key);
          if (!snapshot || !isCellInsideViewport(snapshot, patch.viewport)) {
            continue;
          }
          this.cellSnapshots.delete(key);
          sheetCellKeys.delete(key);
          changedKeys.add(key);
          if (!damagedCellKeys.has(key)) {
            const parsed = parseCellAddress(snapshot.address, snapshot.sheetName);
            damage.push({ cell: [parsed.col, parsed.row] });
            damagedCellKeys.add(key);
          }
        }
      }
    }
    for (const cell of patch.cells) {
      const key = `${patch.viewport.sheetName}!${cell.snapshot.address}`;
      const current = this.cellSnapshots.get(key);
      if (current) {
        const incoming = cell.snapshot;
        if (shouldKeepCurrentSnapshot(current, incoming)) {
          continue;
        }
        if (cellSnapshotSignature(current) === cellSnapshotSignature(incoming)) {
          if (
            incoming.styleId &&
            changedStyleIds.has(incoming.styleId) &&
            !damagedCellKeys.has(key)
          ) {
            damage.push({ cell: [cell.col, cell.row] });
            damagedCellKeys.add(key);
          }
          continue;
        }
      }
      this.cellSnapshots.set(key, cell.snapshot);
      this.touchCellKey(key);
      this.sheetCellKeys(patch.viewport.sheetName).add(key);
      changedKeys.add(key);
      if (!damagedCellKeys.has(key)) {
        damage.push({ cell: [cell.col, cell.row] });
        damagedCellKeys.add(key);
      }
    }

    let axisChanged = false;
    if (patch.columns.length > 0) {
      const nextColumns = applyProjectedViewportAxisPatches({
        patches: patch.columns,
        sizes: this.columnSizesBySheet.get(patch.viewport.sheetName) ?? {},
        renderedSizes: this.columnWidthsBySheet.get(patch.viewport.sheetName) ?? {},
        pendingSizes: this.pendingColumnWidthsBySheet.get(patch.viewport.sheetName) ?? {},
        hiddenAxes: this.hiddenColumnsBySheet.get(patch.viewport.sheetName) ?? {},
      });
      this.columnSizesBySheet.set(patch.viewport.sheetName, nextColumns.sizes);
      this.columnWidthsBySheet.set(patch.viewport.sheetName, nextColumns.renderedSizes);
      this.pendingColumnWidthsBySheet.set(patch.viewport.sheetName, nextColumns.pendingSizes);
      if (Object.keys(nextColumns.hiddenAxes).length === 0) {
        this.hiddenColumnsBySheet.delete(patch.viewport.sheetName);
      } else {
        this.hiddenColumnsBySheet.set(patch.viewport.sheetName, nextColumns.hiddenAxes);
      }
      axisChanged = axisChanged || nextColumns.axisChanged;
    }

    if (patch.rows.length > 0) {
      const nextRows = applyProjectedViewportAxisPatches({
        patches: patch.rows,
        sizes: this.rowSizesBySheet.get(patch.viewport.sheetName) ?? {},
        renderedSizes: this.rowHeightsBySheet.get(patch.viewport.sheetName) ?? {},
        pendingSizes: this.pendingRowHeightsBySheet.get(patch.viewport.sheetName) ?? {},
        hiddenAxes: this.hiddenRowsBySheet.get(patch.viewport.sheetName) ?? {},
      });
      this.rowSizesBySheet.set(patch.viewport.sheetName, nextRows.sizes);
      this.rowHeightsBySheet.set(patch.viewport.sheetName, nextRows.renderedSizes);
      this.pendingRowHeightsBySheet.set(patch.viewport.sheetName, nextRows.pendingSizes);
      if (Object.keys(nextRows.hiddenAxes).length === 0) {
        this.hiddenRowsBySheet.delete(patch.viewport.sheetName);
      } else {
        this.hiddenRowsBySheet.set(patch.viewport.sheetName, nextRows.hiddenAxes);
      }
      axisChanged = axisChanged || nextRows.axisChanged;
    }

    this.pruneSheetCache(patch.viewport.sheetName);
    this.notifyCellSubscriptions(changedKeys);
    if (damage.length > 0 || axisChanged || freezeChanged) {
      this.listeners.forEach((listener) => listener());
    }
    return damage;
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

  private notifyCellSubscriptions(changedKeys: Set<string>): void {
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
