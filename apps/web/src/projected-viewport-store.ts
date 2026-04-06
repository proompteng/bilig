import { parseCellAddress } from "@bilig/formula";
import type { GridEngineLike } from "@bilig/grid";
import { ValueTag, type CellSnapshot, type CellStyleRecord, type Viewport } from "@bilig/protocol";
import {
  decodeViewportPatch,
  type ViewportPatch,
  type WorkerEngineClient,
} from "@bilig/worker-transport";

const EMPTY_WIDTHS: Readonly<Record<number, number>> = Object.freeze({});
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
  private readonly columnWidthsBySheet = new Map<string, Record<number, number>>();
  private readonly pendingColumnWidthsBySheet = new Map<string, Record<number, number>>();
  private readonly rowHeightsBySheet = new Map<string, Record<number, number>>();
  private readonly knownSheets = new Set<string>();
  private readonly activeViewportKeysBySheet = new Map<string, Set<string>>();
  private readonly activeViewports = new Map<string, Viewport>();

  constructor(private readonly client?: WorkerEngineClient) {}

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  peekCell(sheetName: string, address: string): CellSnapshot | undefined {
    return this.cellSnapshots.get(`${sheetName}!${address}`);
  }

  getColumnWidths(sheetName: string): Readonly<Record<number, number>> {
    return this.columnWidthsBySheet.get(sheetName) ?? EMPTY_WIDTHS;
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
    this.knownSheets.add(snapshot.sheetName);
    this.cellSnapshots.set(key, snapshot);
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
    const widths = { ...this.columnWidthsBySheet.get(sheetName) };
    widths[columnIndex] = width;
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
    const widths = { ...this.columnWidthsBySheet.get(sheetName) };
    if (width === undefined) {
      delete widths[columnIndex];
    } else {
      widths[columnIndex] = width;
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
    };
  }

  applyViewportPatch(patch: ViewportPatch): readonly { cell: CellItem }[] {
    return this.applyPatch(patch);
  }

  private applyPatch(patch: ReturnType<typeof decodeViewportPatch>): readonly { cell: CellItem }[] {
    this.knownSheets.add(patch.viewport.sheetName);

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
      this.sheetCellKeys(patch.viewport.sheetName).add(key);
      changedKeys.add(key);
      if (!damagedCellKeys.has(key)) {
        damage.push({ cell: [cell.col, cell.row] });
        damagedCellKeys.add(key);
      }
    }

    let axisChanged = false;
    if (patch.columns.length > 0) {
      const widths = { ...this.columnWidthsBySheet.get(patch.viewport.sheetName) };
      const pendingWidths = { ...this.pendingColumnWidthsBySheet.get(patch.viewport.sheetName) };
      patch.columns.forEach((column: { index: number; size: number }) => {
        const pending = pendingWidths[column.index];
        if (pending !== undefined && pending !== column.size) {
          return;
        }
        if (pending === column.size) {
          delete pendingWidths[column.index];
        }
        widths[column.index] = column.size;
        axisChanged = true;
      });
      this.columnWidthsBySheet.set(patch.viewport.sheetName, widths);
      this.pendingColumnWidthsBySheet.set(patch.viewport.sheetName, pendingWidths);
    }

    if (patch.rows.length > 0) {
      const heights = { ...this.rowHeightsBySheet.get(patch.viewport.sheetName) };
      patch.rows.forEach((row: { index: number; size: number }) => {
        heights[row.index] = row.size;
        axisChanged = true;
      });
      this.rowHeightsBySheet.set(patch.viewport.sheetName, heights);
    }

    this.pruneSheetCache(patch.viewport.sheetName);
    this.notifyCellSubscriptions(changedKeys);
    if (damage.length > 0 || axisChanged) {
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
    if (!activeViewportKeys || activeViewportKeys.size === 0) {
      return;
    }
    const activeViewports = [...activeViewportKeys]
      .map((key) => this.activeViewports.get(key))
      .filter((viewport): viewport is Viewport => viewport !== undefined);
    const pinnedKeys = new Set<string>();
    this.cellSubscriptions.forEach((subscription) => {
      if (subscription.sheetName !== sheetName) {
        return;
      }
      subscription.addresses.forEach((address) => pinnedKeys.add(`${sheetName}!${address}`));
    });
    const keysToInspect = Array.from(sheetCellKeys);
    for (const key of keysToInspect) {
      if (sheetCellKeys.size <= MAX_CACHED_CELLS_PER_SHEET) {
        break;
      }
      if (pinnedKeys.has(key)) {
        continue;
      }
      const snapshot = this.cellSnapshots.get(key);
      if (!snapshot) {
        sheetCellKeys.delete(key);
        continue;
      }
      const parsed = parseCellAddress(snapshot.address, snapshot.sheetName);
      const insideActiveViewport = activeViewports.some((viewport) => {
        return (
          parsed.row >= viewport.rowStart &&
          parsed.row <= viewport.rowEnd &&
          parsed.col >= viewport.colStart &&
          parsed.col <= viewport.colEnd
        );
      });
      if (insideActiveViewport) {
        continue;
      }
      this.cellSnapshots.delete(key);
      sheetCellKeys.delete(key);
    }
  }

  private dropSheetCache(sheetName: string): void {
    this.cellKeysBySheet.get(sheetName)?.forEach((key) => {
      this.cellSnapshots.delete(key);
    });
    this.cellKeysBySheet.delete(sheetName);
    this.columnWidthsBySheet.delete(sheetName);
    this.pendingColumnWidthsBySheet.delete(sheetName);
    this.rowHeightsBySheet.delete(sheetName);
    const viewportKeys = this.activeViewportKeysBySheet.get(sheetName);
    viewportKeys?.forEach((key) => this.activeViewports.delete(key));
    this.activeViewportKeysBySheet.delete(sheetName);
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
