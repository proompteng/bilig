import type { GridEngineLike } from "@bilig/grid";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import { ValueTag, type CellSnapshot, type CellStyleRecord, type Viewport } from "@bilig/protocol";
import {
  decodeViewportPatch,
  type ViewportPatch,
  type WorkerEngineClient,
} from "@bilig/worker-transport";

const EMPTY_WIDTHS: Readonly<Record<number, number>> = Object.freeze({});
const DEFAULT_STYLE_ID = "style-0";
type CellItem = readonly [number, number];

interface CellSubscription {
  sheetName: string;
  addresses: Set<string>;
  listener: () => void;
}

export class WorkerViewportCache implements GridEngineLike {
  readonly workbook = {
    getSheet: (sheetName: string) => {
      if (!this.knownSheets.has(sheetName)) {
        return undefined;
      }
      const entries = [...this.cellSnapshots.values()].filter(
        (snapshot) => snapshot.sheetName === sheetName,
      );
      return {
        grid: {
          forEachCellEntry: (listener: (cellIndex: number, row: number, col: number) => void) => {
            entries.forEach((snapshot, index) => {
              const parsed = parseCellAddress(snapshot.address, snapshot.sheetName);
              listener(index, parsed.row, parsed.col);
            });
          },
        },
      };
    },
  };

  private readonly cellSnapshots = new Map<string, CellSnapshot>();
  private readonly cellStyles = new Map<string, CellStyleRecord>([
    [DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }],
  ]);
  private readonly cellSubscriptions = new Set<CellSubscription>();
  private readonly listeners = new Set<() => void>();
  private readonly columnWidthsBySheet = new Map<string, Record<number, number>>();
  private readonly rowHeightsBySheet = new Map<string, Record<number, number>>();
  private readonly knownSheets = new Set<string>();

  constructor(private readonly client: WorkerEngineClient) {}

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
    this.notifyCellSubscriptions(new Set([key]));
    this.listeners.forEach((listener) => listener());
  }

  setColumnWidth(sheetName: string, columnIndex: number, width: number): void {
    this.knownSheets.add(sheetName);
    const widths = { ...this.columnWidthsBySheet.get(sheetName) };
    widths[columnIndex] = width;
    this.columnWidthsBySheet.set(sheetName, widths);
    this.listeners.forEach((listener) => listener());
  }

  setKnownSheets(sheetNames: readonly string[]): void {
    this.knownSheets.clear();
    sheetNames.forEach((sheetName) => this.knownSheets.add(sheetName));
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
    return this.client.subscribeViewportPatches({ sheetName, ...viewport }, (bytes: Uint8Array) => {
      const damage = this.applyPatch(decodeViewportPatch(bytes));
      listener(damage);
    });
  }

  applyViewportPatch(patch: ViewportPatch): readonly { cell: CellItem }[] {
    return this.applyPatch(patch);
  }

  private applyPatch(patch: ReturnType<typeof decodeViewportPatch>): readonly { cell: CellItem }[] {
    this.knownSheets.add(patch.viewport.sheetName);
    if (patch.full) {
      this.clearViewportRegion(patch.viewport.sheetName, patch.viewport);
      this.clearAxisRange(
        this.columnWidthsBySheet,
        patch.viewport.sheetName,
        patch.viewport.colStart,
        patch.viewport.colEnd,
      );
      this.clearAxisRange(
        this.rowHeightsBySheet,
        patch.viewport.sheetName,
        patch.viewport.rowStart,
        patch.viewport.rowEnd,
      );
    }

    const changedKeys = new Set<string>();
    const damage: { cell: CellItem }[] = [];
    patch.styles.forEach((style) => {
      this.cellStyles.set(style.id, style);
    });
    for (const cell of patch.cells) {
      const key = `${patch.viewport.sheetName}!${cell.snapshot.address}`;
      this.cellSnapshots.set(key, cell.snapshot);
      changedKeys.add(key);
      damage.push({ cell: [cell.col, cell.row] });
    }

    if (patch.columns.length > 0) {
      const widths = { ...this.columnWidthsBySheet.get(patch.viewport.sheetName) };
      patch.columns.forEach((column: { index: number; size: number }) => {
        widths[column.index] = column.size;
      });
      this.columnWidthsBySheet.set(patch.viewport.sheetName, widths);
    }

    if (patch.rows.length > 0) {
      const heights = { ...this.rowHeightsBySheet.get(patch.viewport.sheetName) };
      patch.rows.forEach((row: { index: number; size: number }) => {
        heights[row.index] = row.size;
      });
      this.rowHeightsBySheet.set(patch.viewport.sheetName, heights);
    }

    this.notifyCellSubscriptions(changedKeys);
    this.listeners.forEach((listener) => listener());
    return damage;
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

  private clearViewportRegion(sheetName: string, viewport: Viewport): void {
    for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
      for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
        this.cellSnapshots.delete(`${sheetName}!${formatAddress(row, col)}`);
      }
    }
  }

  private clearAxisRange(
    store: Map<string, Record<number, number>>,
    sheetName: string,
    start: number,
    end: number,
  ): void {
    const current = store.get(sheetName);
    if (!current) {
      return;
    }
    const next = { ...current };
    for (let index = start; index <= end; index += 1) {
      delete next[index];
    }
    store.set(sheetName, next);
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
