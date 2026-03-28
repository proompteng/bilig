/* oxlint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-type-assertion */
import type { Zero } from "@rocicorp/zero";
import type { Viewport } from "@bilig/protocol";
import { queries } from "@bilig/zero-sync";
import type {
  AxisMetadataRow,
  CellSourceRow,
  ComputedCellRow,
  FormatRangeRow,
  StyleRangeRow,
} from "./viewport-projector.js";

export const TILE_ROWS = 128;
export const TILE_COLS = 32;
const OVERSCAN_ROWS = 32;
const OVERSCAN_COLS = 8;

interface TileDescriptor {
  key: string;
  sheetName: string;
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

interface TileData {
  sourceCells: CellSourceRow[];
  computedCells: ComputedCellRow[];
  rowMetadata: AxisMetadataRow[];
  columnMetadata: AxisMetadataRow[];
  styleRanges: StyleRangeRow[];
  formatRanges: FormatRangeRow[];
}

interface TileHandle {
  descriptor: TileDescriptor;
  refCount: number;
  data: TileData;
  listeners: Set<() => void>;
  destroy(): void;
}

export interface TileViewportAttachment {
  getData(): TileData;
  dispose(): void;
}

function bindView(view: any, onData: (data: any) => void): () => void {
  const unsubscribe = view.addListener((data: any) => {
    onData(data);
  });
  return () => {
    unsubscribe();
    view.destroy();
  };
}

function clampViewport(viewport: Viewport): Viewport {
  return {
    rowStart: Math.max(0, viewport.rowStart - OVERSCAN_ROWS),
    rowEnd: Math.max(viewport.rowStart, viewport.rowEnd + OVERSCAN_ROWS),
    colStart: Math.max(0, viewport.colStart - OVERSCAN_COLS),
    colEnd: Math.max(viewport.colStart, viewport.colEnd + OVERSCAN_COLS),
  };
}

function toTileDescriptors(sheetName: string, viewport: Viewport): TileDescriptor[] {
  const expanded = clampViewport(viewport);
  const tileRowStart = Math.floor(expanded.rowStart / TILE_ROWS);
  const tileRowEnd = Math.floor(expanded.rowEnd / TILE_ROWS);
  const tileColStart = Math.floor(expanded.colStart / TILE_COLS);
  const tileColEnd = Math.floor(expanded.colEnd / TILE_COLS);
  const descriptors: TileDescriptor[] = [];
  for (let tileRow = tileRowStart; tileRow <= tileRowEnd; tileRow += 1) {
    for (let tileCol = tileColStart; tileCol <= tileColEnd; tileCol += 1) {
      const rowStart = tileRow * TILE_ROWS;
      const colStart = tileCol * TILE_COLS;
      descriptors.push({
        key: `${sheetName}:${tileRow}:${tileCol}`,
        sheetName,
        rowStart,
        rowEnd: rowStart + TILE_ROWS - 1,
        colStart,
        colEnd: colStart + TILE_COLS - 1,
      });
    }
  }
  return descriptors;
}

export class TileSubscriptionManager {
  private readonly tiles = new Map<string, TileHandle>();

  constructor(
    private readonly zero: Zero,
    private readonly documentId: string,
    private readonly onError: (error: unknown) => void,
  ) {}

  subscribeViewport(sheetName: string, viewport: Viewport, listener: () => void): TileViewportAttachment {
    const descriptors = toTileDescriptors(sheetName, viewport);
    const detachments = descriptors.map((descriptor) => this.attachTile(descriptor, listener));
    return {
      getData: () =>
        detachments.reduce<TileData>(
          (aggregate, handle) => {
            const data = handle.data;
            aggregate.sourceCells.push(...data.sourceCells);
            aggregate.computedCells.push(...data.computedCells);
            aggregate.rowMetadata.push(...data.rowMetadata);
            aggregate.columnMetadata.push(...data.columnMetadata);
            aggregate.styleRanges.push(...data.styleRanges);
            aggregate.formatRanges.push(...data.formatRanges);
            return aggregate;
          },
          {
            sourceCells: [],
            computedCells: [],
            rowMetadata: [],
            columnMetadata: [],
            styleRanges: [],
            formatRanges: [],
          },
        ),
      dispose: () => {
        for (const handle of detachments) {
          handle.listeners.delete(listener);
          handle.refCount -= 1;
          if (handle.refCount <= 0) {
            handle.destroy();
            this.tiles.delete(handle.descriptor.key);
          }
        }
      },
    };
  }

  dispose(): void {
    for (const handle of this.tiles.values()) {
      handle.destroy();
    }
    this.tiles.clear();
  }

  private attachTile(descriptor: TileDescriptor, listener: () => void): TileHandle {
    const existing = this.tiles.get(descriptor.key);
    if (existing) {
      existing.refCount += 1;
      existing.listeners.add(listener);
      return existing;
    }

    const data: TileData = {
      sourceCells: [],
      computedCells: [],
      rowMetadata: [],
      columnMetadata: [],
      styleRanges: [],
      formatRanges: [],
    };
    const listeners = new Set([listener]);

    const notifyListeners = () => {
      for (const nextListener of listeners) {
        try {
          nextListener();
        } catch (error) {
          this.onError(error);
        }
      }
    };

    const destroyers: Array<() => void> = [];
    const pushView = (view: any, assign: (value: any) => void) => {
      assign(view.data);
      destroyers.push(
        bindView(view, (value: any) => {
          assign(value);
          notifyListeners();
        }),
      );
    };

    pushView(
      this.zero.materialize(
        queries.cells.tile({
          documentId: this.documentId,
          sheetName: descriptor.sheetName,
          rowStart: descriptor.rowStart,
          rowEnd: descriptor.rowEnd,
          colStart: descriptor.colStart,
          colEnd: descriptor.colEnd,
        }),
      ),
      (value: any) => {
        data.sourceCells = [...value];
      },
    );
    pushView(
      this.zero.materialize(
        queries.computedCells.tile({
          documentId: this.documentId,
          sheetName: descriptor.sheetName,
          rowStart: descriptor.rowStart,
          rowEnd: descriptor.rowEnd,
          colStart: descriptor.colStart,
          colEnd: descriptor.colEnd,
        }),
      ),
      (value: any) => {
        data.computedCells = [...value];
      },
    );
    pushView(
      this.zero.materialize(
        queries.rowMetadata.tile({
          documentId: this.documentId,
          sheetName: descriptor.sheetName,
          rowStart: descriptor.rowStart,
          rowEnd: descriptor.rowEnd,
        }),
      ),
      (value: any) => {
        data.rowMetadata = [...value];
      },
    );
    pushView(
      this.zero.materialize(
        queries.columnMetadata.tile({
          documentId: this.documentId,
          sheetName: descriptor.sheetName,
          colStart: descriptor.colStart,
          colEnd: descriptor.colEnd,
        }),
      ),
      (value: any) => {
        data.columnMetadata = [...value];
      },
    );
    pushView(
      this.zero.materialize(
        queries.styleRanges.intersectTile({
          documentId: this.documentId,
          sheetName: descriptor.sheetName,
          rowStart: descriptor.rowStart,
          rowEnd: descriptor.rowEnd,
          colStart: descriptor.colStart,
          colEnd: descriptor.colEnd,
        }),
      ),
      (value: any) => {
        data.styleRanges = [...value];
      },
    );
    pushView(
      this.zero.materialize(
        queries.formatRanges.intersectTile({
          documentId: this.documentId,
          sheetName: descriptor.sheetName,
          rowStart: descriptor.rowStart,
          rowEnd: descriptor.rowEnd,
          colStart: descriptor.colStart,
          colEnd: descriptor.colEnd,
        }),
      ),
      (value: any) => {
        data.formatRanges = [...value];
      },
    );

    const handle: TileHandle = {
      descriptor,
      refCount: 1,
      data,
      listeners,
      destroy: () => {
        while (destroyers.length > 0) {
          destroyers.pop()?.();
        }
      },
    };
    this.tiles.set(descriptor.key, handle);
    return handle;
  }
}
