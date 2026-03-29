/* oxlint-disable @typescript-eslint/no-unsafe-type-assertion */
import type { TypedView, Zero } from "@rocicorp/zero";
import type { Viewport } from "@bilig/protocol";
import { queries } from "@bilig/zero-sync";
import type {
  AxisMetadataRow,
  CellSourceRow,
  CellEvalRow,
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
  sourceCells: Map<string, CellSourceRow>;
  cellEval: Map<string, CellEvalRow>;
  rowMetadata: Map<string, AxisMetadataRow>;
  columnMetadata: Map<string, AxisMetadataRow>;
  styleRanges: Map<string, StyleRangeRow>;
  formatRanges: Map<string, FormatRangeRow>;
}

interface TileDataCounts {
  sourceCells: Map<string, number>;
  cellEval: Map<string, number>;
  rowMetadata: Map<string, number>;
  columnMetadata: Map<string, number>;
  styleRanges: Map<string, number>;
  formatRanges: Map<string, number>;
}

interface TileHandle {
  descriptor: TileDescriptor;
  refCount: number;
  version: number;
  data: TileData;
  listeners: Set<() => void>;
  destroy(): void;
}

export interface TileViewportAttachment {
  getData(): TileData;
  getSourceCell(address: string): CellSourceRow | undefined;
  dispose(): void;
}

function bindView<T>(view: TypedView<T>, onData: (data: T) => void): () => void {
  const unsubscribe = view.addListener((data) => {
    onData(data as T);
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

function asRecord(value: unknown): Record<string, unknown> {
  return value as Record<string, unknown>;
}

function createTileData(): TileData {
  return {
    sourceCells: new Map(),
    cellEval: new Map(),
    rowMetadata: new Map(),
    columnMetadata: new Map(),
    styleRanges: new Map(),
    formatRanges: new Map(),
  };
}

function createTileDataCounts(): TileDataCounts {
  return {
    sourceCells: new Map(),
    cellEval: new Map(),
    rowMetadata: new Map(),
    columnMetadata: new Map(),
    styleRanges: new Map(),
    formatRanges: new Map(),
  };
}

function normalizeCellSourceRow(value: unknown): CellSourceRow {
  const row = asRecord(value);
  const normalized: CellSourceRow = {
    workbookId: String(row["workbookId"]),
    sheetName: String(row["sheetName"]),
    address: String(row["address"]),
    inputValue: row["inputValue"],
  };
  if (typeof row["rowNum"] === "number") {
    normalized.rowNum = row["rowNum"];
  }
  if (typeof row["colNum"] === "number") {
    normalized.colNum = row["colNum"];
  }
  if (typeof row["formula"] === "string") {
    normalized.formula = row["formula"];
  }
  if (typeof row["format"] === "string") {
    normalized.format = row["format"];
  }
  if (typeof row["explicitFormatId"] === "string") {
    normalized.explicitFormatId = row["explicitFormatId"];
  }
  return normalized;
}

function normalizeCellSourceKey(row: CellSourceRow): string {
  return row.address;
}

function normalizeCellEvalRow(value: unknown): CellEvalRow {
  const row = asRecord(value);
  const normalized: CellEvalRow = {
    workbookId: String(row["workbookId"]),
    sheetName: String(row["sheetName"]),
    address: String(row["address"]),
    value: row["value"] as CellEvalRow["value"],
    flags: Number(row["flags"] ?? 0),
    version: Number(row["version"] ?? 0),
  };
  if (typeof row["rowNum"] === "number") {
    normalized.rowNum = row["rowNum"];
  }
  if (typeof row["colNum"] === "number") {
    normalized.colNum = row["colNum"];
  }
  return normalized;
}

function normalizeCellEvalKey(row: CellEvalRow): string {
  return row.address;
}

function normalizeAxisMetadataRow(value: unknown): AxisMetadataRow {
  const row = asRecord(value);
  const normalized: AxisMetadataRow = {
    workbookId: String(row["workbookId"]),
    sheetName: String(row["sheetName"]),
    startIndex: Number(row["startIndex"] ?? 0),
    count: Number(row["count"] ?? 0),
  };
  if (typeof row["size"] === "number") {
    normalized.size = row["size"];
  }
  if (typeof row["hidden"] === "boolean") {
    normalized.hidden = row["hidden"];
  }
  return normalized;
}

function normalizeAxisMetadataKey(row: AxisMetadataRow): string {
  return [row.startIndex, row.count, row.size ?? "", row.hidden === true ? 1 : 0].join(":");
}

function replaceMap<T>(
  target: Map<string, T>,
  rows: readonly unknown[],
  normalize: (value: unknown) => T,
  keyFn: (value: T) => string,
): void {
  target.clear();
  for (const row of rows) {
    const normalized = normalize(row);
    target.set(keyFn(normalized), normalized);
  }
}

function cloneMap<T>(source: ReadonlyMap<string, T>): Map<string, T> {
  return new Map(source);
}

function cloneTileData(data: TileData): TileData {
  return {
    sourceCells: cloneMap(data.sourceCells),
    cellEval: cloneMap(data.cellEval),
    rowMetadata: cloneMap(data.rowMetadata),
    columnMetadata: cloneMap(data.columnMetadata),
    styleRanges: cloneMap(data.styleRanges),
    formatRanges: cloneMap(data.formatRanges),
  };
}

function updateAggregateMap<T>(
  aggregate: Map<string, T>,
  counts: Map<string, number>,
  previous: ReadonlyMap<string, T>,
  next: ReadonlyMap<string, T>,
): void {
  for (const key of previous.keys()) {
    if (next.has(key)) {
      continue;
    }
    const remaining = (counts.get(key) ?? 0) - 1;
    if (remaining <= 0) {
      counts.delete(key);
      aggregate.delete(key);
      continue;
    }
    counts.set(key, remaining);
  }

  for (const [key, value] of next) {
    if (!previous.has(key)) {
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    aggregate.set(key, value);
  }
}

function updateAggregateTileData(
  aggregate: TileData,
  counts: TileDataCounts,
  previous: TileData,
  next: TileData,
): void {
  updateAggregateMap(
    aggregate.sourceCells,
    counts.sourceCells,
    previous.sourceCells,
    next.sourceCells,
  );
  updateAggregateMap(aggregate.cellEval, counts.cellEval, previous.cellEval, next.cellEval);
  updateAggregateMap(
    aggregate.rowMetadata,
    counts.rowMetadata,
    previous.rowMetadata,
    next.rowMetadata,
  );
  updateAggregateMap(
    aggregate.columnMetadata,
    counts.columnMetadata,
    previous.columnMetadata,
    next.columnMetadata,
  );
  updateAggregateMap(
    aggregate.styleRanges,
    counts.styleRanges,
    previous.styleRanges,
    next.styleRanges,
  );
  updateAggregateMap(
    aggregate.formatRanges,
    counts.formatRanges,
    previous.formatRanges,
    next.formatRanges,
  );
}

export class TileSubscriptionManager {
  private readonly tiles = new Map<string, TileHandle>();

  constructor(
    private readonly zero: Zero,
    private readonly documentId: string,
    private readonly onError: (error: unknown) => void,
  ) {}

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: () => void,
  ): TileViewportAttachment {
    const descriptors = toTileDescriptors(sheetName, viewport);
    const detachments = descriptors.map((descriptor) => this.attachTile(descriptor, listener));
    const aggregateData = createTileData();
    const aggregateCounts = createTileDataCounts();
    const snapshots = new Map<string, { version: number; data: TileData }>();

    const refreshAggregate = () => {
      for (const handle of detachments) {
        const previous = snapshots.get(handle.descriptor.key);
        if (previous && previous.version === handle.version) {
          continue;
        }
        const next = cloneTileData(handle.data);
        updateAggregateTileData(
          aggregateData,
          aggregateCounts,
          previous?.data ?? createTileData(),
          next,
        );
        snapshots.set(handle.descriptor.key, {
          version: handle.version,
          data: next,
        });
      }
    };

    const getData = () => {
      refreshAggregate();
      return aggregateData;
    };

    return {
      getData,
      getSourceCell: (address: string) => getData().sourceCells.get(address),
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

    const data = createTileData();
    const listeners = new Set([listener]);
    const handle: TileHandle = {
      descriptor,
      refCount: 1,
      version: 0,
      data,
      listeners,
      destroy: () => {
        while (destroyers.length > 0) {
          destroyers.pop()?.();
        }
      },
    };

    const notifyListeners = () => {
      handle.version += 1;
      for (const nextListener of listeners) {
        try {
          nextListener();
        } catch (error) {
          this.onError(error);
        }
      }
    };

    const destroyers: Array<() => void> = [];
    const pushView = <T>(view: TypedView<T>, assign: (value: T) => void) => {
      assign(view.data);
      destroyers.push(
        bindView(view, (value) => {
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
      ) as unknown as TypedView<readonly unknown[]>,
      (value) => {
        replaceMap(data.sourceCells, value, normalizeCellSourceRow, normalizeCellSourceKey);
      },
    );
    pushView(
      this.zero.materialize(
        queries.cellEval.tile({
          documentId: this.documentId,
          sheetName: descriptor.sheetName,
          rowStart: descriptor.rowStart,
          rowEnd: descriptor.rowEnd,
          colStart: descriptor.colStart,
          colEnd: descriptor.colEnd,
        }),
      ) as unknown as TypedView<readonly unknown[]>,
      (value) => {
        replaceMap(data.cellEval, value, normalizeCellEvalRow, normalizeCellEvalKey);
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
      ) as unknown as TypedView<readonly unknown[]>,
      (value) => {
        replaceMap(data.rowMetadata, value, normalizeAxisMetadataRow, normalizeAxisMetadataKey);
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
      ) as unknown as TypedView<readonly unknown[]>,
      (value) => {
        replaceMap(data.columnMetadata, value, normalizeAxisMetadataRow, normalizeAxisMetadataKey);
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
      ) as unknown as TypedView<readonly StyleRangeRow[]>,
      (value) => {
        replaceMap(
          data.styleRanges,
          value,
          (row) => row as StyleRangeRow,
          (row) => row.id,
        );
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
      ) as unknown as TypedView<readonly FormatRangeRow[]>,
      (value) => {
        replaceMap(
          data.formatRanges,
          value,
          (row) => row as FormatRangeRow,
          (row) => row.id,
        );
      },
    );
    this.tiles.set(descriptor.key, handle);
    return handle;
  }
}
