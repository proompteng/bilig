/* oxlint-disable @typescript-eslint/no-unsafe-type-assertion */
import type { TypedView, Zero } from "@rocicorp/zero";
import type { Viewport } from "@bilig/protocol";
import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from "@bilig/protocol";
import { queries } from "@bilig/zero-sync";
import type { AxisMetadataRow, CellSourceRow, CellEvalRow } from "./viewport-projector.js";
import type { CellStyleRecord } from "@bilig/protocol";

export const TILE_ROWS = 128;
export const TILE_COLS = 32;
const OVERSCAN_ROWS = 64;
const OVERSCAN_COLS = 16;
const PRELOAD_ROWS = 128;
const PRELOAD_COLS = 32;

interface TileDescriptor {
  key: string;
  sheetName: string;
  sheetViewId: string | undefined;
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
}

interface TileDataCounts {
  sourceCells: Map<string, number>;
  cellEval: Map<string, number>;
  rowMetadata: Map<string, number>;
  columnMetadata: Map<string, number>;
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

function clampViewport(viewport: Viewport, padRows: number, padCols: number): Viewport {
  return {
    rowStart: Math.max(0, viewport.rowStart - padRows),
    rowEnd: Math.max(viewport.rowStart, viewport.rowEnd + padRows),
    colStart: Math.max(0, viewport.colStart - padCols),
    colEnd: Math.max(viewport.colStart, viewport.colEnd + padCols),
  };
}

function toTileDescriptors(
  sheetName: string,
  viewport: Viewport,
  sheetViewId?: string,
): TileDescriptor[] {
  const tileRowStart = Math.floor(viewport.rowStart / TILE_ROWS);
  const tileRowEnd = Math.floor(viewport.rowEnd / TILE_ROWS);
  const tileColStart = Math.floor(viewport.colStart / TILE_COLS);
  const tileColEnd = Math.floor(viewport.colEnd / TILE_COLS);
  const descriptors: TileDescriptor[] = [];
  for (let tileRow = tileRowStart; tileRow <= tileRowEnd; tileRow += 1) {
    for (let tileCol = tileColStart; tileCol <= tileColEnd; tileCol += 1) {
      const rowStart = tileRow * TILE_ROWS;
      const colStart = tileCol * TILE_COLS;
      descriptors.push({
        key: `${sheetName}:${tileRow}:${tileCol}${sheetViewId ? `:${sheetViewId}` : ""}`,
        sheetName,
        sheetViewId,
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

function asTypedView<T>(view: unknown): TypedView<T> {
  return view as TypedView<T>;
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "number" ||
    typeof value === "string" ||
    typeof value === "boolean"
  );
}

function normalizeLiteralInput(value: unknown): LiteralInput | null | undefined {
  return isLiteralInput(value) ? value : undefined;
}

function normalizeValueTag(value: unknown): ValueTag | null {
  if (typeof value === "number") {
    return value in ValueTag ? (value as ValueTag) : null;
  }
  if (typeof value !== "string") {
    return null;
  }
  switch (value) {
    case "Empty":
      return ValueTag.Empty;
    case "Number":
      return ValueTag.Number;
    case "Boolean":
      return ValueTag.Boolean;
    case "String":
      return ValueTag.String;
    case "Error":
      return ValueTag.Error;
    default:
      return null;
  }
}

function normalizeErrorCode(value: unknown): ErrorCode {
  if (typeof value === "number") {
    return value as ErrorCode;
  }
  if (typeof value === "string" && value.length > 0) {
    const parsed = Number.parseInt(value, 10);
    if (Number.isFinite(parsed)) {
      return parsed as ErrorCode;
    }
    const enumValue = (ErrorCode as Record<string, unknown>)[value];
    if (typeof enumValue === "number") {
      return enumValue as ErrorCode;
    }
  }
  return ErrorCode.Value;
}

function normalizeCellValue(row: Record<string, unknown>): CellValue {
  if (typeof row["value"] === "object" && row["value"] !== null) {
    return row["value"] as CellValue;
  }

  const tag = normalizeValueTag(row["valueTag"]);
  switch (tag) {
    case null:
      return { tag: ValueTag.Empty };
    case ValueTag.Number:
      return typeof row["numberValue"] === "number"
        ? { tag: ValueTag.Number, value: row["numberValue"] }
        : { tag: ValueTag.Empty };
    case ValueTag.Boolean:
      return typeof row["booleanValue"] === "boolean"
        ? { tag: ValueTag.Boolean, value: row["booleanValue"] }
        : { tag: ValueTag.Empty };
    case ValueTag.String:
      return typeof row["stringValue"] === "string"
        ? { tag: ValueTag.String, value: row["stringValue"], stringId: 0 }
        : { tag: ValueTag.Empty };
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: normalizeErrorCode(row["errorCode"]) };
    case ValueTag.Empty:
    default:
      return { tag: ValueTag.Empty };
  }
}

function createTileData(): TileData {
  return {
    sourceCells: new Map(),
    cellEval: new Map(),
    rowMetadata: new Map(),
    columnMetadata: new Map(),
  };
}

function createTileDataCounts(): TileDataCounts {
  return {
    sourceCells: new Map(),
    cellEval: new Map(),
    rowMetadata: new Map(),
    columnMetadata: new Map(),
  };
}

function normalizeCellSourceRow(value: unknown): CellSourceRow {
  const row = asRecord(value);
  const normalized: CellSourceRow = {
    workbookId: String(row["workbookId"]),
    sheetName: String(row["sheetId"] ?? row["sheetName"]),
    address: String(row["address"]),
  };
  if (typeof row["rowNum"] === "number") {
    normalized.rowNum = row["rowNum"];
  }
  if (typeof row["colNum"] === "number") {
    normalized.colNum = row["colNum"];
  }
  const inputValue = normalizeLiteralInput(row["inputJson"] ?? row["inputValue"]);
  if (inputValue !== undefined) {
    normalized.inputValue = inputValue;
  }
  if (typeof row["formulaSource"] === "string") {
    normalized.formula = row["formulaSource"];
  } else if (typeof row["formula"] === "string") {
    normalized.formula = row["formula"];
  }
  if (typeof row["styleId"] === "string") {
    normalized.styleId = row["styleId"];
  }
  if (typeof row["formatId"] === "string") {
    normalized.explicitFormatId = row["formatId"];
  } else if (typeof row["explicitFormatId"] === "string") {
    normalized.explicitFormatId = row["explicitFormatId"];
  }
  if (typeof row["editorText"] === "string") {
    normalized.editorText = row["editorText"];
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
    sheetName: String(row["sheetId"] ?? row["sheetName"]),
    address: String(row["address"]),
    value: normalizeCellValue(row),
    flags: Number(row["flags"] ?? 0),
    version: Number(row["version"] ?? row["calcRevision"] ?? 0),
  };
  if (typeof row["rowNum"] === "number") {
    normalized.rowNum = row["rowNum"];
  }
  if (typeof row["colNum"] === "number") {
    normalized.colNum = row["colNum"];
  }
  if (typeof row["styleId"] === "string") {
    normalized.styleId = row["styleId"];
  }
  if (
    typeof row["styleJson"] === "object" &&
    row["styleJson"] !== null &&
    typeof asRecord(row["styleJson"])["id"] === "string"
  ) {
    normalized.styleJson = row["styleJson"] as CellStyleRecord;
  }
  if (typeof row["formatId"] === "string") {
    normalized.formatId = row["formatId"];
  }
  if (typeof row["formatCode"] === "string") {
    normalized.formatCode = row["formatCode"];
  }
  if (typeof row["calcRevision"] === "number") {
    normalized.calcRevision = row["calcRevision"];
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
    sheetName: String(row["sheetId"] ?? row["sheetName"]),
    startIndex: Number(row["rowNum"] ?? row["colNum"] ?? row["startIndex"] ?? 0),
    count: Number(row["count"] ?? 1),
  };
  if (typeof row["height"] === "number") {
    normalized.size = row["height"];
  }
  if (typeof row["width"] === "number") {
    normalized.size = row["width"];
  }
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

function cloneMap<T>(source: ReadonlyMap<string, T>): Map<string, T> {
  return new Map(source);
}

function cloneTileData(data: TileData): TileData {
  return {
    sourceCells: cloneMap(data.sourceCells),
    cellEval: cloneMap(data.cellEval),
    rowMetadata: cloneMap(data.rowMetadata),
    columnMetadata: cloneMap(data.columnMetadata),
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
}

export class TileSubscriptionManager {
  private readonly tiles = new Map<string, TileHandle>();
  private readonly preloadTiles = new Set<string>();

  constructor(
    private readonly zero: Zero,
    private readonly documentId: string,
    private readonly onError: (error: unknown) => void,
  ) {}

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: () => void,
    sheetViewId?: string,
  ): TileViewportAttachment {
    const clamped = clampViewport(viewport, OVERSCAN_ROWS, OVERSCAN_COLS);
    const descriptors = toTileDescriptors(sheetName, clamped, sheetViewId);
    const detachments = descriptors.map((descriptor) => this.attachTile(descriptor, listener));

    // Background preloading of surrounding tiles
    this.managePreloads(sheetName, viewport, sheetViewId);

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
            this.preloadTiles.delete(handle.descriptor.key);
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
    this.preloadTiles.clear();
  }

  private attachTile(descriptor: TileDescriptor, listener: (() => void) | null): TileHandle {
    const existing = this.tiles.get(descriptor.key);
    if (existing) {
      existing.refCount += 1;
      if (listener) {
        existing.listeners.add(listener);
      }
      return existing;
    }

    const data = createTileData();
    const listeners = new Set(listener ? [listener] : []);
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
    const pushView = <T>(
      view: TypedView<readonly unknown[]>,
      map: Map<string, T>,
      normalize: (value: unknown) => T,
      keyFn: (value: T) => string,
    ) => {
      const update = (rows: readonly unknown[]) => {
        const nextKeys = new Set<string>();
        for (const row of rows) {
          const normalized = normalize(row);
          const key = keyFn(normalized);
          nextKeys.add(key);
          map.set(key, normalized);
        }
        for (const key of map.keys()) {
          if (!nextKeys.has(key)) {
            map.delete(key);
          }
        }
      };

      // Ensure data is synced synchronously
      update(view.data);

      destroyers.push(
        bindView(view, (value) => {
          update(value);
          notifyListeners();
        }),
      );
    };

    pushView(
      asTypedView<readonly unknown[]>(
        this.zero.materialize(
          queries.cellInput.tile({
            documentId: this.documentId,
            sheetName: descriptor.sheetName,
            rowStart: descriptor.rowStart,
            rowEnd: descriptor.rowEnd,
            colStart: descriptor.colStart,
            colEnd: descriptor.colEnd,
          }),
        ),
      ),
      data.sourceCells,
      normalizeCellSourceRow,
      normalizeCellSourceKey,
    );
    pushView(
      asTypedView<readonly unknown[]>(
        this.zero.materialize(
          queries.cellEval.tile({
            documentId: this.documentId,
            sheetName: descriptor.sheetName,
            rowStart: descriptor.rowStart,
            rowEnd: descriptor.rowEnd,
            colStart: descriptor.colStart,
            colEnd: descriptor.colEnd,
          }),
        ),
      ),
      data.cellEval,
      normalizeCellEvalRow,
      normalizeCellEvalKey,
    );
    pushView(
      asTypedView<readonly unknown[]>(
        this.zero.materialize(
          queries.sheetRow.tile({
            documentId: this.documentId,
            sheetName: descriptor.sheetName,
            rowStart: descriptor.rowStart,
            rowEnd: descriptor.rowEnd,
          }),
        ),
      ),
      data.rowMetadata,
      normalizeAxisMetadataRow,
      normalizeAxisMetadataKey,
    );
    pushView(
      asTypedView<readonly unknown[]>(
        this.zero.materialize(
          queries.sheetCol.tile({
            documentId: this.documentId,
            sheetName: descriptor.sheetName,
            colStart: descriptor.colStart,
            colEnd: descriptor.colEnd,
          }),
        ),
      ),
      data.columnMetadata,
      normalizeAxisMetadataRow,
      normalizeAxisMetadataKey,
    );
    this.tiles.set(descriptor.key, handle);
    return handle;
  }

  private managePreloads(sheetName: string, viewport: Viewport, sheetViewId?: string) {
    const preloadViewport = clampViewport(viewport, PRELOAD_ROWS, PRELOAD_COLS);
    const activeDescriptors = new Set(
      toTileDescriptors(
        sheetName,
        clampViewport(viewport, OVERSCAN_ROWS, OVERSCAN_COLS),
        sheetViewId,
      ).map((d) => d.key),
    );
    const preloadDescriptors = toTileDescriptors(sheetName, preloadViewport, sheetViewId);

    for (const descriptor of preloadDescriptors) {
      if (!activeDescriptors.has(descriptor.key) && !this.tiles.has(descriptor.key)) {
        this.attachTile(descriptor, null);
        this.preloadTiles.add(descriptor.key);
      }
    }

    // Evict distant preloads if needed
    for (const key of this.preloadTiles) {
      if (!activeDescriptors.has(key)) {
        const isStillInPreloadRange = preloadDescriptors.some((d) => d.key === key);
        if (!isStillInPreloadRange) {
          const handle = this.tiles.get(key);
          if (handle) {
            handle.refCount -= 1;
            if (handle.refCount <= 0) {
              handle.destroy();
              this.tiles.delete(key);
              this.preloadTiles.delete(key);
            }
          }
        }
      }
    }
  }
}
