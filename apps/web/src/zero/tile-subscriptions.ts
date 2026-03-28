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
  sourceCells: CellSourceRow[];
  cellEval: CellEvalRow[];
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
    return {
      getData: () =>
        detachments.reduce<TileData>(
          (aggregate, handle) => {
            aggregate.sourceCells.push(...handle.data.sourceCells);
            aggregate.cellEval.push(...handle.data.cellEval);
            aggregate.rowMetadata.push(...handle.data.rowMetadata);
            aggregate.columnMetadata.push(...handle.data.columnMetadata);
            aggregate.styleRanges.push(...handle.data.styleRanges);
            aggregate.formatRanges.push(...handle.data.formatRanges);
            return aggregate;
          },
          {
            sourceCells: [],
            cellEval: [],
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
      cellEval: [],
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
        data.sourceCells = value.map((row) => normalizeCellSourceRow(row));
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
        data.cellEval = value.map((row) => normalizeCellEvalRow(row));
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
        data.rowMetadata = value.map((row) => normalizeAxisMetadataRow(row));
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
        data.columnMetadata = value.map((row) => normalizeAxisMetadataRow(row));
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
      ) as unknown as TypedView<readonly FormatRangeRow[]>,
      (value) => {
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
