import { formatAddress } from "@bilig/formula";
import {
  ValueTag,
  formatCellDisplayValue,
  type CellSnapshot,
  type CellStyleRecord,
  type CellValue,
  type LiteralInput,
  type RecalcMetrics,
  type Viewport,
} from "@bilig/protocol";
import type { ViewportPatch, ViewportPatchedCell } from "@bilig/worker-transport";

const DEFAULT_STYLE_ID = "style-0";
const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_ROW_HEIGHT = 22;

export interface WorkbookRow {
  id: string;
  name: string;
  headRevision: number;
  calculatedRevision: number;
}

export interface SheetRow {
  workbookId: string;
  name: string;
  sortOrder: number;
}

export interface CellSourceRow {
  workbookId: string;
  sheetName: string;
  address: string;
  rowNum?: number | null | undefined;
  colNum?: number | null | undefined;
  inputValue?: unknown;
  formula?: string | null | undefined;
  format?: string;
  explicitFormatId?: string;
}

export interface CellEvalRow {
  workbookId: string;
  sheetName: string;
  address: string;
  rowNum?: number | null | undefined;
  colNum?: number | null | undefined;
  value: CellValue;
  flags: number;
  version: number;
}

export interface AxisMetadataRow {
  workbookId: string;
  sheetName: string;
  startIndex: number;
  count: number;
  size?: number | null | undefined;
  hidden?: boolean | null | undefined;
}

export interface StyleRangeRow {
  id: string;
  workbookId: string;
  sheetName: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  styleId: string;
  updatedAt?: number;
}

export interface FormatRangeRow {
  id: string;
  workbookId: string;
  sheetName: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  formatId: string;
  updatedAt?: number;
}

export interface StyleRow {
  workbookId: string;
  id: string;
  recordJSON: CellStyleRecord;
}

export interface NumberFormatRow {
  workbookId: string;
  id: string;
  code: string;
  kind: string;
}

export interface ViewportProjectionState {
  nextVersion: number;
  knownStyleIds: Set<string>;
  lastCellSignatures: Map<string, string>;
  lastColumnSignatures: Map<number, string>;
  lastRowSignatures: Map<number, string>;
}

export interface ViewportProjectionInput {
  viewport: Viewport & { sheetName: string };
  metrics?: RecalcMetrics;
  sourceCells: readonly CellSourceRow[];
  cellEval: readonly CellEvalRow[];
  rowMetadata: readonly AxisMetadataRow[];
  columnMetadata: readonly AxisMetadataRow[];
  styleRanges: readonly StyleRangeRow[];
  formatRanges: readonly FormatRangeRow[];
  stylesById: ReadonlyMap<string, CellStyleRecord>;
  numberFormatCodeById: ReadonlyMap<string, string>;
}

interface AxisViewportEntry {
  index: number;
  size?: number | null;
  hidden?: boolean | null;
}

interface ResolvedNumberFormat {
  code?: string;
  numberFormatId?: string;
}

const EMPTY_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
};

function isLiteralInput(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "number" ||
    typeof value === "string" ||
    typeof value === "boolean"
  );
}

function emptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 0,
  };
}

function compareRectRanges(
  left: Pick<StyleRangeRow, "startRow" | "startCol" | "endRow" | "endCol" | "id" | "updatedAt">,
  right: Pick<StyleRangeRow, "startRow" | "startCol" | "endRow" | "endCol" | "id" | "updatedAt">,
): number {
  return (
    left.startRow - right.startRow ||
    left.startCol - right.startCol ||
    left.endRow - right.endRow ||
    left.endCol - right.endCol ||
    (left.updatedAt ?? 0) - (right.updatedAt ?? 0) ||
    left.id.localeCompare(right.id)
  );
}

function pointInRange(
  row: number,
  col: number,
  range: Pick<StyleRangeRow, "startRow" | "endRow" | "startCol" | "endCol">,
): boolean {
  return (
    row >= range.startRow && row <= range.endRow && col >= range.startCol && col <= range.endCol
  );
}

function resolveStyleId(row: number, col: number, styleRanges: readonly StyleRangeRow[]): string {
  let styleId = DEFAULT_STYLE_ID;
  for (const range of styleRanges) {
    if (pointInRange(row, col, range)) {
      styleId = range.styleId;
    }
  }
  return styleId;
}

function resolveNumberFormat(
  row: number,
  col: number,
  source: CellSourceRow | undefined,
  formatRanges: readonly FormatRangeRow[],
  numberFormatCodeById: ReadonlyMap<string, string>,
): ResolvedNumberFormat {
  const result: ResolvedNumberFormat = {};
  const explicitFormatId = source?.explicitFormatId;
  const sourceFormat = source?.format;
  if (sourceFormat !== undefined) {
    result.code = sourceFormat;
    if (explicitFormatId) {
      result.numberFormatId = explicitFormatId;
    }
    return result;
  }
  if (explicitFormatId) {
    const code = numberFormatCodeById.get(explicitFormatId);
    if (code !== undefined) {
      result.code = code;
    }
    result.numberFormatId = explicitFormatId;
    return result;
  }

  let numberFormatId: string | undefined;
  for (const range of formatRanges) {
    if (pointInRange(row, col, range)) {
      numberFormatId = range.formatId;
    }
  }
  if (!numberFormatId) {
    return result;
  }
  const code = numberFormatCodeById.get(numberFormatId);
  if (code !== undefined) {
    result.code = code;
  }
  result.numberFormatId = numberFormatId;
  return result;
}

function toEditorText(snapshot: CellSnapshot): string {
  if (snapshot.formula) {
    return `=${snapshot.formula}`;
  }
  if (snapshot.input === null || snapshot.input === undefined) {
    return formatCellDisplayValue(snapshot.value, snapshot.format);
  }
  if (typeof snapshot.input === "boolean") {
    return snapshot.input ? "TRUE" : "FALSE";
  }
  return String(snapshot.input);
}

function buildPatchedCell(
  snapshot: CellSnapshot,
  row: number,
  col: number,
  styleId: string,
): ViewportPatchedCell {
  const editorText = toEditorText(snapshot);
  const displayText = formatCellDisplayValue(snapshot.value, snapshot.format);
  return {
    row,
    col,
    snapshot,
    displayText,
    copyText: snapshot.formula ? editorText : displayText,
    editorText,
    formatId: 0,
    styleId,
  };
}

function buildAxisEntries(
  viewport: Viewport,
  rows: readonly AxisMetadataRow[],
  defaultSize: number,
  startKey: "rowStart" | "colStart",
  endKey: "rowEnd" | "colEnd",
): Map<number, AxisViewportEntry> {
  const start = viewport[startKey];
  const end = viewport[endKey];
  const indexed = new Map<number, AxisViewportEntry>();
  for (const entry of rows) {
    const entryEnd = entry.startIndex + Math.max(0, entry.count - 1);
    if (entryEnd < start || entry.startIndex > end) {
      continue;
    }
    for (
      let index = Math.max(start, entry.startIndex);
      index <= Math.min(end, entryEnd);
      index += 1
    ) {
      const axisEntry: AxisViewportEntry = { index, size: entry.size ?? defaultSize };
      if (entry.hidden !== undefined) {
        axisEntry.hidden = entry.hidden;
      }
      indexed.set(index, axisEntry);
    }
  }
  return indexed;
}

function buildAxisPatches(
  start: number,
  end: number,
  entries: Map<number, AxisViewportEntry>,
  defaultSize: number,
  previous: Map<number, string>,
  full: boolean,
): {
  patches: Array<{ index: number; size: number; hidden: boolean }>;
  signatures: Map<number, string>;
} {
  const signatures = new Map<number, string>();
  const patches: Array<{ index: number; size: number; hidden: boolean }> = [];
  for (let index = start; index <= end; index += 1) {
    const entry = entries.get(index);
    const size = entry?.size ?? defaultSize;
    const hidden = entry?.hidden ?? false;
    const signature = `${size}:${hidden ? 1 : 0}`;
    signatures.set(index, signature);
    if (full || previous.get(index) !== signature) {
      patches.push({ index, size, hidden });
    }
  }
  return { patches, signatures };
}

export function createViewportProjectionState(): ViewportProjectionState {
  return {
    nextVersion: 1,
    knownStyleIds: new Set(),
    lastCellSignatures: new Map(),
    lastColumnSignatures: new Map(),
    lastRowSignatures: new Map(),
  };
}

export function buildStylesById(rows: readonly StyleRow[]): Map<string, CellStyleRecord> {
  const map = new Map<string, CellStyleRecord>([[DEFAULT_STYLE_ID, { id: DEFAULT_STYLE_ID }]]);
  for (const row of rows) {
    map.set(row.id, {
      ...row.recordJSON,
      id: row.id,
    });
  }
  return map;
}

export function buildNumberFormatCodeById(rows: readonly NumberFormatRow[]): Map<string, string> {
  return new Map(rows.map((row) => [row.id, row.code]));
}

export function buildSelectedCellSnapshot(
  sheetName: string,
  address: string,
  cached: CellSnapshot | undefined,
  source: CellSourceRow | null,
  numberFormatCodeById: ReadonlyMap<string, string>,
): CellSnapshot {
  const base = cached ?? emptyCellSnapshot(sheetName, address);
  if (!source) {
    return base;
  }

  const next: CellSnapshot = {
    ...base,
    sheetName,
    address,
  };
  if (source.formula != null) {
    next.formula = source.formula ?? undefined;
  }
  const inputValue = source.inputValue;
  if (isLiteralInput(inputValue)) {
    next.input = inputValue;
  }
  const format = source.format;
  if (format !== undefined) {
    next.format = format;
  } else if (source.explicitFormatId) {
    const code = numberFormatCodeById.get(source.explicitFormatId);
    if (code !== undefined) {
      next.format = code;
    }
  }
  if (source.explicitFormatId) {
    next.numberFormatId = source.explicitFormatId;
  }
  return next;
}

export function projectViewportPatch(
  state: ViewportProjectionState,
  input: ViewportProjectionInput,
  full: boolean,
): ViewportPatch {
  const styles: CellStyleRecord[] = [];
  const cells: ViewportPatchedCell[] = [];
  const nextCellSignatures = new Map<string, string>();
  const sourceByAddress = new Map(input.sourceCells.map((row) => [row.address, row]));
  const computedByAddress = new Map(input.cellEval.map((row) => [row.address, row]));
  const sortedStyleRanges = input.styleRanges.toSorted(compareRectRanges);
  const sortedFormatRanges = input.formatRanges.toSorted(compareRectRanges);

  for (let row = input.viewport.rowStart; row <= input.viewport.rowEnd; row += 1) {
    for (let col = input.viewport.colStart; col <= input.viewport.colEnd; col += 1) {
      const address = formatAddress(row, col);
      const source = sourceByAddress.get(address);
      const computed = computedByAddress.get(address);
      const styleId = resolveStyleId(row, col, sortedStyleRanges);
      const style = input.stylesById.get(styleId) ?? { id: DEFAULT_STYLE_ID };
      if (full || !state.knownStyleIds.has(style.id)) {
        state.knownStyleIds.add(style.id);
        styles.push(style);
      }

      const numberFormat = resolveNumberFormat(
        row,
        col,
        source,
        sortedFormatRanges,
        input.numberFormatCodeById,
      );
      const snapshot: CellSnapshot = computed
        ? {
            sheetName: input.viewport.sheetName,
            address,
            value: computed.value,
            flags: computed.flags,
            version: computed.version,
          }
        : emptyCellSnapshot(input.viewport.sheetName, address);
      if (style.id !== DEFAULT_STYLE_ID) {
        snapshot.styleId = style.id;
      }
      if (numberFormat.code !== undefined) {
        snapshot.format = numberFormat.code;
      }
      if (numberFormat.numberFormatId !== undefined) {
        snapshot.numberFormatId = numberFormat.numberFormatId;
      }
      if (source?.formula != null) {
        snapshot.formula = source.formula ?? undefined;
      }
      const sourceInputValue = source?.inputValue;
      if (isLiteralInput(sourceInputValue)) {
        snapshot.input = sourceInputValue;
      }
      const patchedCell = buildPatchedCell(snapshot, row, col, style.id);
      const signature = JSON.stringify([
        patchedCell.snapshot.version,
        patchedCell.snapshot.formula ?? "",
        patchedCell.snapshot.input ?? null,
        patchedCell.snapshot.format ?? "",
        patchedCell.snapshot.numberFormatId ?? "",
        patchedCell.snapshot.styleId ?? "",
        patchedCell.snapshot.value,
        patchedCell.displayText,
        patchedCell.copyText,
        patchedCell.editorText,
      ]);
      const key = `${input.viewport.sheetName}!${address}`;
      nextCellSignatures.set(key, signature);
      if (full || state.lastCellSignatures.get(key) !== signature) {
        cells.push(patchedCell);
      }
    }
  }
  state.lastCellSignatures = nextCellSignatures;

  const columnEntries = buildAxisEntries(
    input.viewport,
    input.columnMetadata,
    PRODUCT_COLUMN_WIDTH,
    "colStart",
    "colEnd",
  );
  const rowEntries = buildAxisEntries(
    input.viewport,
    input.rowMetadata,
    PRODUCT_ROW_HEIGHT,
    "rowStart",
    "rowEnd",
  );
  const { patches: columns, signatures: columnSignatures } = buildAxisPatches(
    input.viewport.colStart,
    input.viewport.colEnd,
    columnEntries,
    PRODUCT_COLUMN_WIDTH,
    state.lastColumnSignatures,
    full,
  );
  const { patches: rows, signatures: rowSignatures } = buildAxisPatches(
    input.viewport.rowStart,
    input.viewport.rowEnd,
    rowEntries,
    PRODUCT_ROW_HEIGHT,
    state.lastRowSignatures,
    full,
  );
  state.lastColumnSignatures = columnSignatures;
  state.lastRowSignatures = rowSignatures;

  return {
    version: state.nextVersion++,
    full,
    viewport: input.viewport,
    metrics: input.metrics ?? EMPTY_METRICS,
    styles,
    cells,
    columns,
    rows,
  };
}
