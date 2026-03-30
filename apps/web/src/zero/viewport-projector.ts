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
  styleId?: string | null | undefined;
  styleJson?: CellStyleRecord | null | undefined;
  formatId?: string | null | undefined;
  formatCode?: string | null | undefined;
}

export interface AxisMetadataRow {
  workbookId: string;
  sheetName: string;
  startIndex: number;
  count: number;
  size?: number | null | undefined;
  hidden?: boolean | null | undefined;
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
  sourceCells: ReadonlyMap<string, CellSourceRow>;
  cellEval: ReadonlyMap<string, CellEvalRow>;
  rowMetadata: ReadonlyMap<string, AxisMetadataRow>;
  columnMetadata: ReadonlyMap<string, AxisMetadataRow>;
}

interface AxisViewportEntry {
  index: number;
  size?: number | null;
  hidden?: boolean | null;
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

function applySourceCellOverrides(snapshot: CellSnapshot, source: CellSourceRow | undefined): void {
  if (!source) {
    return;
  }

  if (source.formula !== undefined) {
    if (source.formula === null) {
      delete snapshot.formula;
    } else {
      snapshot.formula = source.formula;
    }
  }
  if (isLiteralInput(source.inputValue)) {
    snapshot.input = source.inputValue;
  }

  if (source.format !== undefined) {
    snapshot.format = source.format;
  }
  if (source.explicitFormatId !== undefined) {
    snapshot.numberFormatId = source.explicitFormatId;
  }
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

export function buildSelectedCellSnapshot(
  sheetName: string,
  address: string,
  cached: CellSnapshot | undefined,
  source: CellSourceRow | null,
): CellSnapshot {
  const next: CellSnapshot = {
    ...(cached ?? emptyCellSnapshot(sheetName, address)),
    sheetName,
    address,
  };
  applySourceCellOverrides(next, source ?? undefined);
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
  const sourceByAddress = input.sourceCells;
  const computedByAddress = input.cellEval;

  for (let row = input.viewport.rowStart; row <= input.viewport.rowEnd; row += 1) {
    for (let col = input.viewport.colStart; col <= input.viewport.colEnd; col += 1) {
      const address = formatAddress(row, col);
      const source = sourceByAddress.get(address);
      const computed = computedByAddress.get(address);
      const styleId = computed?.styleJson?.id ?? computed?.styleId ?? DEFAULT_STYLE_ID;
      const style = computed?.styleJson ?? undefined;

      if (style && styleId !== DEFAULT_STYLE_ID && (full || !state.knownStyleIds.has(style.id))) {
        state.knownStyleIds.add(style.id);
        styles.push(style);
      }

      const snapshot: CellSnapshot = computed
        ? {
            sheetName: input.viewport.sheetName,
            address,
            value: computed.value,
            flags: computed.flags,
            version: computed.version,
          }
        : emptyCellSnapshot(input.viewport.sheetName, address);

      if (styleId !== DEFAULT_STYLE_ID) {
        snapshot.styleId = styleId;
      }
      if (computed?.formatCode !== undefined && computed.formatCode !== null) {
        snapshot.format = computed.formatCode;
      }
      if (computed?.formatId !== undefined && computed.formatId !== null) {
        snapshot.numberFormatId = computed.formatId;
      }

      applySourceCellOverrides(snapshot, source);

      const patchedCell = buildPatchedCell(snapshot, row, col, styleId);
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
    [...input.columnMetadata.values()],
    PRODUCT_COLUMN_WIDTH,
    "colStart",
    "colEnd",
  );
  const rowEntries = buildAxisEntries(
    input.viewport,
    [...input.rowMetadata.values()],
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
