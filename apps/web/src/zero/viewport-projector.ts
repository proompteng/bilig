import { formatAddress } from "@bilig/formula";
import {
  ValueTag,
  formatCellDisplayValue,
  type CellSnapshot,
  type LiteralInput,
  type CellStyleRecord,
  type CellValue,
  type RecalcMetrics,
  type Viewport,
} from "@bilig/protocol";

export interface CellSourceRow {
  workbookId: string;
  sheetName: string;
  address: string;
  rowNum?: number;
  colNum?: number;
  inputValue?: LiteralInput | null;
  formula?: string;
  styleId?: string | null;
  explicitFormatId?: string | null;
  editorText?: string;
}

export interface CellEvalRow {
  workbookId: string;
  sheetName: string;
  address: string;
  rowNum?: number;
  colNum?: number;
  value: CellValue;
  flags: number;
  version: number;
  styleId?: string | null;
  styleJson?: CellStyleRecord | null;
  formatId?: string | null;
  formatCode?: string | null;
  calcRevision?: number;
}

export interface AxisMetadataRow {
  workbookId: string;
  sheetName: string;
  startIndex: number;
  count: number;
  size?: number;
  hidden?: boolean;
}

export interface WorkbookRow {
  id: string;
  name: string;
  headRevision: number;
  calculatedRevision: number;
}

export interface SheetRow {
  id?: string;
  workbookId: string;
  name: string;
  position?: number;
  sortOrder?: number;
}

export interface CellStyleRow {
  workbookId: string;
  styleId: string;
  styleJson: CellStyleRecord;
}

export interface NumberFormatRow {
  workbookId: string;
  formatId: string;
  code: string;
}

export interface ViewportProjectionInput {
  viewport: Viewport & { sheetName: string };
  sourceCells: Map<string, CellSourceRow>;
  cellEval: Map<string, CellEvalRow>;
  rowMetadata: Map<string, AxisMetadataRow>;
  columnMetadata: Map<string, AxisMetadataRow>;
  stylesById?: ReadonlyMap<string, CellStyleRecord>;
  numberFormatsById?: ReadonlyMap<string, string>;
  metrics?: RecalcMetrics;
}

export interface ViewportProjectionState {
  nextVersion: number;
  knownStyleIds: Set<string>;
  lastCellSignatures: Map<string, string>;
  lastColumnSignatures: Map<number, string>;
  lastRowSignatures: Map<number, string>;
}

export interface ViewportPatch {
  version: number;
  full: boolean;
  viewport: Viewport & { sheetName: string };
  metrics: RecalcMetrics;
  styles: CellStyleRecord[];
  cells: Array<{
    row: number;
    col: number;
    snapshot: CellSnapshot;
    displayText: string;
    copyText: string;
    editorText: string;
    formatId: number;
    styleId: string;
  }>;
  columns: Array<{ index: number; size: number; hidden: boolean }>;
  rows: Array<{ index: number; size: number; hidden: boolean }>;
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

const PRODUCT_COLUMN_WIDTH = 104;
const PRODUCT_ROW_HEIGHT = 22;
const EMPTY_STYLE_MAP = new Map<string, CellStyleRecord>();
const EMPTY_FORMAT_MAP = new Map<string, string>();

function buildAxisPatches(
  start: number,
  end: number,
  entries: AxisMetadataRow[],
  defaultSize: number,
  lastSignatures: Map<number, string>,
  full: boolean,
): {
  patches: Array<{ index: number; size: number; hidden: boolean }>;
  signatures: Map<number, string>;
} {
  const patches: Array<{ index: number; size: number; hidden: boolean }> = [];
  const nextSignatures = new Map<number, string>();

  for (let i = start; i <= end; i += 1) {
    const entry = entries.find((e) => e.startIndex === i);
    const size = entry?.size ?? defaultSize;
    const hidden = entry?.hidden ?? false;
    const signature = `${size}:${hidden}`;
    nextSignatures.set(i, signature);

    if (full || lastSignatures.get(i) !== signature) {
      patches.push({ index: i, size, hidden });
    }
  }

  return { patches, signatures: nextSignatures };
}

function signatureOfCell(row: CellEvalRow, source: CellSourceRow | undefined): string {
  return [
    row.version,
    row.styleId ?? "",
    row.formatId ?? "",
    JSON.stringify(row.value),
    JSON.stringify(source?.inputValue ?? null),
    source?.formula ?? "",
    source?.styleId ?? "",
    source?.explicitFormatId ?? "",
  ].join(":");
}

function emptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty }, // Ensure value is ALWAYS defined
    flags: 0,
    version: 0,
  };
}

function resolveFormatCode(
  formatId: string | null | undefined,
  inlineFormatCode: string | null | undefined,
  numberFormatsById: ReadonlyMap<string, string>,
): string | undefined {
  if (inlineFormatCode !== undefined && inlineFormatCode !== null) {
    return inlineFormatCode;
  }
  if (formatId === undefined || formatId === null) {
    return undefined;
  }
  return numberFormatsById.get(formatId);
}

function literalInputToText(value: LiteralInput | undefined | null): string {
  if (value === undefined || value === null) {
    return "";
  }
  return String(value);
}

function literalInputToCellValue(value: LiteralInput | undefined | null): CellValue {
  if (value === undefined || value === null) {
    return { tag: ValueTag.Empty };
  }
  if (typeof value === "number") {
    return { tag: ValueTag.Number, value };
  }
  if (typeof value === "boolean") {
    return { tag: ValueTag.Boolean, value };
  }
  return { tag: ValueTag.String, value, stringId: 0 };
}

function applyLiteralValueFallback(
  snapshot: CellSnapshot,
  source: CellSourceRow | undefined,
): void {
  if (!source || source.formula !== undefined || source.inputValue === undefined) {
    return;
  }
  if (snapshot.value.tag !== ValueTag.Empty) {
    return;
  }
  snapshot.value = literalInputToCellValue(source.inputValue);
}

function applySourceCellOverrides(
  snapshot: CellSnapshot,
  source: CellSourceRow | undefined,
  numberFormatsById: ReadonlyMap<string, string>,
): void {
  if (!source) {
    return;
  }

  if (source.inputValue !== undefined) {
    snapshot.input = source.inputValue;
  }
  if (source.formula !== undefined) {
    snapshot.formula = source.formula;
  }
  if (source.styleId !== undefined) {
    if (source.styleId === null) {
      delete snapshot.styleId;
    } else {
      snapshot.styleId = source.styleId;
    }
  }
  if (source.explicitFormatId !== undefined) {
    if (source.explicitFormatId === null) {
      delete snapshot.numberFormatId;
    } else {
      snapshot.numberFormatId = source.explicitFormatId;
    }
    const formatCode = resolveFormatCode(source.explicitFormatId, undefined, numberFormatsById);
    if (formatCode !== undefined) {
      snapshot.format = formatCode;
    } else if (source.explicitFormatId === null) {
      delete snapshot.format;
    }
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
  numberFormatsById: ReadonlyMap<string, string> = EMPTY_FORMAT_MAP,
): CellSnapshot {
  const next: CellSnapshot = {
    ...(cached ?? emptyCellSnapshot(sheetName, address)),
    sheetName,
    address,
  };
  // Double check value is set
  if (!next.value) {
    next.value = { tag: ValueTag.Empty };
  }
  applySourceCellOverrides(next, source ?? undefined, numberFormatsById);
  applyLiteralValueFallback(next, source ?? undefined);
  return next;
}

export function projectViewportPatch(
  state: ViewportProjectionState,
  input: ViewportProjectionInput,
  full: boolean,
): ViewportPatch {
  const styles: CellStyleRecord[] = [];
  const cells: ViewportPatch["cells"] = [];
  const stylesById = input.stylesById ?? EMPTY_STYLE_MAP;
  const numberFormatsById = input.numberFormatsById ?? EMPTY_FORMAT_MAP;

  for (let row = input.viewport.rowStart; row <= input.viewport.rowEnd; row += 1) {
    for (let col = input.viewport.colStart; col <= input.viewport.colEnd; col += 1) {
      const address = formatAddress(row, col);
      const evalRow = input.cellEval.get(address);
      const sourceRow = input.sourceCells.get(address);

      if (!evalRow && !sourceRow) {
        if (full) {
          cells.push({
            row,
            col,
            snapshot: emptyCellSnapshot(input.viewport.sheetName, address),
            displayText: "",
            copyText: "",
            editorText: "",
            formatId: 0,
            styleId: "style-0",
          });
        }
        continue;
      }

      const signature = evalRow ? signatureOfCell(evalRow, sourceRow) : "empty";
      if (!full && state.lastCellSignatures.get(address) === signature) {
        continue;
      }
      state.lastCellSignatures.set(address, signature);

      const formatCode = resolveFormatCode(
        evalRow?.formatId,
        evalRow?.formatCode,
        numberFormatsById,
      );
      const snapshot: CellSnapshot = evalRow
        ? {
            sheetName: input.viewport.sheetName,
            address,
            value: evalRow.value,
            flags: evalRow.flags,
            version: evalRow.version,
            ...(evalRow.styleId ? { styleId: evalRow.styleId } : {}),
            ...(formatCode ? { format: formatCode } : {}),
            ...(evalRow.formatId ? { numberFormatId: evalRow.formatId } : {}),
          }
        : emptyCellSnapshot(input.viewport.sheetName, address);

      applySourceCellOverrides(snapshot, sourceRow, numberFormatsById);
      applyLiteralValueFallback(snapshot, sourceRow);

      const styleRecord =
        (snapshot.styleId ? stylesById.get(snapshot.styleId) : undefined) ??
        evalRow?.styleJson ??
        undefined;
      if (styleRecord && !state.knownStyleIds.has(styleRecord.id)) {
        styles.push(styleRecord);
        state.knownStyleIds.add(styleRecord.id);
      }

      const editorText =
        sourceRow?.editorText ??
        (snapshot.formula ? `=${snapshot.formula}` : literalInputToText(snapshot.input));
      cells.push({
        row,
        col,
        snapshot,
        displayText: formatCellDisplayValue(snapshot.value, snapshot.format),
        copyText: snapshot.formula ? `=${snapshot.formula}` : literalInputToText(snapshot.input),
        editorText,
        formatId: 0,
        styleId: snapshot.styleId ?? "style-0",
      });
    }
  }

  const columnEntries = [...input.columnMetadata.values()];
  const rowEntries = [...input.rowMetadata.values()];

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
