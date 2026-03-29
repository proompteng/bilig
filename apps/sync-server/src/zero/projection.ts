import type { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { CellValue, WorkbookSnapshot } from "@bilig/protocol";

export interface WorkbookSourceRow {
  id: string;
  name: string;
  ownerUserId: string;
  headRevision: number;
  calculatedRevision: number;
  calcMode: "automatic" | "manual";
  compatibilityMode: "excel-modern" | "odf-1.4";
  recalcEpoch: number;
  updatedAt: string;
}

export interface SheetSourceRow {
  workbookId: string;
  name: string;
  sortOrder: number;
  freezeRows: number;
  freezeCols: number;
  updatedAt: string;
}

export interface CellSourceRow {
  workbookId: string;
  sheetName: string;
  address: string;
  rowNum: number;
  colNum: number;
  inputValue: unknown;
  formula: string | null;
  format: string | null;
  explicitFormatId: string | null;
  sourceRevision: number;
  updatedBy: string;
  updatedAt: string;
}

export interface AxisMetadataSourceRow {
  workbookId: string;
  sheetName: string;
  startIndex: number;
  count: number;
  size: number | null;
  hidden: boolean | null;
  sourceRevision: number;
  updatedAt: string;
}

export interface DefinedNameSourceRow {
  workbookId: string;
  name: string;
  value: unknown;
}

export interface WorkbookMetadataSourceRow {
  workbookId: string;
  key: string;
  value: unknown;
}

export interface CalculationSettingsSourceRow {
  workbookId: string;
  mode: "automatic" | "manual";
  recalcEpoch: number;
}

export interface StyleSourceRow {
  workbookId: string;
  id: string;
  recordJSON: unknown;
  hash: string;
  createdAt: string;
}

export interface NumberFormatSourceRow {
  workbookId: string;
  id: string;
  code: string;
  kind: string;
  createdAt: string;
}

export interface StyleRangeSourceRow {
  id: string;
  workbookId: string;
  sheetName: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  styleId: string;
  sourceRevision: number;
  updatedAt: string;
}

export interface FormatRangeSourceRow {
  id: string;
  workbookId: string;
  sheetName: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  formatId: string;
  sourceRevision: number;
  updatedAt: string;
}

export interface CellEvalRow {
  workbookId: string;
  sheetName: string;
  address: string;
  rowNum: number;
  colNum: number;
  value: CellValue;
  flags: number;
  version: number;
  styleId: string | null;
  formatId: string | null;
  formatCode: string | null;
  calcRevision: number;
  updatedAt: string;
}

export interface WorkbookSourceProjection {
  workbook: WorkbookSourceRow;
  sheets: readonly SheetSourceRow[];
  cells: readonly CellSourceRow[];
  rowMetadata: readonly AxisMetadataSourceRow[];
  columnMetadata: readonly AxisMetadataSourceRow[];
  definedNames: readonly DefinedNameSourceRow[];
  workbookMetadataEntries: readonly WorkbookMetadataSourceRow[];
  calculationSettings: CalculationSettingsSourceRow;
  styles: readonly StyleSourceRow[];
  numberFormats: readonly NumberFormatSourceRow[];
  styleRanges: readonly StyleRangeSourceRow[];
  formatRanges: readonly FormatRangeSourceRow[];
}

export interface ProjectionDiff<Row> {
  upserts: Row[];
  deletes: string[];
}

export interface WorkbookProjectionOptions {
  revision: number;
  calculatedRevision: number;
  ownerUserId: string;
  updatedBy: string;
  updatedAt: string;
}

function signatureOf(row: unknown): string {
  return JSON.stringify(row);
}

export function diffProjectionRows<Row>(
  previousRows: readonly Row[],
  nextRows: readonly Row[],
  keyOf: (row: Row) => string,
  signatureOfRow: (row: Row) => string = signatureOf,
): ProjectionDiff<Row> {
  const previous = new Map(
    previousRows.map((row) => [
      keyOf(row),
      {
        row,
        signature: signatureOfRow(row),
      },
    ]),
  );
  const next = new Map(
    nextRows.map((row) => [
      keyOf(row),
      {
        row,
        signature: signatureOfRow(row),
      },
    ]),
  );
  const upserts: Row[] = [];
  const deletes: string[] = [];

  for (const [key, entry] of next) {
    const current = previous.get(key);
    if (!current || current.signature !== entry.signature) {
      upserts.push(entry.row);
    }
  }

  for (const key of previous.keys()) {
    if (!next.has(key)) {
      deletes.push(key);
    }
  }

  return { upserts, deletes };
}

function rangeId(
  prefix: "style-range" | "format-range",
  documentId: string,
  sheetName: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): string {
  return `${prefix}:${documentId}:${sheetName}:${startRow}:${startCol}:${endRow}:${endCol}`;
}

export function buildWorkbookHeaderRow(
  documentId: string,
  snapshot: WorkbookSnapshot,
  options: WorkbookProjectionOptions,
): WorkbookSourceRow {
  const calcSettings = snapshot.workbook.metadata?.calculationSettings;
  const recalcEpoch = snapshot.workbook.metadata?.volatileContext?.recalcEpoch ?? 0;
  return {
    id: documentId,
    name: snapshot.workbook.name,
    ownerUserId: options.ownerUserId,
    headRevision: options.revision,
    calculatedRevision: options.calculatedRevision,
    calcMode: calcSettings?.mode ?? "automatic",
    compatibilityMode: calcSettings?.compatibilityMode ?? "excel-modern",
    recalcEpoch,
    updatedAt: options.updatedAt,
  };
}

export function buildCalculationSettingsRow(
  documentId: string,
  snapshot: WorkbookSnapshot,
): CalculationSettingsSourceRow {
  const calcSettings = snapshot.workbook.metadata?.calculationSettings;
  const recalcEpoch = snapshot.workbook.metadata?.volatileContext?.recalcEpoch ?? 0;
  return {
    workbookId: documentId,
    mode: calcSettings?.mode ?? "automatic",
    recalcEpoch,
  };
}

function buildCellSourceRow(
  documentId: string,
  sheetName: string,
  cell: WorkbookSnapshot["sheets"][number]["cells"][number],
  options: WorkbookProjectionOptions,
): CellSourceRow {
  const parsed = parseCellAddress(cell.address, sheetName);
  return {
    workbookId: documentId,
    sheetName,
    address: cell.address,
    rowNum: parsed.row,
    colNum: parsed.col,
    inputValue: cell.formula ? null : (cell.value ?? null),
    formula: cell.formula ?? null,
    format: cell.format ?? null,
    explicitFormatId: null,
    sourceRevision: options.revision,
    updatedBy: options.updatedBy,
    updatedAt: options.updatedAt,
  };
}

function findSheet(snapshot: WorkbookSnapshot, sheetName: string) {
  return snapshot.sheets.find((sheet) => sheet.name === sheetName);
}

function buildStyleSourceRow(
  documentId: string,
  style: NonNullable<NonNullable<WorkbookSnapshot["workbook"]["metadata"]>["styles"]>[number],
  options: WorkbookProjectionOptions,
): StyleSourceRow {
  return {
    workbookId: documentId,
    id: style.id,
    recordJSON: style,
    hash: JSON.stringify(style),
    createdAt: options.updatedAt,
  };
}

function buildNumberFormatSourceRow(
  documentId: string,
  format: NonNullable<NonNullable<WorkbookSnapshot["workbook"]["metadata"]>["formats"]>[number],
  options: WorkbookProjectionOptions,
): NumberFormatSourceRow {
  return {
    workbookId: documentId,
    id: format.id,
    code: format.code,
    kind: format.kind,
    createdAt: options.updatedAt,
  };
}

export function buildSingleCellSourceRow(
  documentId: string,
  snapshot: WorkbookSnapshot,
  sheetName: string,
  address: string,
  options: WorkbookProjectionOptions,
): CellSourceRow | null {
  const sheet = snapshot.sheets.find((entry) => entry.name === sheetName);
  const cell = sheet?.cells.find((entry) => entry.address === address);
  if (!cell) {
    return null;
  }
  return buildCellSourceRow(documentId, sheetName, cell, options);
}

export function buildSheetColumnMetadataRows(
  documentId: string,
  snapshot: WorkbookSnapshot,
  sheetName: string,
  options: WorkbookProjectionOptions,
): AxisMetadataSourceRow[] {
  const rows: AxisMetadataSourceRow[] = [];
  for (const entry of findSheet(snapshot, sheetName)?.metadata?.columnMetadata ?? []) {
    rows.push({
      workbookId: documentId,
      sheetName,
      startIndex: entry.start,
      count: entry.count,
      size: entry.size ?? null,
      hidden: entry.hidden ?? null,
      sourceRevision: options.revision,
      updatedAt: options.updatedAt,
    });
  }
  return rows;
}

export function buildWorkbookStyleRows(
  documentId: string,
  snapshot: WorkbookSnapshot,
  options: WorkbookProjectionOptions,
): StyleSourceRow[] {
  const rows: StyleSourceRow[] = [];
  for (const style of snapshot.workbook.metadata?.styles ?? []) {
    rows.push(buildStyleSourceRow(documentId, style, options));
  }
  return rows;
}

export function buildSheetStyleRangeRows(
  documentId: string,
  snapshot: WorkbookSnapshot,
  sheetName: string,
  options: WorkbookProjectionOptions,
): StyleRangeSourceRow[] {
  const rows: StyleRangeSourceRow[] = [];
  for (const entry of findSheet(snapshot, sheetName)?.metadata?.styleRanges ?? []) {
    const start = parseCellAddress(entry.range.startAddress, sheetName);
    const end = parseCellAddress(entry.range.endAddress, sheetName);
    rows.push({
      id: rangeId("style-range", documentId, sheetName, start.row, start.col, end.row, end.col),
      workbookId: documentId,
      sheetName,
      startRow: start.row,
      endRow: end.row,
      startCol: start.col,
      endCol: end.col,
      styleId: entry.styleId,
      sourceRevision: options.revision,
      updatedAt: options.updatedAt,
    });
  }
  return rows;
}

export function buildWorkbookNumberFormatRows(
  documentId: string,
  snapshot: WorkbookSnapshot,
  options: WorkbookProjectionOptions,
): NumberFormatSourceRow[] {
  const rows: NumberFormatSourceRow[] = [];
  for (const format of snapshot.workbook.metadata?.formats ?? []) {
    rows.push(buildNumberFormatSourceRow(documentId, format, options));
  }
  return rows;
}

export function buildSheetFormatRangeRows(
  documentId: string,
  snapshot: WorkbookSnapshot,
  sheetName: string,
  options: WorkbookProjectionOptions,
): FormatRangeSourceRow[] {
  const rows: FormatRangeSourceRow[] = [];
  for (const entry of findSheet(snapshot, sheetName)?.metadata?.formatRanges ?? []) {
    const start = parseCellAddress(entry.range.startAddress, sheetName);
    const end = parseCellAddress(entry.range.endAddress, sheetName);
    rows.push({
      id: rangeId("format-range", documentId, sheetName, start.row, start.col, end.row, end.col),
      workbookId: documentId,
      sheetName,
      startRow: start.row,
      endRow: end.row,
      startCol: start.col,
      endCol: end.col,
      formatId: entry.formatId,
      sourceRevision: options.revision,
      updatedAt: options.updatedAt,
    });
  }
  return rows;
}

export function buildWorkbookSourceProjection(
  documentId: string,
  snapshot: WorkbookSnapshot,
  options: WorkbookProjectionOptions,
): WorkbookSourceProjection {
  const sheets: SheetSourceRow[] = [];
  const cells: CellSourceRow[] = [];
  const rowMetadata: AxisMetadataSourceRow[] = [];
  const columnMetadata: AxisMetadataSourceRow[] = [];
  const styleRanges: StyleRangeSourceRow[] = [];
  const formatRanges: FormatRangeSourceRow[] = [];
  const styles: StyleSourceRow[] = [];
  const numberFormats: NumberFormatSourceRow[] = [];
  const definedNames: DefinedNameSourceRow[] = [];
  const workbookMetadataEntries: WorkbookMetadataSourceRow[] = [];

  for (const sheet of snapshot.sheets) {
    sheets.push({
      workbookId: documentId,
      name: sheet.name,
      sortOrder: sheet.order,
      freezeRows: sheet.metadata?.freezePane?.rows ?? 0,
      freezeCols: sheet.metadata?.freezePane?.cols ?? 0,
      updatedAt: options.updatedAt,
    });

    for (const cell of sheet.cells) {
      cells.push(buildCellSourceRow(documentId, sheet.name, cell, options));
    }

    for (const entry of sheet.metadata?.rowMetadata ?? []) {
      rowMetadata.push({
        workbookId: documentId,
        sheetName: sheet.name,
        startIndex: entry.start,
        count: entry.count,
        size: entry.size ?? null,
        hidden: entry.hidden ?? null,
        sourceRevision: options.revision,
        updatedAt: options.updatedAt,
      });
    }

    for (const entry of sheet.metadata?.columnMetadata ?? []) {
      columnMetadata.push({
        workbookId: documentId,
        sheetName: sheet.name,
        startIndex: entry.start,
        count: entry.count,
        size: entry.size ?? null,
        hidden: entry.hidden ?? null,
        sourceRevision: options.revision,
        updatedAt: options.updatedAt,
      });
    }

    for (const entry of sheet.metadata?.styleRanges ?? []) {
      const start = parseCellAddress(entry.range.startAddress, sheet.name);
      const end = parseCellAddress(entry.range.endAddress, sheet.name);
      styleRanges.push({
        id: rangeId("style-range", documentId, sheet.name, start.row, start.col, end.row, end.col),
        workbookId: documentId,
        sheetName: sheet.name,
        startRow: start.row,
        endRow: end.row,
        startCol: start.col,
        endCol: end.col,
        styleId: entry.styleId,
        sourceRevision: options.revision,
        updatedAt: options.updatedAt,
      });
    }

    for (const entry of sheet.metadata?.formatRanges ?? []) {
      const start = parseCellAddress(entry.range.startAddress, sheet.name);
      const end = parseCellAddress(entry.range.endAddress, sheet.name);
      formatRanges.push({
        id: rangeId("format-range", documentId, sheet.name, start.row, start.col, end.row, end.col),
        workbookId: documentId,
        sheetName: sheet.name,
        startRow: start.row,
        endRow: end.row,
        startCol: start.col,
        endCol: end.col,
        formatId: entry.formatId,
        sourceRevision: options.revision,
        updatedAt: options.updatedAt,
      });
    }
  }

  styles.push(...buildWorkbookStyleRows(documentId, snapshot, options));
  numberFormats.push(...buildWorkbookNumberFormatRows(documentId, snapshot, options));

  for (const entry of snapshot.workbook.metadata?.definedNames ?? []) {
    definedNames.push({
      workbookId: documentId,
      name: entry.name,
      value: entry.value,
    });
  }

  for (const entry of snapshot.workbook.metadata?.properties ?? []) {
    workbookMetadataEntries.push({
      workbookId: documentId,
      key: entry.key,
      value: entry.value,
    });
  }

  return {
    workbook: buildWorkbookHeaderRow(documentId, snapshot, options),
    sheets,
    cells,
    rowMetadata,
    columnMetadata,
    definedNames,
    workbookMetadataEntries,
    calculationSettings: buildCalculationSettingsRow(documentId, snapshot),
    styles,
    numberFormats,
    styleRanges,
    formatRanges,
  };
}

export function materializeCellEvalProjection(
  engine: SpreadsheetEngine,
  documentId: string,
  revision: number,
  updatedAt: string,
): CellEvalRow[] {
  const entries: CellEvalRow[] = [];

  for (const sheet of engine.workbook.sheetsByName.values()) {
    sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
      const address = formatAddress(row, col);
      const cell = engine.getCell(sheet.name, address);
      entries.push({
        workbookId: documentId,
        sheetName: sheet.name,
        address,
        rowNum: row,
        colNum: col,
        value: cell.value,
        flags: cell.flags,
        version: cell.version,
        styleId: cell.styleId ?? null,
        formatId: cell.numberFormatId ?? null,
        formatCode: cell.format ?? null,
        calcRevision: revision,
        updatedAt,
      });
    });
  }

  return entries;
}

export const sourceProjectionKeys = {
  sheet: (row: SheetSourceRow) => JSON.stringify([row.workbookId, row.name]),
  cell: (row: CellSourceRow) => JSON.stringify([row.workbookId, row.sheetName, row.address]),
  axisMetadata: (row: AxisMetadataSourceRow) =>
    JSON.stringify([row.workbookId, row.sheetName, row.startIndex]),
  definedName: (row: DefinedNameSourceRow) => JSON.stringify([row.workbookId, row.name]),
  workbookMetadata: (row: WorkbookMetadataSourceRow) => JSON.stringify([row.workbookId, row.key]),
  style: (row: StyleSourceRow) => JSON.stringify([row.workbookId, row.id]),
  numberFormat: (row: NumberFormatSourceRow) => JSON.stringify([row.workbookId, row.id]),
  styleRange: (row: StyleRangeSourceRow) => row.id,
  formatRange: (row: FormatRangeSourceRow) => row.id,
  cellEval: (row: CellEvalRow) => JSON.stringify([row.workbookId, row.sheetName, row.address]),
} as const;
