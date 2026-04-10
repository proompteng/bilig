import type { SpreadsheetEngine } from "@bilig/core";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import {
  ValueTag,
  type CellRangeRef,
  type CellStyleRecord,
  type CellValue,
  type WorkbookSnapshot,
} from "@bilig/protocol";

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
  sheetId: number;
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
  styleId: string | null;
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
  styleJson: CellStyleRecord | null;
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

function buildWorkbookHeaderRow(
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

export function buildWorkbookHeaderRowFromEngine(
  documentId: string,
  engine: SpreadsheetEngine,
  options: WorkbookProjectionOptions,
): WorkbookSourceRow {
  const calcSettings = engine.getCalculationSettings();
  const recalcEpoch = engine.getVolatileContext().recalcEpoch ?? 0;
  return {
    id: documentId,
    name: engine.workbook.workbookName,
    ownerUserId: options.ownerUserId,
    headRevision: options.revision,
    calculatedRevision: options.calculatedRevision,
    calcMode: calcSettings.mode ?? "automatic",
    compatibilityMode: calcSettings.compatibilityMode ?? "excel-modern",
    recalcEpoch,
    updatedAt: options.updatedAt,
  };
}

function buildCalculationSettingsRow(
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

export function buildCalculationSettingsRowFromEngine(
  documentId: string,
  engine: SpreadsheetEngine,
): CalculationSettingsSourceRow {
  const calcSettings = engine.getCalculationSettings();
  const recalcEpoch = engine.getVolatileContext().recalcEpoch ?? 0;
  return {
    workbookId: documentId,
    mode: calcSettings.mode ?? "automatic",
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
    styleId: null,
    explicitFormatId: null,
    sourceRevision: options.revision,
    updatedBy: options.updatedBy,
    updatedAt: options.updatedAt,
  };
}

function buildCellSourceRowFromEngineCell(
  documentId: string,
  sheetName: string,
  address: string,
  row: number,
  col: number,
  cell: {
    value: CellValue;
    version: number;
    formula?: string;
    format?: string;
    styleId?: string;
    numberFormatId?: string;
  },
  options: WorkbookProjectionOptions,
): CellSourceRow | null {
  if (
    cell.formula === undefined &&
    (cell.value.tag === ValueTag.Empty || cell.value.tag === ValueTag.Error) &&
    cell.version === 0 &&
    cell.styleId === undefined &&
    cell.numberFormatId === undefined &&
    cell.format === undefined
  ) {
    return null;
  }
  return {
    workbookId: documentId,
    sheetName,
    address,
    rowNum: row,
    colNum: col,
    inputValue: cell.formula ? null : literalInputFromCellValue(cell.value),
    formula: cell.formula ?? null,
    format: cell.format ?? null,
    styleId: cell.styleId ?? null,
    explicitFormatId: cell.numberFormatId ?? null,
    sourceRevision: options.revision,
    updatedBy: options.updatedBy,
    updatedAt: options.updatedAt,
  };
}

function literalInputFromCellValue(value: CellValue): number | boolean | string | null {
  switch (value.tag) {
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value;
    case ValueTag.Empty:
    case ValueTag.Error:
    default:
      return null;
  }
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

export function buildSingleCellSourceRowFromEngine(
  documentId: string,
  engine: SpreadsheetEngine,
  sheetName: string,
  address: string,
  options: WorkbookProjectionOptions,
): CellSourceRow | null {
  const { row, col } = parseCellAddress(address, sheetName);
  return buildCellSourceRowFromEngineCell(
    documentId,
    sheetName,
    address,
    row,
    col,
    engine.getCell(sheetName, address),
    options,
  );
}

export function buildSheetColumnMetadataRowsFromEngine(
  documentId: string,
  engine: SpreadsheetEngine,
  sheetName: string,
  options: WorkbookProjectionOptions,
): AxisMetadataSourceRow[] {
  return engine.getColumnMetadata(sheetName).map((entry) => ({
    workbookId: documentId,
    sheetName,
    startIndex: entry.start,
    count: entry.count,
    size: entry.size ?? null,
    hidden: entry.hidden ?? null,
    sourceRevision: options.revision,
    updatedAt: options.updatedAt,
  }));
}

export function buildSheetRowMetadataRowsFromEngine(
  documentId: string,
  engine: SpreadsheetEngine,
  sheetName: string,
  options: WorkbookProjectionOptions,
): AxisMetadataSourceRow[] {
  return engine.getRowMetadata(sheetName).map((entry) => ({
    workbookId: documentId,
    sheetName,
    startIndex: entry.start,
    count: entry.count,
    size: entry.size ?? null,
    hidden: entry.hidden ?? null,
    sourceRevision: options.revision,
    updatedAt: options.updatedAt,
  }));
}

function buildWorkbookStyleRows(
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

export function buildWorkbookStyleRowsFromEngine(
  documentId: string,
  engine: SpreadsheetEngine,
  options: WorkbookProjectionOptions,
): StyleSourceRow[] {
  const referencedIds = collectReferencedFormattingIds(engine).styleIds;
  return engine.workbook
    .listCellStyles()
    .filter((style) => referencedIds.has(style.id))
    .map((style) => ({
      workbookId: documentId,
      id: style.id,
      recordJSON: style,
      hash: JSON.stringify(style),
      createdAt: options.updatedAt,
    }));
}

function buildWorkbookNumberFormatRows(
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

export function buildWorkbookNumberFormatRowsFromEngine(
  documentId: string,
  engine: SpreadsheetEngine,
  options: WorkbookProjectionOptions,
): NumberFormatSourceRow[] {
  const referencedIds = collectReferencedFormattingIds(engine).formatIds;
  return engine.workbook
    .listCellNumberFormats()
    .filter((format) => referencedIds.has(format.id))
    .map((format) => ({
      workbookId: documentId,
      id: format.id,
      code: format.code,
      kind: format.kind,
      createdAt: options.updatedAt,
    }));
}

function collectReferencedFormattingIds(engine: SpreadsheetEngine): {
  styleIds: Set<string>;
  formatIds: Set<string>;
} {
  const styleIds = new Set<string>();
  const formatIds = new Set<string>();
  engine.workbook.sheetsByName.forEach((sheet) => {
    sheet.styleRanges.forEach((record) => styleIds.add(record.styleId));
    sheet.formatRanges.forEach((record) => formatIds.add(record.formatId));
  });
  for (let cellIndex = 0; cellIndex < engine.workbook.cellStore.size; cellIndex += 1) {
    const explicitFormat = engine.workbook.getCellFormat(cellIndex);
    if (explicitFormat !== undefined) {
      formatIds.add(engine.workbook.internCellNumberFormat(explicitFormat).id);
    }
  }
  return { styleIds, formatIds };
}

function normalizeRangeBounds(
  sheetName: string,
  startAddress: string,
  endAddress: string,
): { rowStart: number; rowEnd: number; colStart: number; colEnd: number } {
  const start = parseCellAddress(startAddress, sheetName);
  const end = parseCellAddress(endAddress, sheetName);
  return {
    rowStart: Math.min(start.row, end.row),
    rowEnd: Math.max(start.row, end.row),
    colStart: Math.min(start.col, end.col),
    colEnd: Math.max(start.col, end.col),
  };
}

function addressWithinBounds(
  row: number,
  col: number,
  bounds?: { rowStart: number; rowEnd: number; colStart: number; colEnd: number },
): boolean {
  if (!bounds) {
    return true;
  }
  return (
    row >= bounds.rowStart && row <= bounds.rowEnd && col >= bounds.colStart && col <= bounds.colEnd
  );
}

function upsertCellSourceRow(
  rows: Map<string, CellSourceRow>,
  documentId: string,
  sheetName: string,
  row: number,
  col: number,
  options: WorkbookProjectionOptions,
): CellSourceRow {
  const address = formatAddress(row, col);
  const existing = rows.get(address);
  if (existing) {
    return existing;
  }
  const next: CellSourceRow = {
    workbookId: documentId,
    sheetName,
    address,
    rowNum: row,
    colNum: col,
    inputValue: null,
    formula: null,
    format: null,
    styleId: null,
    explicitFormatId: null,
    sourceRevision: options.revision,
    updatedBy: options.updatedBy,
    updatedAt: options.updatedAt,
  };
  rows.set(address, next);
  return next;
}

export function buildSheetCellSourceRows(
  documentId: string,
  snapshot: WorkbookSnapshot,
  sheetName: string,
  options: WorkbookProjectionOptions,
  range?: CellRangeRef,
): CellSourceRow[] {
  const sheet = findSheet(snapshot, sheetName);
  if (!sheet) {
    return [];
  }
  const formatCodeById = new Map(
    (snapshot.workbook.metadata?.formats ?? []).map((entry) => [entry.id, entry.code]),
  );
  const bounds = range
    ? normalizeRangeBounds(range.sheetName, range.startAddress, range.endAddress)
    : undefined;
  const rows = new Map<string, CellSourceRow>();

  for (const cell of sheet.cells) {
    const parsed = parseCellAddress(cell.address, sheetName);
    if (!addressWithinBounds(parsed.row, parsed.col, bounds)) {
      continue;
    }
    rows.set(cell.address, buildCellSourceRow(documentId, sheetName, cell, options));
  }

  for (const entry of sheet.metadata?.styleRanges ?? []) {
    const rangeBounds = normalizeRangeBounds(
      sheetName,
      entry.range.startAddress,
      entry.range.endAddress,
    );
    for (let row = rangeBounds.rowStart; row <= rangeBounds.rowEnd; row += 1) {
      for (let col = rangeBounds.colStart; col <= rangeBounds.colEnd; col += 1) {
        if (!addressWithinBounds(row, col, bounds)) {
          continue;
        }
        upsertCellSourceRow(rows, documentId, sheetName, row, col, options).styleId = entry.styleId;
      }
    }
  }

  for (const entry of sheet.metadata?.formatRanges ?? []) {
    const rangeBounds = normalizeRangeBounds(
      sheetName,
      entry.range.startAddress,
      entry.range.endAddress,
    );
    const formatCode = formatCodeById.get(entry.formatId) ?? null;
    for (let row = rangeBounds.rowStart; row <= rangeBounds.rowEnd; row += 1) {
      for (let col = rangeBounds.colStart; col <= rangeBounds.colEnd; col += 1) {
        if (!addressWithinBounds(row, col, bounds)) {
          continue;
        }
        const sourceRow = upsertCellSourceRow(rows, documentId, sheetName, row, col, options);
        sourceRow.explicitFormatId = entry.formatId;
        sourceRow.format = formatCode;
      }
    }
  }

  return [...rows.values()];
}

export function buildSheetCellSourceRowsFromEngine(
  documentId: string,
  engine: SpreadsheetEngine,
  sheetName: string,
  options: WorkbookProjectionOptions,
  range?: CellRangeRef,
): CellSourceRow[] {
  const sheet = engine.workbook.sheetsByName.get(sheetName);
  if (!sheet) {
    return [];
  }
  const bounds = range
    ? normalizeRangeBounds(range.sheetName, range.startAddress, range.endAddress)
    : undefined;
  const rows = new Map<string, CellSourceRow>();

  sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
    if (!addressWithinBounds(row, col, bounds)) {
      return;
    }
    const address = formatAddress(row, col);
    const sourceRow = buildCellSourceRowFromEngineCell(
      documentId,
      sheetName,
      address,
      row,
      col,
      engine.getCell(sheetName, address),
      options,
    );
    if (sourceRow) {
      rows.set(address, sourceRow);
    }
  });

  for (const entry of engine.workbook.listStyleRanges(sheetName)) {
    const rangeBounds = normalizeRangeBounds(
      sheetName,
      entry.range.startAddress,
      entry.range.endAddress,
    );
    for (let row = rangeBounds.rowStart; row <= rangeBounds.rowEnd; row += 1) {
      for (let col = rangeBounds.colStart; col <= rangeBounds.colEnd; col += 1) {
        if (!addressWithinBounds(row, col, bounds)) {
          continue;
        }
        upsertCellSourceRow(rows, documentId, sheetName, row, col, options).styleId = entry.styleId;
      }
    }
  }

  for (const entry of engine.workbook.listFormatRanges(sheetName)) {
    const rangeBounds = normalizeRangeBounds(
      sheetName,
      entry.range.startAddress,
      entry.range.endAddress,
    );
    const formatCode = engine.workbook.getCellNumberFormat(entry.formatId)?.code ?? null;
    for (let row = rangeBounds.rowStart; row <= rangeBounds.rowEnd; row += 1) {
      for (let col = rangeBounds.colStart; col <= rangeBounds.colEnd; col += 1) {
        if (!addressWithinBounds(row, col, bounds)) {
          continue;
        }
        const sourceRow = upsertCellSourceRow(rows, documentId, sheetName, row, col, options);
        sourceRow.explicitFormatId = entry.formatId;
        sourceRow.format = formatCode;
      }
    }
  }

  return [...rows.values()];
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
  const styles: StyleSourceRow[] = [];
  const numberFormats: NumberFormatSourceRow[] = [];
  const definedNames: DefinedNameSourceRow[] = [];
  const workbookMetadataEntries: WorkbookMetadataSourceRow[] = [];

  for (const sheet of snapshot.sheets) {
    sheets.push({
      workbookId: documentId,
      sheetId: sheet.id ?? sheet.order + 1,
      name: sheet.name,
      sortOrder: sheet.order,
      freezeRows: sheet.metadata?.freezePane?.rows ?? 0,
      freezeCols: sheet.metadata?.freezePane?.cols ?? 0,
      updatedAt: options.updatedAt,
    });

    cells.push(...buildSheetCellSourceRows(documentId, snapshot, sheet.name, options));

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
  };
}

export function buildWorkbookSourceProjectionFromEngine(
  documentId: string,
  engine: SpreadsheetEngine,
  options: WorkbookProjectionOptions,
): WorkbookSourceProjection {
  const sheets: SheetSourceRow[] = [];
  const cells: CellSourceRow[] = [];
  const rowMetadata: AxisMetadataSourceRow[] = [];
  const columnMetadata: AxisMetadataSourceRow[] = [];
  const styles = buildWorkbookStyleRowsFromEngine(documentId, engine, options);
  const numberFormats = buildWorkbookNumberFormatRowsFromEngine(documentId, engine, options);
  const definedNames = engine.getDefinedNames().map((entry) => ({
    workbookId: documentId,
    name: entry.name,
    value: entry.value,
  }));
  const workbookMetadataEntries = engine.getWorkbookMetadataEntries().map((entry) => ({
    workbookId: documentId,
    key: entry.key,
    value: entry.value,
  }));

  for (const sheet of [...engine.workbook.sheetsByName.values()].toSorted(
    (left, right) => left.order - right.order,
  )) {
    const freezePane = engine.getFreezePane(sheet.name);
    sheets.push({
      workbookId: documentId,
      sheetId: sheet.id,
      name: sheet.name,
      sortOrder: sheet.order,
      freezeRows: freezePane?.rows ?? 0,
      freezeCols: freezePane?.cols ?? 0,
      updatedAt: options.updatedAt,
    });
    cells.push(...buildSheetCellSourceRowsFromEngine(documentId, engine, sheet.name, options));
    rowMetadata.push(
      ...buildSheetRowMetadataRowsFromEngine(documentId, engine, sheet.name, options),
    );
    columnMetadata.push(
      ...buildSheetColumnMetadataRowsFromEngine(documentId, engine, sheet.name, options),
    );
  }

  return {
    workbook: buildWorkbookHeaderRowFromEngine(documentId, engine, options),
    sheets,
    cells,
    rowMetadata,
    columnMetadata,
    definedNames,
    workbookMetadataEntries,
    calculationSettings: buildCalculationSettingsRowFromEngine(documentId, engine),
    styles,
    numberFormats,
  };
}

export function materializeCellEvalProjection(
  engine: SpreadsheetEngine,
  documentId: string,
  revision: number,
  updatedAt: string,
  changedCellIndices?: readonly number[],
): CellEvalRow[] {
  const entries: CellEvalRow[] = [];

  if (changedCellIndices) {
    for (let i = 0; i < changedCellIndices.length; i += 1) {
      const cellIndex = changedCellIndices[i]!;
      const qualifiedAddress = engine.workbook.getQualifiedAddress(cellIndex);
      const separatorIndex = qualifiedAddress.lastIndexOf("!");
      if (separatorIndex <= 0 || separatorIndex >= qualifiedAddress.length - 1) {
        continue;
      }
      const sheetName = qualifiedAddress.slice(0, separatorIndex);
      const address = qualifiedAddress.slice(separatorIndex + 1);
      const { row, col } = parseCellAddress(address, sheetName);
      const cell = engine.getCell(sheetName, address);
      entries.push({
        workbookId: documentId,
        sheetName,
        address,
        rowNum: row,
        colNum: col,
        value: cell.value,
        flags: cell.flags,
        version: cell.version,
        styleId: cell.styleId ?? null,
        styleJson: engine.getCellStyle(cell.styleId) ?? null,
        formatId: cell.numberFormatId ?? null,
        formatCode: cell.format ?? null,
        calcRevision: revision,
        updatedAt,
      });
    }
    return entries;
  }

  for (const sheet of engine.workbook.sheetsByName.values()) {
    const addresses = new Set<string>();
    sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
      const address = formatAddress(row, col);
      addresses.add(address);
    });

    for (const range of engine.workbook.listStyleRanges(sheet.name)) {
      const start = parseCellAddress(range.range.startAddress, sheet.name);
      const end = parseCellAddress(range.range.endAddress, sheet.name);
      for (let row = start.row; row <= end.row; row += 1) {
        for (let col = start.col; col <= end.col; col += 1) {
          addresses.add(formatAddress(row, col));
        }
      }
    }

    for (const range of engine.workbook.listFormatRanges(sheet.name)) {
      const start = parseCellAddress(range.range.startAddress, sheet.name);
      const end = parseCellAddress(range.range.endAddress, sheet.name);
      for (let row = start.row; row <= end.row; row += 1) {
        for (let col = start.col; col <= end.col; col += 1) {
          addresses.add(formatAddress(row, col));
        }
      }
    }

    for (const address of addresses) {
      const { row, col } = parseCellAddress(address, sheet.name);
      const cell = engine.getCell(sheet.name, address);
      if (
        cell.value.tag === ValueTag.Empty &&
        cell.flags === 0 &&
        cell.styleId === undefined &&
        cell.numberFormatId === undefined &&
        cell.format === undefined
      ) {
        continue;
      }
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
        styleJson: engine.getCellStyle(cell.styleId) ?? null,
        formatId: cell.numberFormatId ?? null,
        formatCode: cell.format ?? null,
        calcRevision: revision,
        updatedAt,
      });
    }
  }

  return entries;
}

export const sourceProjectionKeys = {
  sheet: (row: SheetSourceRow) => JSON.stringify([row.workbookId, row.sheetId]),
  cell: (row: CellSourceRow) => JSON.stringify([row.workbookId, row.sheetName, row.address]),
  axisMetadata: (row: AxisMetadataSourceRow) =>
    JSON.stringify([row.workbookId, row.sheetName, row.startIndex]),
  definedName: (row: DefinedNameSourceRow) => JSON.stringify([row.workbookId, row.name]),
  workbookMetadata: (row: WorkbookMetadataSourceRow) => JSON.stringify([row.workbookId, row.key]),
  style: (row: StyleSourceRow) => JSON.stringify([row.workbookId, row.id]),
  numberFormat: (row: NumberFormatSourceRow) => JSON.stringify([row.workbookId, row.id]),
  cellEval: (row: CellEvalRow) => JSON.stringify([row.workbookId, row.sheetName, row.address]),
} as const;
