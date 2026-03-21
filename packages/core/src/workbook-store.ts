import {
  MAX_COLS,
  MAX_ROWS,
  type LiteralInput,
  type WorkbookPivotSnapshot,
  type WorkbookPivotValueSnapshot
} from "@bilig/protocol";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import { SheetGrid } from "./sheet-grid.js";
import { CellStore } from "./cell-store.js";

const SHEET_STRIDE = MAX_ROWS * MAX_COLS;

export interface WorkbookDefinedNameRecord {
  name: string;
  value: LiteralInput;
}

export interface WorkbookSpillRecord {
  sheetName: string;
  address: string;
  rows: number;
  cols: number;
}

export interface WorkbookPivotRecord extends WorkbookPivotSnapshot {
  values: WorkbookPivotValueSnapshot[];
}

export interface WorkbookMetadataRecord {
  definedNames: Map<string, WorkbookDefinedNameRecord>;
  spills: Map<string, WorkbookSpillRecord>;
  pivots: Map<string, WorkbookPivotRecord>;
}

export function normalizeDefinedName(name: string): string {
  const normalized = name.trim().toUpperCase();
  if (normalized.length === 0) {
    throw new Error("Defined names must be non-empty");
  }
  return normalized;
}

export interface SheetRecord {
  id: number;
  name: string;
  order: number;
  grid: SheetGrid;
}

export interface EnsuredCell {
  cellIndex: number;
  created: boolean;
}

export class WorkbookStore {
  readonly cellStore = new CellStore();
  readonly sheetsByName = new Map<string, SheetRecord>();
  readonly sheetsById = new Map<number, SheetRecord>();
  readonly cellKeyToIndex = new Map<number, number>();
  readonly cellFormats = new Map<number, string>();
  readonly metadata: WorkbookMetadataRecord = {
    definedNames: new Map(),
    spills: new Map(),
    pivots: new Map()
  };
  workbookName: string;
  private nextSheetId = 1;

  constructor(workbookName = "Workbook") {
    this.workbookName = workbookName;
  }

  createSheet(name: string, order = this.sheetsByName.size): SheetRecord {
    const existing = this.sheetsByName.get(name);
    if (existing) {
      existing.order = order;
      return existing;
    }
    const sheet: SheetRecord = {
      id: this.nextSheetId++,
      name,
      order,
      grid: new SheetGrid()
    };
    this.sheetsByName.set(name, sheet);
    this.sheetsById.set(sheet.id, sheet);
    return sheet;
  }

  deleteSheet(name: string): void {
    const sheet = this.sheetsByName.get(name);
    if (!sheet) return;
    sheet.grid.forEachCell((cellIndex) => {
      const key = makeCellKey(sheet.id, this.cellStore.rows[cellIndex]!, this.cellStore.cols[cellIndex]!);
      this.cellKeyToIndex.delete(key);
      this.cellFormats.delete(cellIndex);
    });
    for (const [key, spill] of this.metadata.spills.entries()) {
      if (spill.sheetName === name) {
        this.metadata.spills.delete(key);
      }
    }
    for (const [key, pivot] of this.metadata.pivots.entries()) {
      if (pivot.sheetName === name) {
        this.metadata.pivots.delete(key);
      }
    }
    this.sheetsByName.delete(name);
    this.sheetsById.delete(sheet.id);
  }

  getSheet(name: string): SheetRecord | undefined {
    return this.sheetsByName.get(name);
  }

  getSheetById(id: number): SheetRecord | undefined {
    return this.sheetsById.get(id);
  }

  getOrCreateSheet(name: string): SheetRecord {
    return this.getSheet(name) ?? this.createSheet(name);
  }

  ensureCell(sheetName: string, address: string): number {
    return this.ensureCellRecord(sheetName, address).cellIndex;
  }

  ensureCellRecord(sheetName: string, address: string): EnsuredCell {
    const sheet = this.getOrCreateSheet(sheetName);
    const parsed = parseCellAddress(address, sheetName);
    return this.ensureCellAt(sheet.id, parsed.row, parsed.col);
  }

  ensureCellAt(sheetId: number, row: number, col: number): EnsuredCell {
    const sheet = this.getSheetById(sheetId);
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`);
    }
    const key = makeCellKey(sheet.id, row, col);
    const existing = this.cellKeyToIndex.get(key);
    if (existing !== undefined) {
      return { cellIndex: existing, created: false };
    }
    const cellIndex = this.cellStore.allocate(sheet.id, row, col);
    this.cellKeyToIndex.set(key, cellIndex);
    sheet.grid.set(row, col, cellIndex);
    return { cellIndex, created: true };
  }

  getCellIndex(sheetName: string, address: string): number | undefined {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return undefined;
    const parsed = parseCellAddress(address, sheetName);
    return this.cellKeyToIndex.get(makeCellKey(sheet.id, parsed.row, parsed.col));
  }

  getSheetNameById(id: number): string {
    return this.sheetsById.get(id)?.name ?? "";
  }

  getAddress(index: number): string {
    return formatAddress(this.cellStore.rows[index]!, this.cellStore.cols[index]!);
  }

  getQualifiedAddress(index: number): string {
    return `${this.getSheetNameById(this.cellStore.sheetIds[index]!)}!${this.getAddress(index)}`;
  }

  setCellFormat(index: number, format: string | null | undefined): void {
    if (format === undefined || format === null || format === "") {
      this.cellFormats.delete(index);
      return;
    }
    this.cellFormats.set(index, format);
  }

  getCellFormat(index: number): string | undefined {
    return this.cellFormats.get(index);
  }

  setDefinedName(name: string, value: LiteralInput): WorkbookDefinedNameRecord {
    const trimmedName = name.trim();
    const record: WorkbookDefinedNameRecord = { name: trimmedName, value };
    this.metadata.definedNames.set(normalizeDefinedName(trimmedName), record);
    return record;
  }

  getDefinedName(name: string): WorkbookDefinedNameRecord | undefined {
    return this.metadata.definedNames.get(normalizeDefinedName(name));
  }

  deleteDefinedName(name: string): boolean {
    return this.metadata.definedNames.delete(normalizeDefinedName(name));
  }

  listDefinedNames(): WorkbookDefinedNameRecord[] {
    return [...this.metadata.definedNames.values()].sort((left, right) =>
      normalizeDefinedName(left.name).localeCompare(normalizeDefinedName(right.name))
    );
  }

  setSpill(sheetName: string, address: string, rows: number, cols: number): WorkbookSpillRecord {
    const record: WorkbookSpillRecord = { sheetName, address, rows, cols };
    this.metadata.spills.set(spillKey(sheetName, address), record);
    return record;
  }

  getSpill(sheetName: string, address: string): WorkbookSpillRecord | undefined {
    return this.metadata.spills.get(spillKey(sheetName, address));
  }

  deleteSpill(sheetName: string, address: string): boolean {
    return this.metadata.spills.delete(spillKey(sheetName, address));
  }

  listSpills(): WorkbookSpillRecord[] {
    return [...this.metadata.spills.values()].sort((left, right) =>
      `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`)
    );
  }

  setPivot(record: WorkbookPivotSnapshot): WorkbookPivotRecord {
    const stored: WorkbookPivotRecord = {
      ...record,
      name: record.name.trim(),
      groupBy: [...record.groupBy],
      values: record.values.map((value) => ({ ...value })),
      source: { ...record.source }
    };
    this.metadata.pivots.set(pivotKey(record.sheetName, record.address), stored);
    return stored;
  }

  getPivot(sheetName: string, address: string): WorkbookPivotRecord | undefined {
    return this.metadata.pivots.get(pivotKey(sheetName, address));
  }

  getPivotByKey(key: string): WorkbookPivotRecord | undefined {
    return this.metadata.pivots.get(key);
  }

  deletePivot(sheetName: string, address: string): boolean {
    return this.metadata.pivots.delete(pivotKey(sheetName, address));
  }

  listPivots(): WorkbookPivotRecord[] {
    return [...this.metadata.pivots.values()].sort((left, right) =>
      `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`)
    );
  }

  reset(workbookName = "Workbook"): void {
    this.workbookName = workbookName;
    this.sheetsByName.clear();
    this.sheetsById.clear();
    this.cellKeyToIndex.clear();
    this.cellFormats.clear();
    this.metadata.definedNames.clear();
    this.metadata.spills.clear();
    this.metadata.pivots.clear();
    this.nextSheetId = 1;
    this.cellStore.reset();
  }
}

export function makeCellKey(sheetId: number, row: number, col: number): number {
  return sheetId * SHEET_STRIDE + row * MAX_COLS + col;
}

function spillKey(sheetName: string, address: string): string {
  return `${sheetName}!${address}`;
}

export function pivotKey(sheetName: string, address: string): string {
  return `${sheetName}!${address}`;
}
