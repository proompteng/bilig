import { MAX_COLS, MAX_ROWS } from "@bilig/protocol";
import { formatAddress, parseCellAddress } from "@bilig/formula";
import { SheetGrid } from "./sheet-grid.js";
import { CellStore } from "./cell-store.js";

const SHEET_STRIDE = MAX_ROWS * MAX_COLS;

export interface SheetRecord {
  id: number;
  name: string;
  order: number;
  grid: SheetGrid;
}

export class WorkbookStore {
  readonly cellStore = new CellStore();
  readonly sheetsByName = new Map<string, SheetRecord>();
  readonly sheetsById = new Map<number, SheetRecord>();
  readonly cellKeyToIndex = new Map<number, number>();
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
    });
    this.sheetsByName.delete(name);
    this.sheetsById.delete(sheet.id);
  }

  getSheet(name: string): SheetRecord | undefined {
    return this.sheetsByName.get(name);
  }

  getOrCreateSheet(name: string): SheetRecord {
    return this.getSheet(name) ?? this.createSheet(name);
  }

  ensureCell(sheetName: string, address: string): number {
    const sheet = this.getOrCreateSheet(sheetName);
    const parsed = parseCellAddress(address, sheetName);
    const key = makeCellKey(sheet.id, parsed.row, parsed.col);
    const existing = this.cellKeyToIndex.get(key);
    if (existing !== undefined) {
      return existing;
    }
    const cellIndex = this.cellStore.allocate(sheet.id, parsed.row, parsed.col);
    this.cellKeyToIndex.set(key, cellIndex);
    sheet.grid.set(parsed.row, parsed.col, cellIndex);
    return cellIndex;
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

  reset(workbookName = "Workbook"): void {
    this.workbookName = workbookName;
    this.sheetsByName.clear();
    this.sheetsById.clear();
    this.cellKeyToIndex.clear();
    this.nextSheetId = 1;
    this.cellStore.reset();
  }
}

export function makeCellKey(sheetId: number, row: number, col: number): number {
  return sheetId * SHEET_STRIDE + row * MAX_COLS + col;
}
