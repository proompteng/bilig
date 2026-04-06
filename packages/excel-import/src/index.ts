import * as XLSX from "xlsx";

import type { WorkbookAxisEntrySnapshot, WorkbookSnapshot } from "@bilig/protocol";

export interface ImportedWorkbook {
  snapshot: WorkbookSnapshot;
  workbookName: string;
  sheetNames: string[];
  warnings: string[];
}

interface SheetColumnInfo {
  index: number;
  size: number;
}

interface SheetRowInfo {
  index: number;
  size: number;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function normalizeWorkbookName(fileName: string): string {
  const trimmed = fileName.trim();
  if (trimmed.length === 0) {
    return "Imported workbook";
  }
  return trimmed.replace(/\.xlsx$/i, "") || "Imported workbook";
}

function toLiteralInput(value: unknown) {
  if (value === null || value === undefined) {
    return undefined;
  }
  if (typeof value === "number" || typeof value === "string" || typeof value === "boolean") {
    return value;
  }
  if (value instanceof Date) {
    return value.getTime();
  }
  return undefined;
}

function toPixelSize(value: number | undefined, unit: "pt" | "ch"): number | null {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    return null;
  }
  if (unit === "pt") {
    return Math.round((value * 96) / 72);
  }
  return Math.round(value * 8 + 5);
}

function buildColumnEntries(
  columns: unknown[] | undefined,
): WorkbookAxisEntrySnapshot[] | undefined {
  if (!Array.isArray(columns) || columns.length === 0) {
    return undefined;
  }
  const entries: SheetColumnInfo[] = [];
  columns.forEach((entry, index) => {
    if (!isRecord(entry)) {
      return;
    }
    const size =
      typeof entry["wpx"] === "number"
        ? Math.round(entry["wpx"])
        : typeof entry["wch"] === "number"
          ? toPixelSize(entry["wch"], "ch")
          : null;
    if (size === null) {
      return;
    }
    entries.push({ index, size });
  });
  if (entries.length === 0) {
    return undefined;
  }
  return entries.map(({ index, size }) => ({
    id: `col:${index}`,
    index,
    size,
  }));
}

function buildRowEntries(rows: unknown[] | undefined): WorkbookAxisEntrySnapshot[] | undefined {
  if (!Array.isArray(rows) || rows.length === 0) {
    return undefined;
  }
  const entries: SheetRowInfo[] = [];
  rows.forEach((entry, index) => {
    if (!isRecord(entry)) {
      return;
    }
    const size =
      typeof entry["hpx"] === "number"
        ? Math.round(entry["hpx"])
        : typeof entry["hpt"] === "number"
          ? toPixelSize(entry["hpt"], "pt")
          : null;
    if (size === null) {
      return;
    }
    entries.push({ index, size });
  });
  if (entries.length === 0) {
    return undefined;
  }
  return entries.map(({ index, size }) => ({
    id: `row:${index}`,
    index,
    size,
  }));
}

function addWorkbookWarnings(workbook: XLSX.WorkBook, warnings: string[]): void {
  if (workbook.vbaraw) {
    warnings.push("Macros were ignored during XLSX import.");
  }
  const definedNames = workbook.Workbook?.Names;
  if (Array.isArray(definedNames) && definedNames.length > 0) {
    warnings.push("Defined names were ignored during XLSX import.");
  }
}

function addSheetWarnings(
  sheetName: string,
  sheet: XLSX.WorkSheet,
  warnings: string[],
  ignoredComments: { seen: boolean },
): void {
  const merges = sheet["!merges"];
  if (Array.isArray(merges) && merges.length > 0) {
    warnings.push(`Merged cells on ${sheetName} were ignored during XLSX import.`);
  }
  Object.values(sheet).forEach((value) => {
    if (!isRecord(value)) {
      return;
    }
    if (!ignoredComments.seen && Array.isArray(value["c"]) && value["c"].length > 0) {
      ignoredComments.seen = true;
      warnings.push("Cell comments were ignored during XLSX import.");
    }
  });
}

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const workbook = XLSX.read(data, {
    type: "array",
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    cellText: false,
    cellDates: false,
  });
  const workbookName = normalizeWorkbookName(fileName);
  const warnings: string[] = [];
  addWorkbookWarnings(workbook, warnings);

  const ignoredComments = { seen: false };
  const sheets = workbook.SheetNames.map((sheetName, order) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      return {
        id: order + 1,
        name: sheetName,
        order,
        cells: [],
      };
    }

    addSheetWarnings(sheetName, sheet, warnings, ignoredComments);
    const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null;
    const cells: WorkbookSnapshot["sheets"][number]["cells"] = [];
    if (range) {
      for (let row = range.s.r; row <= range.e.r; row += 1) {
        for (let col = range.s.c; col <= range.e.c; col += 1) {
          const address = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = sheet[address];
          if (!cell) {
            continue;
          }
          const nextCell: WorkbookSnapshot["sheets"][number]["cells"][number] = { address };
          if (typeof cell.f === "string" && cell.f.trim().length > 0) {
            nextCell.formula = cell.f;
          } else {
            const literal = toLiteralInput(cell.v);
            if (literal !== undefined) {
              nextCell.value = literal;
            }
          }
          if (typeof cell.z === "string" && cell.z.trim().length > 0) {
            nextCell.format = cell.z;
          }
          if (
            nextCell.value !== undefined ||
            nextCell.formula !== undefined ||
            nextCell.format !== undefined
          ) {
            cells.push(nextCell);
          }
        }
      }
    }

    const rows = buildRowEntries(sheet["!rows"]);
    const columns = buildColumnEntries(sheet["!cols"]);
    const metadata =
      rows || columns
        ? {
            ...(rows ? { rows } : {}),
            ...(columns ? { columns } : {}),
          }
        : undefined;

    return {
      id: order + 1,
      name: sheetName,
      order,
      ...(metadata ? { metadata } : {}),
      cells,
    };
  });

  return {
    snapshot: {
      version: 1,
      workbook: {
        name: workbookName,
      },
      sheets,
    },
    workbookName,
    sheetNames: workbook.SheetNames,
    warnings,
  };
}
