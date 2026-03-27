import type {
  SheetMetadataSnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookDefinedNameSnapshot,
  WorkbookPropertySnapshot,
  WorkbookSnapshot,
} from "@bilig/protocol";
import { isLiteralInput } from "./mutators.js";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function asString(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function asNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function asBoolean(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

function asArray(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  return (
    isRecord(value) &&
    value["version"] === 1 &&
    isRecord(value["workbook"]) &&
    typeof value["workbook"]["name"] === "string" &&
    Array.isArray(value["sheets"])
  );
}

export function createEmptyWorkbookSnapshot(documentId: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: documentId,
    },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells: [],
      },
    ],
  };
}

function parseAxisMetadata(entries: unknown[]): WorkbookAxisMetadataSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null;
      }
      const start = asNumber(entry["startIndex"]);
      const count = asNumber(entry["count"]);
      if (start === undefined || count === undefined) {
        return null;
      }
      const next: WorkbookAxisMetadataSnapshot = {
        start,
        count,
      };
      const size = asNumber(entry["size"]);
      const hiddenFlag = asBoolean(entry["hidden"]);
      if (size !== undefined) {
        next.size = size;
      }
      if (hiddenFlag !== undefined) {
        next.hidden = hiddenFlag;
      }
      return next;
    })
    .filter((entry): entry is WorkbookAxisMetadataSnapshot => entry !== null);
}

function parseWorkbookProperties(entries: unknown[]): WorkbookPropertySnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null;
      }
      const key = asString(entry["key"]);
      const value = entry["value"];
      if (!key || !isLiteralInput(value)) {
        return null;
      }
      return { key, value };
    })
    .filter((entry): entry is WorkbookPropertySnapshot => entry !== null);
}

function parseDefinedNames(entries: unknown[]): WorkbookDefinedNameSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null;
      }
      const name = asString(entry["name"]);
      const value = entry["value"];
      if (!name || !isLiteralInput(value)) {
        return null;
      }
      const next: WorkbookDefinedNameSnapshot = { name, value };
      return next;
    })
    .filter((entry): entry is WorkbookDefinedNameSnapshot => entry !== null);
}

function withSheetMetadataFallback(
  fallback: SheetMetadataSnapshot | undefined,
  rowEntries: WorkbookAxisMetadataSnapshot[],
  columnEntries: WorkbookAxisMetadataSnapshot[],
): SheetMetadataSnapshot | undefined {
  const next: SheetMetadataSnapshot = {};
  if (fallback?.rows) {
    next.rows = fallback.rows;
  }
  if (fallback?.columns) {
    next.columns = fallback.columns;
  }
  if (fallback?.freezePane) {
    next.freezePane = fallback.freezePane;
  }
  if (fallback?.filters) {
    next.filters = fallback.filters;
  }
  if (fallback?.sorts) {
    next.sorts = fallback.sorts;
  }
  if (rowEntries.length > 0) {
    next.rowMetadata = rowEntries;
  } else if (fallback?.rowMetadata) {
    next.rowMetadata = fallback.rowMetadata;
  }
  if (columnEntries.length > 0) {
    next.columnMetadata = columnEntries;
  } else if (fallback?.columnMetadata) {
    next.columnMetadata = fallback.columnMetadata;
  }
  return Object.keys(next).length > 0 ? next : undefined;
}

export function projectWorkbookToSnapshot(
  value: unknown,
  documentId: string,
): WorkbookSnapshot | null {
  if (!isRecord(value)) {
    return null;
  }

  const baseSnapshot = isWorkbookSnapshot(value["snapshot"])
    ? value["snapshot"]
    : createEmptyWorkbookSnapshot(documentId);
  const workbookName = asString(value["name"]) ?? baseSnapshot.workbook.name ?? documentId;

  const workbookMetadata = parseWorkbookProperties(asArray(value["workbookMetadataEntries"]));
  const definedNames = parseDefinedNames(asArray(value["definedNames"]));
  const calculationSettingsRecord = isRecord(value["calculationSettings"])
    ? value["calculationSettings"]
    : null;
  const calculationMode = calculationSettingsRecord
    ? asString(calculationSettingsRecord["mode"])
    : undefined;
  const recalcEpoch = calculationSettingsRecord
    ? asNumber(calculationSettingsRecord["recalcEpoch"])
    : undefined;

  const fallbackSheets = new Map(baseSnapshot.sheets.map((sheet) => [sheet.name, sheet]));
  const projectedSheets = asArray(value["sheets"])
    .map((sheetEntry) => {
      if (!isRecord(sheetEntry)) {
        return null;
      }
      const sheetName = asString(sheetEntry["name"]);
      const sortOrder = asNumber(sheetEntry["sortOrder"]);
      if (!sheetName || sortOrder === undefined) {
        return null;
      }

      const cells = asArray(sheetEntry["cells"])
        .map((cellEntry) => {
          if (!isRecord(cellEntry)) {
            return null;
          }
          const address = asString(cellEntry["address"]);
          if (!address) {
            return null;
          }
          const inputValue = cellEntry["inputValue"];
          const formula = asString(cellEntry["formula"]);
          const format = asString(cellEntry["format"]);
          const nextCell: WorkbookSnapshot["sheets"][number]["cells"][number] = { address };
          if (formula) {
            nextCell.formula = formula;
          } else if (isLiteralInput(inputValue)) {
            nextCell.value = inputValue;
          }
          if (format) {
            nextCell.format = format;
          }
          return nextCell;
        })
        .filter(
          (entry): entry is WorkbookSnapshot["sheets"][number]["cells"][number] => entry !== null,
        );

      const fallbackSheet = fallbackSheets.get(sheetName);
      const metadata = withSheetMetadataFallback(
        fallbackSheet?.metadata,
        parseAxisMetadata(asArray(sheetEntry["rowMetadata"])),
        parseAxisMetadata(asArray(sheetEntry["columnMetadata"])),
      );

      return metadata
        ? { name: sheetName, order: sortOrder, metadata, cells }
        : { name: sheetName, order: sortOrder, cells };
    })
    .filter((entry): entry is WorkbookSnapshot["sheets"][number] => entry !== null);

  const workbookMetadataSnapshot = {
    ...baseSnapshot.workbook.metadata,
  };

  if (workbookMetadata.length > 0) {
    workbookMetadataSnapshot.properties = workbookMetadata;
  }
  if (definedNames.length > 0) {
    workbookMetadataSnapshot.definedNames = definedNames;
  }
  if (calculationMode === "automatic" || calculationMode === "manual") {
    workbookMetadataSnapshot.calculationSettings = {
      mode: calculationMode,
    };
  }
  if (recalcEpoch !== undefined) {
    workbookMetadataSnapshot.volatileContext = {
      recalcEpoch,
    };
  }

  const workbook =
    Object.keys(workbookMetadataSnapshot).length > 0
      ? { name: workbookName, metadata: workbookMetadataSnapshot }
      : { name: workbookName };

  return {
    version: 1,
    workbook,
    sheets: projectedSheets.length > 0 ? projectedSheets : baseSnapshot.sheets,
  };
}
