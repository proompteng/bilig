import { formatAddress } from "@bilig/formula";
import type {
  CellNumberFormatRecord,
  CellStyleRecord,
  CompatibilityMode,
  SheetFormatRangeSnapshot,
  SheetMetadataSnapshot,
  SheetStyleRangeSnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookDefinedNameSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookFreezePaneSnapshot,
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

function isCellNumberFormatKind(value: unknown): value is CellNumberFormatRecord["kind"] {
  return (
    value === "general" ||
    value === "number" ||
    value === "currency" ||
    value === "accounting" ||
    value === "percent" ||
    value === "date" ||
    value === "time" ||
    value === "datetime" ||
    value === "text"
  );
}

function isCompatibilityMode(value: unknown): value is CompatibilityMode {
  return value === "excel-modern" || value === "odf-1.4";
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
        id: 1,
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

function isWorkbookDefinedNameValueSnapshot(
  value: unknown,
): value is WorkbookDefinedNameValueSnapshot {
  return (
    value === null ||
    typeof value === "number" ||
    typeof value === "string" ||
    typeof value === "boolean" ||
    isRecord(value)
  );
}

function parseDefinedNames(entries: unknown[]): WorkbookDefinedNameSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null;
      }
      const name = asString(entry["name"]);
      const value = entry["value"];
      if (!name || !isWorkbookDefinedNameValueSnapshot(value)) {
        return null;
      }
      return { name, value };
    })
    .filter((entry): entry is WorkbookDefinedNameSnapshot => entry !== null);
}

function parseStyleRecords(entries: unknown[]): CellStyleRecord[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null;
      }
      const id = asString(entry["id"]);
      const recordJSON = entry["recordJSON"];
      if (!id || !isRecord(recordJSON)) {
        return null;
      }
      return {
        ...(recordJSON as Omit<CellStyleRecord, "id">),
        id,
      };
    })
    .filter((entry): entry is CellStyleRecord => entry !== null);
}

function parseNumberFormats(entries: unknown[]): CellNumberFormatRecord[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null;
      }
      const id = asString(entry["id"]);
      const code = asString(entry["code"]);
      const kind = asString(entry["kind"]);
      if (!id || !code || !isCellNumberFormatKind(kind)) {
        return null;
      }
      return {
        id,
        code,
        kind,
      };
    })
    .filter((entry): entry is CellNumberFormatRecord => entry !== null);
}

function parseFreezePane(
  freezeRows: unknown,
  freezeCols: unknown,
  fallback?: WorkbookFreezePaneSnapshot,
): WorkbookFreezePaneSnapshot | undefined {
  const rows = asNumber(freezeRows);
  const cols = asNumber(freezeCols);
  if ((rows ?? 0) > 0 || (cols ?? 0) > 0) {
    return {
      rows: rows ?? 0,
      cols: cols ?? 0,
    };
  }
  return fallback;
}

function parseStyleRanges(entries: unknown[]): SheetStyleRangeSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null;
      }
      const startRow = asNumber(entry["startRow"]);
      const endRow = asNumber(entry["endRow"]);
      const startCol = asNumber(entry["startCol"]);
      const endCol = asNumber(entry["endCol"]);
      const styleId = asString(entry["styleId"]);
      if (
        startRow === undefined ||
        endRow === undefined ||
        startCol === undefined ||
        endCol === undefined ||
        !styleId
      ) {
        return null;
      }
      return {
        range: {
          sheetName: "",
          startAddress: formatAddress(startRow, startCol),
          endAddress: formatAddress(endRow, endCol),
        },
        styleId,
      };
    })
    .filter((entry): entry is SheetStyleRangeSnapshot => entry !== null);
}

function parseFormatRanges(entries: unknown[]): SheetFormatRangeSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null;
      }
      const startRow = asNumber(entry["startRow"]);
      const endRow = asNumber(entry["endRow"]);
      const startCol = asNumber(entry["startCol"]);
      const endCol = asNumber(entry["endCol"]);
      const formatId = asString(entry["formatId"]);
      if (
        startRow === undefined ||
        endRow === undefined ||
        startCol === undefined ||
        endCol === undefined ||
        !formatId
      ) {
        return null;
      }
      return {
        range: {
          sheetName: "",
          startAddress: formatAddress(startRow, startCol),
          endAddress: formatAddress(endRow, endCol),
        },
        formatId,
      };
    })
    .filter((entry): entry is SheetFormatRangeSnapshot => entry !== null);
}

function withSheetMetadataFallback(
  sheetName: string,
  rowEntries: WorkbookAxisMetadataSnapshot[],
  columnEntries: WorkbookAxisMetadataSnapshot[],
  styleRanges: SheetStyleRangeSnapshot[],
  formatRanges: SheetFormatRangeSnapshot[],
  freezePane: WorkbookFreezePaneSnapshot | undefined,
  fallback?: SheetMetadataSnapshot,
) {
  const next: SheetMetadataSnapshot = {};
  if (fallback?.rows) {
    next.rows = fallback.rows;
  }
  if (fallback?.columns) {
    next.columns = fallback.columns;
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
  if (styleRanges.length > 0) {
    next.styleRanges = styleRanges.map((entry) => ({
      ...entry,
      range: {
        ...entry.range,
        sheetName,
      },
    }));
  } else if (fallback?.styleRanges) {
    next.styleRanges = fallback.styleRanges;
  }
  if (formatRanges.length > 0) {
    next.formatRanges = formatRanges.map((entry) => ({
      ...entry,
      range: {
        ...entry.range,
        sheetName,
      },
    }));
  } else if (fallback?.formatRanges) {
    next.formatRanges = fallback.formatRanges;
  }
  if (freezePane) {
    next.freezePane = freezePane;
  } else if (fallback?.freezePane) {
    next.freezePane = fallback.freezePane;
  }
  return Object.keys(next).length > 0 ? next : undefined;
}

export function projectWorkbookToSnapshot(value: unknown, documentId: string) {
  if (!isRecord(value)) {
    return null;
  }

  const baseSnapshot = isWorkbookSnapshot(value["snapshot"])
    ? value["snapshot"]
    : createEmptyWorkbookSnapshot(documentId);
  const workbookName = asString(value["name"]) ?? baseSnapshot.workbook.name ?? documentId;

  const workbookMetadata = parseWorkbookProperties(asArray(value["workbookMetadataEntries"]));
  const definedNames = parseDefinedNames(asArray(value["definedNames"]));
  const styles = parseStyleRecords(asArray(value["styles"]));
  const numberFormats = parseNumberFormats(asArray(value["numberFormats"]));
  const numberFormatCodeById = new Map(numberFormats.map((entry) => [entry.id, entry.code]));

  const calculationSettingsRecord = isRecord(value["calculationSettings"])
    ? value["calculationSettings"]
    : null;
  const calculationMode = calculationSettingsRecord
    ? asString(calculationSettingsRecord["mode"])
    : undefined;
  const compatibilityMode = asString(value["compatibilityMode"]);
  const recalcEpoch =
    calculationSettingsRecord?.["recalcEpoch"] !== undefined
      ? asNumber(calculationSettingsRecord["recalcEpoch"])
      : asNumber(value["recalcEpoch"]);

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
          const explicitFormatId = asString(cellEntry["explicitFormatId"]);
          const address =
            asString(cellEntry["address"]) ??
            (asNumber(cellEntry["rowNum"]) !== undefined &&
            asNumber(cellEntry["colNum"]) !== undefined
              ? formatAddress(
                  asNumber(cellEntry["rowNum"]) ?? 0,
                  asNumber(cellEntry["colNum"]) ?? 0,
                )
              : undefined);
          if (!address) {
            return null;
          }
          const inputValue = cellEntry["inputValue"];
          const formula = asString(cellEntry["formula"]);
          const format =
            asString(cellEntry["format"]) ??
            (explicitFormatId ? numberFormatCodeById.get(explicitFormatId) : undefined);
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
        sheetName,
        parseAxisMetadata(asArray(sheetEntry["rowMetadata"])),
        parseAxisMetadata(asArray(sheetEntry["columnMetadata"])),
        parseStyleRanges(asArray(sheetEntry["styleRanges"])),
        parseFormatRanges(asArray(sheetEntry["formatRanges"])),
        parseFreezePane(
          sheetEntry["freezeRows"],
          sheetEntry["freezeCols"],
          fallbackSheet?.metadata?.freezePane,
        ),
        fallbackSheet?.metadata,
      );

      const id = asNumber(sheetEntry["id"]) ?? fallbackSheet?.id;
      const nextSheet: WorkbookSnapshot["sheets"][number] = metadata
        ? { name: sheetName, order: sortOrder, metadata, cells }
        : { name: sheetName, order: sortOrder, cells };
      if (id !== undefined) {
        nextSheet.id = id;
      }
      return nextSheet;
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
  if (styles.length > 0) {
    workbookMetadataSnapshot.styles = styles;
  }
  if (numberFormats.length > 0) {
    workbookMetadataSnapshot.formats = numberFormats;
  }
  if (
    (calculationMode === "automatic" || calculationMode === "manual") &&
    isCompatibilityMode(compatibilityMode)
  ) {
    workbookMetadataSnapshot.calculationSettings = {
      mode: calculationMode,
      compatibilityMode,
    };
  } else if (calculationMode === "automatic" || calculationMode === "manual") {
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
