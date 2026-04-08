import { isCellRangeRef, isLiteralInput } from "@bilig/protocol";

const HORIZONTAL_ALIGNMENT_VALUES = new Set(["general", "left", "center", "right"]);
const VERTICAL_ALIGNMENT_VALUES = new Set(["top", "middle", "bottom"]);
const BORDER_STYLE_VALUES = new Set(["solid", "dashed", "dotted", "double"]);
const BORDER_WEIGHT_VALUES = new Set(["thin", "medium", "thick"]);
const NUMBER_FORMAT_KIND_VALUES = new Set([
  "general",
  "number",
  "currency",
  "accounting",
  "percent",
  "date",
  "time",
  "datetime",
  "text",
]);
const COMPATIBILITY_MODE_VALUES = new Set(["excel-modern", "odf-1.4"]);
const SORT_DIRECTION_VALUES = new Set(["asc", "desc"]);
const PIVOT_AGGREGATION_VALUES = new Set(["sum", "count"]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === "string");
}

function isOptionalString(value: unknown): value is string | undefined {
  return value === undefined || typeof value === "string";
}

function isOptionalNumber(value: unknown): value is number | undefined {
  return value === undefined || isFiniteNumber(value);
}

function isOptionalNullableNumber(value: unknown): value is number | null | undefined {
  return value === undefined || value === null || isFiniteNumber(value);
}

function isOptionalBoolean(value: unknown): value is boolean | undefined {
  return value === undefined || typeof value === "boolean";
}

function isOptionalNullableBoolean(value: unknown): value is boolean | null | undefined {
  return value === undefined || value === null || typeof value === "boolean";
}

function hasString(value: Record<string, unknown>, key: string): boolean {
  return typeof value[key] === "string";
}

function hasFiniteNumber(value: Record<string, unknown>, key: string): boolean {
  return isFiniteNumber(value[key]);
}

function isWorkbookAxisEntry(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "id") &&
    hasFiniteNumber(value, "index") &&
    isOptionalNullableNumber(value["size"]) &&
    isOptionalNullableBoolean(value["hidden"])
  );
}

function isCellBorderSide(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "color") &&
    typeof value["style"] === "string" &&
    BORDER_STYLE_VALUES.has(value["style"]) &&
    typeof value["weight"] === "string" &&
    BORDER_WEIGHT_VALUES.has(value["weight"])
  );
}

function isCellStyleRecord(value: unknown): boolean {
  if (!isRecord(value) || !hasString(value, "id")) {
    return false;
  }

  const fill = value["fill"];
  if (fill !== undefined && (!isRecord(fill) || typeof fill["backgroundColor"] !== "string")) {
    return false;
  }

  const font = value["font"];
  if (
    font !== undefined &&
    (!isRecord(font) ||
      !isOptionalString(font["family"]) ||
      !isOptionalNumber(font["size"]) ||
      !isOptionalBoolean(font["bold"]) ||
      !isOptionalBoolean(font["italic"]) ||
      !isOptionalBoolean(font["underline"]) ||
      !isOptionalString(font["color"]))
  ) {
    return false;
  }

  const alignment = value["alignment"];
  if (
    alignment !== undefined &&
    (!isRecord(alignment) ||
      !(
        alignment["horizontal"] === undefined ||
        (typeof alignment["horizontal"] === "string" &&
          HORIZONTAL_ALIGNMENT_VALUES.has(alignment["horizontal"]))
      ) ||
      !(
        alignment["vertical"] === undefined ||
        (typeof alignment["vertical"] === "string" &&
          VERTICAL_ALIGNMENT_VALUES.has(alignment["vertical"]))
      ) ||
      !isOptionalBoolean(alignment["wrap"]) ||
      !isOptionalNumber(alignment["indent"]))
  ) {
    return false;
  }

  const borders = value["borders"];
  if (
    borders !== undefined &&
    (!isRecord(borders) ||
      !(borders["top"] === undefined || isCellBorderSide(borders["top"])) ||
      !(borders["right"] === undefined || isCellBorderSide(borders["right"])) ||
      !(borders["bottom"] === undefined || isCellBorderSide(borders["bottom"])) ||
      !(borders["left"] === undefined || isCellBorderSide(borders["left"])))
  ) {
    return false;
  }

  return true;
}

function isCellNumberFormatRecord(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "id") &&
    hasString(value, "code") &&
    typeof value["kind"] === "string" &&
    NUMBER_FORMAT_KIND_VALUES.has(value["kind"])
  );
}

function isWorkbookCalculationSettings(value: unknown): boolean {
  return (
    isRecord(value) &&
    (value["mode"] === "automatic" || value["mode"] === "manual") &&
    (value["compatibilityMode"] === undefined ||
      (typeof value["compatibilityMode"] === "string" &&
        COMPATIBILITY_MODE_VALUES.has(value["compatibilityMode"])))
  );
}

function isWorkbookVolatileContext(value: unknown): boolean {
  return isRecord(value) && hasFiniteNumber(value, "recalcEpoch");
}

function isWorkbookSortKey(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "keyAddress") &&
    typeof value["direction"] === "string" &&
    SORT_DIRECTION_VALUES.has(value["direction"])
  );
}

function isWorkbookDefinedNameValue(value: unknown): boolean {
  if (isLiteralInput(value)) {
    return true;
  }

  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }

  switch (value["kind"]) {
    case "scalar":
      return isLiteralInput(value["value"]);
    case "cell-ref":
      return hasString(value, "sheetName") && hasString(value, "address");
    case "range-ref":
      return isCellRangeRef(value);
    case "structured-ref":
      return hasString(value, "tableName") && hasString(value, "columnName");
    case "formula":
      return hasString(value, "formula");
    default:
      return false;
  }
}

function isWorkbookTableOp(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "name") &&
    hasString(value, "sheetName") &&
    hasString(value, "startAddress") &&
    hasString(value, "endAddress") &&
    isStringArray(value["columnNames"]) &&
    typeof value["headerRow"] === "boolean" &&
    typeof value["totalsRow"] === "boolean"
  );
}

function isWorkbookPivotValue(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "sourceColumn") &&
    typeof value["summarizeBy"] === "string" &&
    PIVOT_AGGREGATION_VALUES.has(value["summarizeBy"]) &&
    isOptionalString(value["outputLabel"])
  );
}

export function isWorkbookOp(value: unknown): value is import("./index.js").WorkbookOp {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }

  switch (value["kind"]) {
    case "upsertWorkbook":
      return hasString(value, "name");
    case "setWorkbookMetadata":
      return hasString(value, "key") && isLiteralInput(value["value"]);
    case "setCalculationSettings":
      return isWorkbookCalculationSettings(value["settings"]);
    case "setVolatileContext":
      return isWorkbookVolatileContext(value["context"]);
    case "upsertSheet":
      return (
        hasString(value, "name") && hasFiniteNumber(value, "order") && isOptionalNumber(value["id"])
      );
    case "renameSheet":
      return hasString(value, "oldName") && hasString(value, "newName");
    case "deleteSheet":
      return hasString(value, "name");
    case "insertRows":
    case "insertColumns":
      return (
        hasString(value, "sheetName") &&
        hasFiniteNumber(value, "start") &&
        hasFiniteNumber(value, "count") &&
        (value["entries"] === undefined ||
          (Array.isArray(value["entries"]) &&
            value["entries"].every((entry) => isWorkbookAxisEntry(entry))))
      );
    case "deleteRows":
    case "deleteColumns":
      return (
        hasString(value, "sheetName") &&
        hasFiniteNumber(value, "start") &&
        hasFiniteNumber(value, "count")
      );
    case "moveRows":
    case "moveColumns":
      return (
        hasString(value, "sheetName") &&
        hasFiniteNumber(value, "start") &&
        hasFiniteNumber(value, "count") &&
        hasFiniteNumber(value, "target")
      );
    case "updateRowMetadata":
    case "updateColumnMetadata":
      return (
        hasString(value, "sheetName") &&
        hasFiniteNumber(value, "start") &&
        hasFiniteNumber(value, "count") &&
        isOptionalNullableNumber(value["size"]) &&
        isOptionalNullableBoolean(value["hidden"])
      );
    case "setFreezePane":
      return (
        hasString(value, "sheetName") &&
        hasFiniteNumber(value, "rows") &&
        hasFiniteNumber(value, "cols")
      );
    case "clearFreezePane":
      return hasString(value, "sheetName");
    case "setFilter":
    case "clearFilter":
    case "clearSort":
      return hasString(value, "sheetName") && isCellRangeRef(value["range"]);
    case "setSort":
      return (
        hasString(value, "sheetName") &&
        isCellRangeRef(value["range"]) &&
        Array.isArray(value["keys"]) &&
        value["keys"].every((entry) => isWorkbookSortKey(entry))
      );
    case "setCellValue":
      return (
        hasString(value, "sheetName") &&
        hasString(value, "address") &&
        isLiteralInput(value["value"])
      );
    case "setCellFormula":
      return (
        hasString(value, "sheetName") && hasString(value, "address") && hasString(value, "formula")
      );
    case "setCellFormat":
      return (
        hasString(value, "sheetName") &&
        hasString(value, "address") &&
        (value["format"] === null || typeof value["format"] === "string")
      );
    case "upsertCellStyle":
      return isCellStyleRecord(value["style"]);
    case "upsertCellNumberFormat":
      return isCellNumberFormatRecord(value["format"]);
    case "setStyleRange":
      return isCellRangeRef(value["range"]) && hasString(value, "styleId");
    case "setFormatRange":
      return isCellRangeRef(value["range"]) && hasString(value, "formatId");
    case "clearCell":
      return hasString(value, "sheetName") && hasString(value, "address");
    case "upsertDefinedName":
      return hasString(value, "name") && isWorkbookDefinedNameValue(value["value"]);
    case "deleteDefinedName":
    case "deleteTable":
      return hasString(value, "name");
    case "upsertTable":
      return isWorkbookTableOp(value["table"]);
    case "upsertSpillRange":
      return (
        hasString(value, "sheetName") &&
        hasString(value, "address") &&
        hasFiniteNumber(value, "rows") &&
        hasFiniteNumber(value, "cols")
      );
    case "deleteSpillRange":
    case "deletePivotTable":
      return hasString(value, "sheetName") && hasString(value, "address");
    case "upsertPivotTable":
      return (
        hasString(value, "name") &&
        hasString(value, "sheetName") &&
        hasString(value, "address") &&
        isCellRangeRef(value["source"]) &&
        isStringArray(value["groupBy"]) &&
        Array.isArray(value["values"]) &&
        value["values"].every((entry) => isWorkbookPivotValue(entry)) &&
        hasFiniteNumber(value, "rows") &&
        hasFiniteNumber(value, "cols")
      );
    default:
      return false;
  }
}

export function isEngineOp(value: unknown): value is import("./index.js").EngineOp {
  return isWorkbookOp(value);
}

export function isEngineOps(value: unknown): value is import("./index.js").EngineOp[] {
  return Array.isArray(value) && value.every((entry) => isEngineOp(entry));
}

export function isEngineOpBatch(value: unknown): value is import("./index.js").EngineOpBatch {
  return (
    isRecord(value) &&
    hasString(value, "id") &&
    hasString(value, "replicaId") &&
    isRecord(value["clock"]) &&
    hasFiniteNumber(value["clock"], "counter") &&
    isEngineOps(value["ops"])
  );
}
