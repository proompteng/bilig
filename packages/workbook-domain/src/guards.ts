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
const VALIDATION_COMPARISON_OPERATOR_VALUES = new Set([
  "between",
  "notBetween",
  "equal",
  "notEqual",
  "greaterThan",
  "greaterThanOrEqual",
  "lessThan",
  "lessThanOrEqual",
]);
const VALIDATION_ERROR_STYLE_VALUES = new Set(["stop", "warning", "information"]);

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

function isOptionalLiteralInput(value: unknown): boolean {
  return value === undefined || isLiteralInput(value);
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

function isWorkbookValidationListSource(value: unknown): boolean {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  switch (value["kind"]) {
    case "named-range":
      return hasString(value, "name");
    case "cell-ref":
      return hasString(value, "sheetName") && hasString(value, "address");
    case "range-ref":
      return isCellRangeRef(value);
    case "structured-ref":
      return hasString(value, "tableName") && hasString(value, "columnName");
    default:
      return false;
  }
}

function isWorkbookDataValidationRule(value: unknown): boolean {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  switch (value["kind"]) {
    case "list": {
      const hasValues =
        Array.isArray(value["values"]) && value["values"].every((entry) => isLiteralInput(entry));
      const hasSource =
        value["source"] !== undefined && isWorkbookValidationListSource(value["source"]);
      return (hasValues ? 1 : 0) + (hasSource ? 1 : 0) === 1;
    }
    case "checkbox":
      return (
        isOptionalLiteralInput(value["checkedValue"]) &&
        isOptionalLiteralInput(value["uncheckedValue"])
      );
    case "whole":
    case "decimal":
    case "date":
    case "time":
    case "textLength":
      return (
        typeof value["operator"] === "string" &&
        VALIDATION_COMPARISON_OPERATOR_VALUES.has(value["operator"]) &&
        Array.isArray(value["values"]) &&
        value["values"].every((entry) => isLiteralInput(entry)) &&
        (value["operator"] === "between" || value["operator"] === "notBetween"
          ? value["values"].length === 2
          : value["values"].length === 1)
      );
    default:
      return false;
  }
}

function isWorkbookDataValidation(value: unknown): boolean {
  return (
    isRecord(value) &&
    isCellRangeRef(value["range"]) &&
    isWorkbookDataValidationRule(value["rule"]) &&
    isOptionalBoolean(value["allowBlank"]) &&
    isOptionalBoolean(value["showDropdown"]) &&
    isOptionalString(value["promptTitle"]) &&
    isOptionalString(value["promptMessage"]) &&
    (value["errorStyle"] === undefined ||
      (typeof value["errorStyle"] === "string" &&
        VALIDATION_ERROR_STYLE_VALUES.has(value["errorStyle"]))) &&
    isOptionalString(value["errorTitle"]) &&
    isOptionalString(value["errorMessage"])
  );
}

function isCellStylePatch(value: unknown): boolean {
  if (!isRecord(value)) {
    return false;
  }

  const fill = value["fill"];
  if (
    fill !== undefined &&
    fill !== null &&
    (!isRecord(fill) ||
      !(
        fill["backgroundColor"] === undefined ||
        fill["backgroundColor"] === null ||
        hasString(fill, "backgroundColor")
      ))
  ) {
    return false;
  }

  const font = value["font"];
  if (
    font !== undefined &&
    font !== null &&
    (!isRecord(font) ||
      !(font["family"] === undefined || font["family"] === null || hasString(font, "family")) ||
      !isOptionalNullableNumber(font["size"]) ||
      !isOptionalNullableBoolean(font["bold"]) ||
      !isOptionalNullableBoolean(font["italic"]) ||
      !isOptionalNullableBoolean(font["underline"]) ||
      !(font["color"] === undefined || font["color"] === null || hasString(font, "color")))
  ) {
    return false;
  }

  const alignment = value["alignment"];
  if (
    alignment !== undefined &&
    alignment !== null &&
    (!isRecord(alignment) ||
      !(
        alignment["horizontal"] === undefined ||
        alignment["horizontal"] === null ||
        (typeof alignment["horizontal"] === "string" &&
          HORIZONTAL_ALIGNMENT_VALUES.has(alignment["horizontal"]))
      ) ||
      !(
        alignment["vertical"] === undefined ||
        alignment["vertical"] === null ||
        (typeof alignment["vertical"] === "string" &&
          VERTICAL_ALIGNMENT_VALUES.has(alignment["vertical"]))
      ) ||
      !isOptionalNullableBoolean(alignment["wrap"]) ||
      !isOptionalNullableNumber(alignment["indent"]))
  ) {
    return false;
  }

  const borders = value["borders"];
  if (borders !== undefined && borders !== null) {
    if (!isRecord(borders)) {
      return false;
    }
    for (const side of ["top", "right", "bottom", "left"] as const) {
      const sideValue = borders[side];
      if (sideValue === undefined || sideValue === null) {
        continue;
      }
      if (
        !isRecord(sideValue) ||
        !(
          sideValue["style"] === undefined ||
          sideValue["style"] === null ||
          (typeof sideValue["style"] === "string" && BORDER_STYLE_VALUES.has(sideValue["style"]))
        ) ||
        !(
          sideValue["weight"] === undefined ||
          sideValue["weight"] === null ||
          (typeof sideValue["weight"] === "string" && BORDER_WEIGHT_VALUES.has(sideValue["weight"]))
        ) ||
        !(
          sideValue["color"] === undefined ||
          sideValue["color"] === null ||
          hasString(sideValue, "color")
        )
      ) {
        return false;
      }
    }
  }

  return true;
}

function isWorkbookConditionalFormatRule(value: unknown): boolean {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  switch (value["kind"]) {
    case "cellIs":
      return (
        typeof value["operator"] === "string" &&
        VALIDATION_COMPARISON_OPERATOR_VALUES.has(value["operator"]) &&
        Array.isArray(value["values"]) &&
        value["values"].every((entry) => isLiteralInput(entry)) &&
        (value["operator"] === "between" || value["operator"] === "notBetween"
          ? value["values"].length === 2
          : value["values"].length === 1)
      );
    case "textContains":
      return hasString(value, "text") && isOptionalBoolean(value["caseSensitive"]);
    case "formula":
      return hasString(value, "formula");
    case "blanks":
    case "notBlanks":
      return true;
    default:
      return false;
  }
}

function isWorkbookConditionalFormat(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "id") &&
    isCellRangeRef(value["range"]) &&
    isWorkbookConditionalFormatRule(value["rule"]) &&
    isCellStylePatch(value["style"]) &&
    isOptionalBoolean(value["stopIfTrue"]) &&
    isOptionalNumber(value["priority"])
  );
}

function isWorkbookSheetProtection(value: unknown): boolean {
  return (
    isRecord(value) && hasString(value, "sheetName") && isOptionalBoolean(value["hideFormulas"])
  );
}

function isWorkbookRangeProtection(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "id") &&
    isCellRangeRef(value["range"]) &&
    isOptionalBoolean(value["hideFormulas"])
  );
}

function isWorkbookCommentEntry(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "id") &&
    hasString(value, "body") &&
    isOptionalString(value["authorUserId"]) &&
    isOptionalString(value["authorDisplayName"]) &&
    isOptionalNumber(value["createdAtUnixMs"])
  );
}

function isWorkbookCommentThread(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "threadId") &&
    hasString(value, "sheetName") &&
    hasString(value, "address") &&
    Array.isArray(value["comments"]) &&
    value["comments"].length > 0 &&
    value["comments"].every((entry) => isWorkbookCommentEntry(entry)) &&
    isOptionalBoolean(value["resolved"]) &&
    isOptionalString(value["resolvedByUserId"]) &&
    isOptionalNumber(value["resolvedAtUnixMs"])
  );
}

function isWorkbookNote(value: unknown): boolean {
  return (
    isRecord(value) &&
    hasString(value, "sheetName") &&
    hasString(value, "address") &&
    hasString(value, "text")
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
    case "setSheetProtection":
      return isWorkbookSheetProtection(value["protection"]);
    case "clearSheetProtection":
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
    case "setDataValidation":
      return isWorkbookDataValidation(value["validation"]);
    case "clearDataValidation":
      return hasString(value, "sheetName") && isCellRangeRef(value["range"]);
    case "upsertConditionalFormat":
      return isWorkbookConditionalFormat(value["format"]);
    case "deleteConditionalFormat":
      return hasString(value, "id") && hasString(value, "sheetName");
    case "upsertRangeProtection":
      return isWorkbookRangeProtection(value["protection"]);
    case "deleteRangeProtection":
      return hasString(value, "id") && hasString(value, "sheetName");
    case "upsertCommentThread":
      return isWorkbookCommentThread(value["thread"]);
    case "deleteCommentThread":
    case "deleteNote":
      return hasString(value, "sheetName") && hasString(value, "address");
    case "upsertNote":
      return isWorkbookNote(value["note"]);
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
