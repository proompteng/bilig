import { ErrorCode, ValueTag, type CellValue, type LiteralInput } from "@bilig/protocol";
import { StringPool } from "./string-pool.js";

export function normalizePivotLookupText(value: string): string {
  return value.trim().toUpperCase();
}

export function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

export function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

export function literalToValue(input: LiteralInput, stringPool: StringPool): CellValue {
  if (input === null) return emptyValue();
  if (typeof input === "number") return { tag: ValueTag.Number, value: input };
  if (typeof input === "boolean") return { tag: ValueTag.Boolean, value: input };
  return { tag: ValueTag.String, value: input, stringId: stringPool.intern(input) };
}

export function areCellValuesEqual(left: CellValue, right: CellValue): boolean {
  if (left.tag !== right.tag) {
    return false;
  }
  switch (left.tag) {
    case ValueTag.Empty:
      return true;
    case ValueTag.Number:
      return right.tag === ValueTag.Number && Object.is(left.value, right.value);
    case ValueTag.Boolean:
      return right.tag === ValueTag.Boolean && left.value === right.value;
    case ValueTag.String:
      return right.tag === ValueTag.String && left.value === right.value;
    case ValueTag.Error:
      return right.tag === ValueTag.Error && left.code === right.code;
  }
}

export function cellValueDisplayText(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return Object.is(value.value, -0) ? "-0" : String(value.value);
    case ValueTag.Boolean:
      return value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return value.value;
    case ValueTag.Error:
      return `#${ErrorCode[value.code] ?? "ERROR"}!`;
  }
}

export function pivotItemMatches(cell: CellValue, item: CellValue): boolean {
  if (areCellValuesEqual(cell, item)) {
    return true;
  }
  if (item.tag === ValueTag.String) {
    return (
      normalizePivotLookupText(cellValueDisplayText(cell)) === normalizePivotLookupText(item.value)
    );
  }
  return false;
}
