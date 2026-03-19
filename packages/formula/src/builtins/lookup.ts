import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";

export interface RangeBuiltinArgument {
  kind: "range";
  values: CellValue[];
  refKind: "cells" | "rows" | "cols";
  rows: number;
  cols: number;
}

export type LookupBuiltinArgument = CellValue | RangeBuiltinArgument;
export type LookupBuiltin = (...args: LookupBuiltinArgument[]) => CellValue;

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function isError(value: CellValue): value is Extract<CellValue, { tag: ValueTag.Error }> {
  return value.tag === ValueTag.Error;
}

function isRangeArg(value: LookupBuiltinArgument): value is RangeBuiltinArgument {
  return typeof value === "object" && value !== null && "kind" in value && value.kind === "range";
}

function toNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value;
    case ValueTag.Boolean:
      return value.value ? 1 : 0;
    case ValueTag.Empty:
      return 0;
    default:
      return undefined;
  }
}

function toInteger(value: CellValue): number | undefined {
  const numeric = toNumber(value);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined;
  }
  return Math.trunc(numeric);
}

function toBoolean(value: CellValue): boolean | undefined {
  switch (value.tag) {
    case ValueTag.Boolean:
      return value.value;
    case ValueTag.Number:
      return value.value !== 0;
    case ValueTag.Empty:
      return false;
    default:
      return undefined;
  }
}

function toStringValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return String(value.value);
    case ValueTag.Boolean:
      return value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return value.value;
    case ValueTag.Error:
      return "";
  }
}

function compareScalars(left: CellValue, right: CellValue): number | undefined {
  if ((left.tag === ValueTag.String || left.tag === ValueTag.Empty) && (right.tag === ValueTag.String || right.tag === ValueTag.Empty)) {
    const normalizedLeft = toStringValue(left).toUpperCase();
    const normalizedRight = toStringValue(right).toUpperCase();
    if (normalizedLeft === normalizedRight) {
      return 0;
    }
    return normalizedLeft < normalizedRight ? -1 : 1;
  }

  const leftNum = toNumber(left);
  const rightNum = toNumber(right);
  if (leftNum === undefined || rightNum === undefined) {
    return undefined;
  }
  if (leftNum === rightNum) {
    return 0;
  }
  return leftNum < rightNum ? -1 : 1;
}

function requireCellVector(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg)) {
    return errorValue(ErrorCode.Value);
  }
  if (arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  if (arg.rows !== 1 && arg.cols !== 1) {
    return errorValue(ErrorCode.NA);
  }
  return arg;
}

function requireCellRange(arg: LookupBuiltinArgument): RangeBuiltinArgument | CellValue {
  if (!isRangeArg(arg) || arg.refKind !== "cells") {
    return errorValue(ErrorCode.Value);
  }
  return arg;
}

function getRangeValue(range: RangeBuiltinArgument, row: number, col: number): CellValue {
  const index = row * range.cols + col;
  return range.values[index] ?? { tag: ValueTag.Empty };
}

function exactMatch(lookupValue: CellValue, range: RangeBuiltinArgument): number {
  for (let index = 0; index < range.values.length; index += 1) {
    const comparison = compareScalars(range.values[index]!, lookupValue);
    if (comparison === 0) {
      return index + 1;
    }
  }
  return -1;
}

function approximateMatchAscending(lookupValue: CellValue, range: RangeBuiltinArgument): number {
  let best = -1;
  for (let index = 0; index < range.values.length; index += 1) {
    const comparison = compareScalars(range.values[index]!, lookupValue);
    if (comparison === undefined) {
      return -1;
    }
    if (comparison <= 0) {
      best = index + 1;
    } else {
      break;
    }
  }
  return best;
}

function approximateMatchDescending(lookupValue: CellValue, range: RangeBuiltinArgument): number {
  let best = -1;
  for (let index = 0; index < range.values.length; index += 1) {
    const comparison = compareScalars(range.values[index]!, lookupValue);
    if (comparison === undefined) {
      return -1;
    }
    if (comparison >= 0) {
      best = index + 1;
      continue;
    }
    break;
  }
  return best;
}

export const lookupBuiltins: Record<string, LookupBuiltin> = {
  MATCH: (lookupValue, lookupArray, matchTypeValue = { tag: ValueTag.Number, value: 1 }) => {
    if (isRangeArg(lookupValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(matchTypeValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(lookupValue)) {
      return lookupValue;
    }
    if (isError(matchTypeValue)) {
      return matchTypeValue;
    }

    const rangeOrError = requireCellVector(lookupArray);
    if (!isRangeArg(rangeOrError)) {
      return rangeOrError;
    }

    const matchType = toInteger(matchTypeValue);
    if (matchType === undefined || ![-1, 0, 1].includes(matchType)) {
      return errorValue(ErrorCode.Value);
    }

    const position =
      matchType === 0
        ? exactMatch(lookupValue, rangeOrError)
        : matchType === 1
          ? approximateMatchAscending(lookupValue, rangeOrError)
          : approximateMatchDescending(lookupValue, rangeOrError);

    return position === -1 ? errorValue(ErrorCode.NA) : { tag: ValueTag.Number, value: position };
  },
  INDEX: (array, rowNumValue, colNumValue = { tag: ValueTag.Number, value: 1 }) => {
    if (!isRangeArg(array) || array.refKind !== "cells") {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(rowNumValue) || isRangeArg(colNumValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(rowNumValue)) {
      return rowNumValue;
    }
    if (isError(colNumValue)) {
      return colNumValue;
    }

    const rawRowNum = toInteger(rowNumValue);
    const rawColNum = toInteger(colNumValue);
    if (rawRowNum === undefined || rawColNum === undefined) {
      return errorValue(ErrorCode.Value);
    }

    let rowNum = rawRowNum;
    let colNum = rawColNum;
    if (array.rows === 1 && rawColNum === 1) {
      rowNum = 1;
      colNum = rawRowNum;
    }

    if (rowNum < 1 || colNum < 1 || rowNum > array.rows || colNum > array.cols) {
      return errorValue(ErrorCode.Ref);
    }

    return getRangeValue(array, rowNum - 1, colNum - 1);
  },
  VLOOKUP: (lookupValue, tableArray, colIndexValue, rangeLookupValue = { tag: ValueTag.Boolean, value: true }) => {
    if (isRangeArg(lookupValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (!isRangeArg(tableArray) || tableArray.refKind !== "cells") {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(colIndexValue) || isRangeArg(rangeLookupValue)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(lookupValue)) {
      return lookupValue;
    }
    if (isError(colIndexValue)) {
      return colIndexValue;
    }
    if (isError(rangeLookupValue)) {
      return rangeLookupValue;
    }

    const colIndex = toInteger(colIndexValue);
    const rangeLookup = toBoolean(rangeLookupValue);
    if (colIndex === undefined || colIndex < 1 || colIndex > tableArray.cols || rangeLookup === undefined) {
      return errorValue(ErrorCode.Value);
    }

    let matchedRow = -1;
    for (let row = 0; row < tableArray.rows; row += 1) {
      const comparison = compareScalars(getRangeValue(tableArray, row, 0), lookupValue);
      if (comparison === undefined) {
        return errorValue(ErrorCode.Value);
      }
      if (comparison === 0) {
        matchedRow = row;
        break;
      }
      if (rangeLookup && comparison < 0) {
        matchedRow = row;
        continue;
      }
      if (rangeLookup && comparison > 0) {
        break;
      }
    }

    if (matchedRow === -1) {
      return errorValue(ErrorCode.NA);
    }
    return getRangeValue(tableArray, matchedRow, colIndex - 1);
  },
  XLOOKUP: (
    lookupValue,
    lookupArray,
    returnArray,
    ifNotFound = { tag: ValueTag.Error, code: ErrorCode.NA },
    matchMode = { tag: ValueTag.Number, value: 0 },
    searchMode = { tag: ValueTag.Number, value: 1 }
  ) => {
    if (isRangeArg(lookupValue) || isRangeArg(ifNotFound) || isRangeArg(matchMode) || isRangeArg(searchMode)) {
      return errorValue(ErrorCode.Value);
    }
    const lookupRange = requireCellVector(lookupArray);
    const returnRange = requireCellVector(returnArray);
    if (!isRangeArg(lookupRange)) {
      return lookupRange;
    }
    if (!isRangeArg(returnRange)) {
      return returnRange;
    }
    if (lookupRange.values.length !== returnRange.values.length) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(lookupValue)) {
      return lookupValue;
    }
    if (isError(matchMode)) {
      return matchMode;
    }
    if (isError(searchMode)) {
      return searchMode;
    }

    const matchModeNumber = toInteger(matchMode);
    const searchModeNumber = toInteger(searchMode);
    if ((matchModeNumber ?? 0) !== 0 || (searchModeNumber !== 1 && searchModeNumber !== -1)) {
      return errorValue(ErrorCode.Value);
    }

    if (searchModeNumber === -1) {
      for (let index = lookupRange.values.length - 1; index >= 0; index -= 1) {
        if (compareScalars(lookupRange.values[index]!, lookupValue) === 0) {
          return returnRange.values[index] ?? errorValue(ErrorCode.NA);
        }
      }
      return ifNotFound;
    }

    for (let index = 0; index < lookupRange.values.length; index += 1) {
      if (compareScalars(lookupRange.values[index]!, lookupValue) === 0) {
        return returnRange.values[index] ?? errorValue(ErrorCode.NA);
      }
    }
    return ifNotFound;
  },
  COUNTIF: (rangeArg, criteriaArg) => {
    const range = requireCellRange(rangeArg);
    if (!isRangeArg(range)) {
      return range;
    }
    if (isRangeArg(criteriaArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(criteriaArg)) {
      return criteriaArg;
    }
    let count = 0;
    for (const value of range.values) {
      if (matchesCriteria(value, criteriaArg)) {
        count += 1;
      }
    }
    return { tag: ValueTag.Number, value: count };
  },
  AVERAGEIF: (rangeArg, criteriaArg, averageRangeArg = rangeArg) => {
    const range = requireCellRange(rangeArg);
    const averageRange = requireCellRange(averageRangeArg);
    if (!isRangeArg(range)) {
      return range;
    }
    if (!isRangeArg(averageRange)) {
      return averageRange;
    }
    if (range.values.length !== averageRange.values.length) {
      return errorValue(ErrorCode.Value);
    }
    if (isRangeArg(criteriaArg)) {
      return errorValue(ErrorCode.Value);
    }
    if (isError(criteriaArg)) {
      return criteriaArg;
    }

    let count = 0;
    let sum = 0;
    for (let index = 0; index < range.values.length; index += 1) {
      if (!matchesCriteria(range.values[index]!, criteriaArg)) {
        continue;
      }
      const numeric = toNumber(averageRange.values[index]!);
      if (numeric === undefined) {
        continue;
      }
      count += 1;
      sum += numeric;
    }

    if (count === 0) {
      return errorValue(ErrorCode.Div0);
    }
    return { tag: ValueTag.Number, value: sum / count };
  }
};

type CriteriaOperator = "=" | "<>" | ">" | ">=" | "<" | "<=";

function matchesCriteria(value: CellValue, criteria: CellValue): boolean {
  if (isError(value)) {
    return false;
  }
  const { operator, operand } = parseCriteria(criteria);
  const comparison = compareScalars(value, operand);
  if (comparison === undefined) {
    return false;
  }
  switch (operator) {
    case "=":
      return comparison === 0;
    case "<>":
      return comparison !== 0;
    case ">":
      return comparison > 0;
    case ">=":
      return comparison >= 0;
    case "<":
      return comparison < 0;
    case "<=":
      return comparison <= 0;
  }
}

function parseCriteria(criteria: CellValue): { operator: CriteriaOperator; operand: CellValue } {
  if (criteria.tag !== ValueTag.String) {
    return { operator: "=", operand: criteria };
  }

  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(criteria.value);
  if (!match) {
    return { operator: "=", operand: criteria };
  }

  return {
    operator: match[1] as CriteriaOperator,
    operand: parseCriteriaOperand(match[2] ?? "")
  };
}

function parseCriteriaOperand(raw: string): CellValue {
  const trimmed = raw.trim();
  if (trimmed === "") {
    return { tag: ValueTag.String, value: "", stringId: 0 };
  }
  const upper = trimmed.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") {
    return { tag: ValueTag.Boolean, value: upper === "TRUE" };
  }
  const numeric = Number(trimmed);
  if (Number.isFinite(numeric)) {
    return { tag: ValueTag.Number, value: numeric };
  }
  return { tag: ValueTag.String, value: trimmed, stringId: 0 };
}

export function getLookupBuiltin(name: string): LookupBuiltin | undefined {
  return lookupBuiltins[name.toUpperCase()];
}
