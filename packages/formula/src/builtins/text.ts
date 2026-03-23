import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { createBlockedBuiltinMap, textPlaceholderBuiltinNames } from "./placeholder.js";

export type TextBuiltin = (...args: CellValue[]) => CellValue;

function error(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function stringResult(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function firstError(args: readonly CellValue[]): CellValue | undefined {
  return args.find((arg) => arg.tag === ValueTag.Error);
}

function coerceText(value: CellValue): string {
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

function coerceNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value;
    case ValueTag.Boolean:
      return value.value ? 1 : 0;
    case ValueTag.Empty:
      return 0;
    case ValueTag.String: {
      const trimmed = value.value.trim();
      if (trimmed === "") {
        return 0;
      }
      const parsed = Number(trimmed);
      return Number.isFinite(parsed) ? parsed : undefined;
    }
    case ValueTag.Error:
      return undefined;
  }
}

function coercePositiveStart(value: CellValue | undefined, defaultValue: number): number | CellValue {
  if (value === undefined) {
    return defaultValue;
  }
  const numeric = coerceNumber(value);
  if (numeric === undefined) {
    return error(ErrorCode.Value);
  }
  const truncated = Math.trunc(numeric);
  return truncated >= 1 ? truncated : error(ErrorCode.Value);
}

function coerceLength(value: CellValue | undefined, defaultValue: number): number | CellValue {
  if (value === undefined) {
    return defaultValue;
  }
  const numeric = coerceNumber(value);
  if (numeric === undefined) {
    return error(ErrorCode.Value);
  }
  const truncated = Math.trunc(numeric);
  return truncated >= 0 ? truncated : error(ErrorCode.Value);
}

function isErrorValue(value: number | CellValue): value is CellValue {
  return typeof value !== "number";
}

function excelTrim(input: string): string {
  let start = 0;
  let end = input.length;

  while (start < end && input.charCodeAt(start) === 32) {
    start += 1;
  }
  while (end > start && input.charCodeAt(end - 1) === 32) {
    end -= 1;
  }

  return input.slice(start, end).replace(/ {2,}/g, " ");
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function indexOfWithMode(text: string, delimiter: string, start: number, matchMode: number): number {
  if (matchMode === 1) {
    return text.toLowerCase().indexOf(delimiter.toLowerCase(), start);
  }
  return text.indexOf(delimiter, start);
}

function lastIndexOfWithMode(text: string, delimiter: string, start: number, matchMode: number): number {
  if (matchMode === 1) {
    return text.toLowerCase().lastIndexOf(delimiter.toLowerCase(), start);
  }
  return text.lastIndexOf(delimiter, start);
}

function hasSearchSyntax(pattern: string): boolean {
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]!;
    if (char === "~") {
      return true;
    }
    if (char === "*" || char === "?") {
      return true;
    }
  }
  return false;
}

function buildSearchRegex(pattern: string): RegExp {
  let source = "^";

  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]!;
    if (char === "~") {
      const next = pattern[index + 1];
      if (next === undefined) {
        source += escapeRegExp(char);
      } else {
        source += escapeRegExp(next);
        index += 1;
      }
      continue;
    }
    if (char === "*") {
      source += "[\\s\\S]*";
      continue;
    }
    if (char === "?") {
      source += "[\\s\\S]";
      continue;
    }
    source += escapeRegExp(char);
  }

  return new RegExp(source, "i");
}

function findPosition(needle: string, haystack: string, start: number, caseSensitive: boolean, wildcardAware: boolean): number | CellValue {
  const startIndex = start - 1;

  if (needle === "") {
    return start;
  }
  if (startIndex > haystack.length) {
    return error(ErrorCode.Value);
  }

  if (!wildcardAware || !hasSearchSyntax(needle)) {
    const normalizedHaystack = caseSensitive ? haystack : haystack.toLowerCase();
    const normalizedNeedle = caseSensitive ? needle : needle.toLowerCase();
    const found = normalizedHaystack.indexOf(normalizedNeedle, startIndex);
    return found === -1 ? error(ErrorCode.Value) : found + 1;
  }

  const regex = buildSearchRegex(needle);
  for (let index = startIndex; index <= haystack.length; index += 1) {
    if (regex.test(haystack.slice(index))) {
      return index + 1;
    }
  }
  return error(ErrorCode.Value);
}

const textPlaceholderBuiltins = createBlockedBuiltinMap(textPlaceholderBuiltinNames);

export const textBuiltins: Record<string, TextBuiltin> = {
  LEN: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    return numberResult(coerceText(value).length);
  },
  CONCAT: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    return stringResult(args.map(coerceText).join(""));
  },
  EXACT: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [leftValue, rightValue] = args;
    if (leftValue === undefined || rightValue === undefined) {
      return error(ErrorCode.Value);
    }
    return { tag: ValueTag.Boolean, value: coerceText(leftValue) === coerceText(rightValue) };
  },
  LEFT: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, countValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    const count = coerceLength(countValue, 1);
    if (isErrorValue(count)) {
      return count;
    }
    return stringResult(coerceText(textValue).slice(0, count));
  },
  RIGHT: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, countValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    const count = coerceLength(countValue, 1);
    if (isErrorValue(count)) {
      return count;
    }
    const text = coerceText(textValue);
    return stringResult(count === 0 ? "" : text.slice(-count));
  },
  MID: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, startValue, countValue] = args;
    if (textValue === undefined || startValue === undefined || countValue === undefined) {
      return error(ErrorCode.Value);
    }
    const start = coercePositiveStart(startValue, 1);
    if (isErrorValue(start)) {
      return start;
    }
    const count = coerceLength(countValue, 0);
    if (isErrorValue(count)) {
      return count;
    }
    const text = coerceText(textValue);
    return stringResult(text.slice(start - 1, start - 1 + count));
  },
  TRIM: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(excelTrim(coerceText(value)));
  },
  UPPER: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(coerceText(value).toUpperCase());
  },
  LOWER: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(coerceText(value).toLowerCase());
  },
  FIND: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [findTextValue, withinTextValue, startValue] = args;
    if (findTextValue === undefined || withinTextValue === undefined) {
      return error(ErrorCode.Value);
    }
    const start = coercePositiveStart(startValue, 1);
    if (isErrorValue(start)) {
      return start;
    }
    const found = findPosition(coerceText(findTextValue), coerceText(withinTextValue), start, true, false);
    return isErrorValue(found) ? found : numberResult(found);
  },
  SEARCH: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [findTextValue, withinTextValue, startValue] = args;
    if (findTextValue === undefined || withinTextValue === undefined) {
      return error(ErrorCode.Value);
    }
    const start = coercePositiveStart(startValue, 1);
    if (isErrorValue(start)) {
      return start;
    }
    const found = findPosition(coerceText(findTextValue), coerceText(withinTextValue), start, false, true);
    return isErrorValue(found) ? found : numberResult(found);
  },
  VALUE: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    const coerced = coerceNumber(value);
    return coerced === undefined ? error(ErrorCode.Value) : numberResult(coerced);
  },
  TEXTBEFORE: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, delimiterValue, instanceValue, matchModeValue, matchEndValue, ifNotFoundValue] = args;
    if (textValue === undefined || delimiterValue === undefined) {
      return error(ErrorCode.Value);
    }

    const text = coerceText(textValue);
    const delimiter = coerceText(delimiterValue);
    if (delimiter === "") {
      return error(ErrorCode.Value);
    }

    const instanceNumber = instanceValue === undefined ? 1 : coerceNumber(instanceValue);
    const matchMode = matchModeValue === undefined ? 0 : coerceNumber(matchModeValue);
    const matchEndNumber = matchEndValue === undefined ? 0 : coerceNumber(matchEndValue);
    if (
      instanceNumber === undefined
      || matchMode === undefined
      || matchEndNumber === undefined
      || !Number.isInteger(instanceNumber)
      || instanceNumber === 0
      || !Number.isInteger(matchMode)
      || (matchMode !== 0 && matchMode !== 1)
    ) {
      return error(ErrorCode.Value);
    }

    const matchEnd = matchEndNumber !== 0;
    if (instanceNumber > 0) {
      let searchFrom = 0;
      let found = -1;
      for (let count = 0; count < instanceNumber; count += 1) {
        found = indexOfWithMode(text, delimiter, searchFrom, matchMode);
        if (found === -1) {
          return ifNotFoundValue ?? error(ErrorCode.NA);
        }
        searchFrom = found + delimiter.length;
      }
      return stringResult(text.slice(0, found));
    }

    let searchFrom = text.length;
    let found = matchEnd ? text.length : -1;
    for (let count = 0; count < Math.abs(instanceNumber); count += 1) {
      found = lastIndexOfWithMode(text, delimiter, searchFrom, matchMode);
      if (found === -1) {
        return ifNotFoundValue ?? error(ErrorCode.NA);
      }
      searchFrom = found - 1;
    }
    return stringResult(text.slice(0, found));
  },
  ...textPlaceholderBuiltins
};

export function getTextBuiltin(name: string): TextBuiltin | undefined {
  return textBuiltins[name.toUpperCase()];
}
