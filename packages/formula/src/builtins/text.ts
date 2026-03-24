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

const utf8Encoder = new TextEncoder();
const utf8Decoder = new TextDecoder();

function utf8Bytes(value: string): Uint8Array {
  return utf8Encoder.encode(value);
}

function utf8Text(bytes: Uint8Array): string {
  return utf8Decoder.decode(bytes);
}

function findSubBytes(haystack: Uint8Array, needle: Uint8Array, start: number): number {
  if (needle.length === 0) {
    return Math.max(0, Math.min(start, haystack.length));
  }

  for (let index = start; index + needle.length <= haystack.length; index += 1) {
    let match = true;
    for (let offset = 0; offset < needle.length; offset += 1) {
      if (haystack[index + offset] !== needle[offset]) {
        match = false;
        break;
      }
    }
    if (match) {
      return index;
    }
  }
  return -1;
}

function leftBytes(text: string, byteCount: number): string {
  const bytes = utf8Bytes(text);
  const normalizedCount = Math.max(0, Math.min(byteCount, bytes.length));
  return utf8Text(bytes.slice(0, normalizedCount));
}

function rightBytes(text: string, byteCount: number): string {
  const bytes = utf8Bytes(text);
  const normalizedCount = Math.max(0, Math.min(byteCount, bytes.length));
  return utf8Text(bytes.slice(bytes.length - normalizedCount));
}

function midBytes(text: string, start: number, byteCount: number): string {
  const bytes = utf8Bytes(text);
  if (byteCount <= 0) {
    return "";
  }

  const zeroBasedStart = Math.max(0, start - 1);
  const zeroBasedEnd = Math.min(bytes.length, zeroBasedStart + byteCount);
  if (zeroBasedStart >= bytes.length) {
    return "";
  }
  return utf8Text(bytes.slice(zeroBasedStart, zeroBasedEnd));
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

function coerceBoolean(value: CellValue, fallback: boolean): boolean | CellValue {
  if (value.tag === ValueTag.Boolean) {
    return value.value;
  }
  if (value.tag === ValueTag.Empty) {
    return fallback;
  }
  const numeric = coerceNumber(value);
  return numeric === undefined ? error(ErrorCode.Value) : numeric !== 0;
}

function coercePositiveStart(
  value: CellValue | undefined,
  defaultValue: number,
): number | CellValue {
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

function coerceNonNegativeInt(
  value: CellValue | undefined,
  defaultValue: number,
): number | CellValue {
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

function replaceSingle(text: string, start: number, count: number, replacement: string): string {
  const index = start - 1;
  if (index >= text.length) {
    return text;
  }
  return text.slice(0, index) + replacement + text.slice(index + count);
}

function substituteText(text: string, oldText: string, newText: string, instance?: number): string {
  if (oldText === "") {
    return text;
  }
  if (instance === undefined) {
    if (!text.includes(oldText)) {
      return text;
    }
    return text.split(oldText).join(newText);
  }

  let occurrence = 0;
  let searchIndex = 0;
  while (searchIndex <= text.length) {
    const foundAt = text.indexOf(oldText, searchIndex);
    if (foundAt === -1) {
      return text;
    }
    occurrence += 1;
    if (occurrence === instance) {
      return text.slice(0, foundAt) + newText + text.slice(foundAt + oldText.length);
    }
    searchIndex = foundAt + oldText.length;
  }
  return text;
}

function createReplaceBuiltin(): TextBuiltin {
  return (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, startValue, countValue, replacementValue] = args;
    if (
      textValue === undefined ||
      startValue === undefined ||
      countValue === undefined ||
      replacementValue === undefined
    ) {
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
    const replacement = coerceText(replacementValue);
    const replaced = replaceSingle(coerceText(textValue), start, count, replacement);
    return stringResult(replaced);
  };
}

function createSubstituteBuiltin(): TextBuiltin {
  return (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, oldValue, newValue, instanceValue] = args;
    if (textValue === undefined || oldValue === undefined || newValue === undefined) {
      return error(ErrorCode.Value);
    }
    const text = coerceText(textValue);
    const oldText = coerceText(oldValue);
    if (oldText === "") {
      return error(ErrorCode.Value);
    }
    const newText = coerceText(newValue);
    if (instanceValue === undefined) {
      return stringResult(substituteText(text, oldText, newText));
    }

    const instance = coercePositiveStart(instanceValue, 1);
    if (isErrorValue(instance)) {
      return instance;
    }
    return stringResult(substituteText(text, oldText, newText, instance));
  };
}

function createReptBuiltin(): TextBuiltin {
  return (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, countValue] = args;
    if (textValue === undefined || countValue === undefined) {
      return error(ErrorCode.Value);
    }
    const count = coerceNonNegativeInt(countValue, 0);
    if (isErrorValue(count)) {
      return count;
    }
    const text = coerceText(textValue);
    let repeated = "";
    for (let index = 0; index < count; index += 1) {
      repeated += text;
    }
    return stringResult(repeated);
  };
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

function stripControlCharacters(input: string): string {
  let output = "";
  for (let index = 0; index < input.length; index += 1) {
    const char = input.charCodeAt(index);
    if ((char >= 0 && char <= 31) || char === 127) {
      continue;
    }
    output += input[index] ?? "";
  }
  return output;
}

function toTitleCase(input: string): string {
  let result = "";
  let capitalizeNext = true;

  for (let index = 0; index < input.length; index += 1) {
    const char = input[index] ?? "";
    const code = char.charCodeAt(0);
    const isAlpha = (code >= 65 && code <= 90) || (code >= 97 && code <= 122);

    if (!isAlpha) {
      capitalizeNext = true;
      result += char;
      continue;
    }

    result += capitalizeNext ? char.toUpperCase() : char.toLowerCase();
    capitalizeNext = false;
  }

  return result;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function indexOfWithMode(
  text: string,
  delimiter: string,
  start: number,
  matchMode: number,
): number {
  if (matchMode === 1) {
    return text.toLowerCase().indexOf(delimiter.toLowerCase(), start);
  }
  return text.indexOf(delimiter, start);
}

function lastIndexOfWithMode(
  text: string,
  delimiter: string,
  start: number,
  matchMode: number,
): number {
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

function findPosition(
  needle: string,
  haystack: string,
  start: number,
  caseSensitive: boolean,
  wildcardAware: boolean,
): number | CellValue {
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

function charCodeFromArgument(value: CellValue | undefined): number | CellValue {
  if (value === undefined) {
    return error(ErrorCode.Value);
  }
  const code = coerceNumber(value);
  if (code === undefined) {
    return error(ErrorCode.Value);
  }
  const integerCode = Math.trunc(code);
  if (!Number.isFinite(integerCode) || integerCode < 1 || integerCode > 255) {
    return error(ErrorCode.Value);
  }
  return integerCode;
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
  CHAR: (...args) => {
    const [codeValue] = args;
    const codePoint = charCodeFromArgument(codeValue);
    if (isErrorValue(codePoint)) {
      return codePoint;
    }
    return stringResult(String.fromCodePoint(codePoint));
  },
  CODE: (...args) => {
    const [textValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    const text = coerceText(textValue);
    if (text.length === 0) {
      return error(ErrorCode.Value);
    }
    const codePoint = text.codePointAt(0);
    return codePoint === undefined ? error(ErrorCode.Value) : numberResult(codePoint);
  },
  UNICODE: (...args) => {
    const [textValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    const text = coerceText(textValue);
    if (text.length === 0) {
      return error(ErrorCode.Value);
    }
    const codePoint = text.codePointAt(0);
    return codePoint === undefined ? error(ErrorCode.Value) : numberResult(codePoint);
  },
  UNICHAR: (...args) => {
    const [codeValue] = args;
    if (codeValue === undefined) {
      return error(ErrorCode.Value);
    }
    const code = coerceNumber(codeValue);
    if (code === undefined) {
      return error(ErrorCode.Value);
    }
    const integerCode = Math.trunc(code);
    if (!Number.isFinite(integerCode) || integerCode < 0 || integerCode > 0x10ffff) {
      return error(ErrorCode.Value);
    }
    return stringResult(String.fromCodePoint(integerCode));
  },
  CLEAN: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(stripControlCharacters(coerceText(textValue)));
  },
  CONCATENATE: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    if (args.length === 0) {
      return error(ErrorCode.Value);
    }
    return stringResult(args.map(coerceText).join(""));
  },
  CONCAT: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    return stringResult(args.map(coerceText).join(""));
  },
  PROPER: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(toTitleCase(coerceText(textValue)));
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
    const found = findPosition(
      coerceText(findTextValue),
      coerceText(withinTextValue),
      start,
      true,
      false,
    );
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
    const found = findPosition(
      coerceText(findTextValue),
      coerceText(withinTextValue),
      start,
      false,
      true,
    );
    return isErrorValue(found) ? found : numberResult(found);
  },
  ENCODEURL: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(encodeURI(coerceText(value)));
  },
  FINDB: (...args) => {
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
    const findBytes = utf8Bytes(coerceText(findTextValue));
    const withinBytes = utf8Bytes(coerceText(withinTextValue));
    if (start > withinBytes.length + 1) {
      return error(ErrorCode.Value);
    }
    const found = findSubBytes(withinBytes, findBytes, start - 1);
    return found === -1 ? error(ErrorCode.Value) : numberResult(found + 1);
  },
  LEFTB: (...args) => {
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
    return stringResult(leftBytes(coerceText(textValue), count));
  },
  MIDB: (...args) => {
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
    return stringResult(midBytes(coerceText(textValue), start, count));
  },
  RIGHTB: (...args) => {
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
    return stringResult(rightBytes(coerceText(textValue), count));
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
    const [
      textValue,
      delimiterValue,
      instanceValue,
      matchModeValue,
      matchEndValue,
      ifNotFoundValue,
    ] = args;
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
      instanceNumber === undefined ||
      matchMode === undefined ||
      matchEndNumber === undefined ||
      !Number.isInteger(instanceNumber) ||
      instanceNumber === 0 ||
      !Number.isInteger(matchMode) ||
      (matchMode !== 0 && matchMode !== 1)
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
  TEXTAFTER: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [
      textValue,
      delimiterValue,
      instanceValue,
      matchModeValue,
      matchEndValue,
      ifNotFoundValue,
    ] = args;
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
      instanceNumber === undefined ||
      matchMode === undefined ||
      matchEndNumber === undefined ||
      !Number.isInteger(instanceNumber) ||
      instanceNumber === 0 ||
      !Number.isInteger(matchMode) ||
      (matchMode !== 0 && matchMode !== 1)
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
      return stringResult(text.slice(found + delimiter.length));
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
    const start = found + delimiter.length;
    return stringResult(text.slice(start));
  },
  TEXTJOIN: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [delimiterValue, ignoreEmptyValue, ...values] = args;
    if (delimiterValue === undefined || ignoreEmptyValue === undefined || values.length === 0) {
      return error(ErrorCode.Value);
    }

    const delimiter = coerceText(delimiterValue);
    const ignoreEmpty = coerceBoolean(ignoreEmptyValue, false);
    if (ignoreEmpty === undefined) {
      return error(ErrorCode.Value);
    }

    const valuesJoined: string[] = [];
    for (const value of values) {
      if (value === undefined) {
        continue;
      }
      if (value.tag === ValueTag.Empty) {
        if (!ignoreEmpty) {
          valuesJoined.push("");
        }
        continue;
      }
      if (value.tag === ValueTag.String && value.value === "" && !ignoreEmpty) {
        valuesJoined.push("");
        continue;
      }
      if (value.tag === ValueTag.String && value.value === "" && ignoreEmpty) {
        continue;
      }
      valuesJoined.push(coerceText(value));
    }

    return stringResult(valuesJoined.join(delimiter));
  },
  REPLACE: createReplaceBuiltin(),
  SUBSTITUTE: createSubstituteBuiltin(),
  REPT: createReptBuiltin(),
  ...textPlaceholderBuiltins,
};

export function getTextBuiltin(name: string): TextBuiltin | undefined {
  return textBuiltins[name.toUpperCase()];
}
