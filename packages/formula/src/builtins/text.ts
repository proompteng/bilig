import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { excelSerialToDateParts } from "./datetime.js";
import { createBlockedBuiltinMap, textPlaceholderBuiltinNames } from "./placeholder.js";
import type { EvaluationResult } from "../runtime-values.js";

export type TextBuiltin = (...args: CellValue[]) => EvaluationResult;

function error(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function stringResult(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function booleanResult(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value };
}

function firstError(args: readonly (CellValue | undefined)[]): CellValue | undefined {
  return args.find((arg) => arg?.tag === ValueTag.Error);
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

function replaceBytes(text: string, start: number, byteCount: number, replacement: string): string {
  const bytes = utf8Bytes(text);
  const replacementBytes = utf8Bytes(replacement);
  const zeroBasedStart = Math.max(0, start - 1);
  if (zeroBasedStart >= bytes.length) {
    return text;
  }
  const zeroBasedEnd = Math.min(bytes.length, zeroBasedStart + Math.max(0, byteCount));
  return utf8Text(
    new Uint8Array([
      ...bytes.slice(0, zeroBasedStart),
      ...replacementBytes,
      ...bytes.slice(zeroBasedEnd),
    ]),
  );
}

function bytePositionToCharPosition(text: string, startByte: number): number {
  if (startByte <= 1) {
    return 1;
  }
  return utf8Text(utf8Bytes(text).slice(0, startByte - 1)).length + 1;
}

function charPositionToBytePosition(text: string, charPosition: number): number {
  return utf8Bytes(text.slice(0, Math.max(0, charPosition - 1))).length + 1;
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

function coerceInteger(value: CellValue | undefined, defaultValue: number): number | CellValue {
  if (value === undefined) {
    return defaultValue;
  }
  const numeric = coerceNumber(value);
  if (numeric === undefined || !Number.isInteger(numeric)) {
    return error(ErrorCode.Value);
  }
  return numeric;
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

function regexFlags(caseSensitivity: number, global = false): string {
  return `${global ? "g" : ""}${caseSensitivity === 1 ? "i" : ""}`;
}

function compileRegex(
  pattern: string,
  caseSensitivity: number,
  global = false,
): RegExp | CellValue {
  try {
    return new RegExp(pattern, regexFlags(caseSensitivity, global));
  } catch {
    return error(ErrorCode.Value);
  }
}

function isRegexError(value: RegExp | CellValue): value is CellValue {
  return !(value instanceof RegExp);
}

function applyReplacementTemplate(
  template: string,
  match: string,
  captures: readonly (string | undefined)[],
): string {
  return template.replace(/\$(\$|&|[0-9]{1,2})/g, (_whole, token: string) => {
    if (token === "$") {
      return "$";
    }
    if (token === "&") {
      return match;
    }
    const index = Number(token);
    if (!Number.isInteger(index) || index <= 0) {
      return "";
    }
    return captures[index - 1] ?? "";
  });
}

function valueToTextResult(value: CellValue, format: number): CellValue {
  if (format !== 0 && format !== 1) {
    return error(ErrorCode.Value);
  }
  switch (value.tag) {
    case ValueTag.Empty:
      return stringResult("");
    case ValueTag.Number:
      return stringResult(String(value.value));
    case ValueTag.Boolean:
      return stringResult(value.value ? "TRUE" : "FALSE");
    case ValueTag.String:
      return stringResult(format === 1 ? JSON.stringify(value.value) : value.value);
    case ValueTag.Error: {
      const label =
        value.code === ErrorCode.Div0
          ? "#DIV/0!"
          : value.code === ErrorCode.Ref
            ? "#REF!"
            : value.code === ErrorCode.Value
              ? "#VALUE!"
              : value.code === ErrorCode.Name
                ? "#NAME?"
                : value.code === ErrorCode.NA
                  ? "#N/A"
                  : value.code === ErrorCode.Cycle
                    ? "#CYCLE!"
                    : value.code === ErrorCode.Spill
                      ? "#SPILL!"
                      : value.code === ErrorCode.Blocked
                        ? "#BLOCKED!"
                        : "#ERROR!";
      return stringResult(label);
    }
  }
}

const shortMonthNames = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
] as const;
const fullMonthNames = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
] as const;
const shortWeekdayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"] as const;
const fullWeekdayNames = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
] as const;

function splitFormatSections(format: string): string[] {
  const sections: string[] = [];
  let current = "";
  let inQuotes = false;
  let bracketDepth = 0;
  let escaped = false;
  for (let index = 0; index < format.length; index += 1) {
    const char = format[index]!;
    if (escaped) {
      current += char;
      escaped = false;
      continue;
    }
    if (char === "\\") {
      current += char;
      escaped = true;
      continue;
    }
    if (char === '"') {
      current += char;
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && char === "[") {
      bracketDepth += 1;
      current += char;
      continue;
    }
    if (!inQuotes && char === "]" && bracketDepth > 0) {
      bracketDepth -= 1;
      current += char;
      continue;
    }
    if (!inQuotes && bracketDepth === 0 && char === ";") {
      sections.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  sections.push(current);
  return sections;
}

function stripFormatDecorations(section: string): string {
  let output = "";
  let inQuotes = false;
  for (let index = 0; index < section.length; index += 1) {
    const char = section[index]!;
    if (inQuotes) {
      if (char === '"') {
        inQuotes = false;
      } else {
        output += char;
      }
      continue;
    }
    if (char === '"') {
      inQuotes = true;
      continue;
    }
    if (char === "\\") {
      output += section[index + 1] ?? "";
      index += 1;
      continue;
    }
    if (char === "_") {
      output += " ";
      index += 1;
      continue;
    }
    if (char === "*") {
      index += 1;
      continue;
    }
    if (char === "[") {
      const end = section.indexOf("]", index + 1);
      if (end === -1) {
        continue;
      }
      index = end;
      continue;
    }
    output += char;
  }
  return output;
}

function formatThousandsText(integerPart: string): string {
  return integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function zeroPadText(value: number, width: number): string {
  return String(Math.trunc(Math.abs(value))).padStart(width, "0");
}

function roundToDigits(value: number, digits: number): number {
  if (!Number.isFinite(value)) {
    return Number.NaN;
  }
  const factor = 10 ** Math.max(0, digits);
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function excelSecondOfDay(serial: number): number | undefined {
  if (!Number.isFinite(serial)) {
    return undefined;
  }
  const whole = Math.floor(serial);
  let fraction = serial - whole;
  if (fraction < 0) {
    fraction += 1;
  }
  let seconds = Math.floor(fraction * 86_400 + 1e-9);
  if (seconds >= 86_400) {
    seconds = 0;
  }
  return seconds;
}

function excelWeekdayIndex(serial: number): number | undefined {
  if (!Number.isFinite(serial)) {
    return undefined;
  }
  const whole = Math.floor(serial);
  if (whole < 0) {
    return undefined;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  return ((adjustedWhole % 7) + 7) % 7;
}

function isDateTimeFormat(section: string): boolean {
  const cleaned = stripFormatDecorations(section).toUpperCase();
  return (
    cleaned.includes("AM/PM") ||
    cleaned.includes("A/P") ||
    /[YDSH]/.test(cleaned) ||
    /(^|[^0#?])M+([^0#?]|$)/.test(cleaned)
  );
}

function isTextFormat(section: string): boolean {
  return stripFormatDecorations(section).includes("@");
}

function chooseFormatSection(
  value: CellValue,
  formatText: string,
): { section: string; numeric?: number; autoNegative: boolean } | CellValue {
  const sections = splitFormatSections(formatText);
  if (value.tag === ValueTag.String) {
    return { section: sections[3] ?? sections[0] ?? "", autoNegative: false };
  }
  const numeric = coerceNumber(value);
  if (numeric === undefined) {
    return error(ErrorCode.Value);
  }
  if (numeric < 0) {
    if (sections[1] !== undefined) {
      return { section: sections[1], numeric: -numeric, autoNegative: false };
    }
    return { section: sections[0] ?? "", numeric: -numeric, autoNegative: true };
  }
  if (numeric === 0 && sections[2] !== undefined) {
    return { section: sections[2], numeric, autoNegative: false };
  }
  return { section: sections[0] ?? "", numeric, autoNegative: false };
}

function formatTextSectionValue(value: string, section: string): string {
  const cleaned = stripFormatDecorations(section);
  return cleaned.includes("@") ? cleaned.replace(/@/g, value) : cleaned;
}

interface DateTimeToken {
  kind: "literal" | "year" | "month" | "minute" | "day" | "hour" | "second" | "ampm";
  text: string;
}

function tokenizeDateTimeFormat(section: string): DateTimeToken[] {
  const cleaned = stripFormatDecorations(section);
  const tokens: DateTimeToken[] = [];
  let index = 0;
  while (index < cleaned.length) {
    const remainder = cleaned.slice(index);
    const upperRemainder = remainder.toUpperCase();
    if (upperRemainder.startsWith("AM/PM")) {
      tokens.push({ kind: "ampm", text: cleaned.slice(index, index + 5) });
      index += 5;
      continue;
    }
    if (upperRemainder.startsWith("A/P")) {
      tokens.push({ kind: "ampm", text: cleaned.slice(index, index + 3) });
      index += 3;
      continue;
    }
    const char = cleaned[index]!;
    const lower = char.toLowerCase();
    if ("ymdhms".includes(lower)) {
      let end = index + 1;
      while (end < cleaned.length && cleaned[end]!.toLowerCase() === lower) {
        end += 1;
      }
      const tokenText = cleaned.slice(index, end);
      const baseKind =
        lower === "y"
          ? "year"
          : lower === "d"
            ? "day"
            : lower === "h"
              ? "hour"
              : lower === "s"
                ? "second"
                : "month";
      tokens.push({ kind: baseKind, text: tokenText });
      index = end;
      continue;
    }
    tokens.push({ kind: "literal", text: char });
    index += 1;
  }
  return tokens.map((token, tokenIndex, allTokens) => {
    if (token.kind !== "month") {
      return token;
    }
    const previous = allTokens.slice(0, tokenIndex).findLast((entry) => entry.kind !== "literal");
    const next = allTokens.slice(tokenIndex + 1).find((entry) => entry.kind !== "literal");
    if (previous?.kind === "hour" || previous?.kind === "minute" || next?.kind === "second") {
      return { kind: "minute", text: token.text };
    }
    return token;
  });
}

function formatAmPmToken(token: string, hour: number): string {
  const isPm = hour >= 12;
  const upper = token.toUpperCase();
  if (upper === "A/P") {
    const letter = isPm ? "P" : "A";
    return token === token.toLowerCase() ? letter.toLowerCase() : letter;
  }
  if (token === token.toLowerCase()) {
    return isPm ? "pm" : "am";
  }
  return isPm ? "PM" : "AM";
}

function formatDateTimeSectionValue(serial: number, section: string): string | undefined {
  const dateParts = excelSerialToDateParts(serial);
  const weekdayIndex = excelWeekdayIndex(serial);
  const secondOfDay = excelSecondOfDay(serial);
  if (!dateParts || weekdayIndex === undefined || secondOfDay === undefined) {
    return undefined;
  }
  const hour24 = Math.floor(secondOfDay / 3600);
  const minute = Math.floor((secondOfDay % 3600) / 60);
  const second = secondOfDay % 60;
  const tokens = tokenizeDateTimeFormat(section);
  const hasAmPm = tokens.some((token) => token.kind === "ampm");
  return tokens
    .map((token) => {
      switch (token.kind) {
        case "literal":
          return token.text;
        case "year":
          return token.text.length === 2
            ? zeroPadText(dateParts.year % 100, 2)
            : String(dateParts.year).padStart(Math.max(4, token.text.length), "0");
        case "month":
          return token.text.length === 1
            ? String(dateParts.month)
            : token.text.length === 2
              ? zeroPadText(dateParts.month, 2)
              : token.text.length === 3
                ? shortMonthNames[dateParts.month - 1]!
                : fullMonthNames[dateParts.month - 1]!;
        case "minute":
          return token.text.length >= 2 ? zeroPadText(minute, 2) : String(minute);
        case "day":
          return token.text.length === 1
            ? String(dateParts.day)
            : token.text.length === 2
              ? zeroPadText(dateParts.day, 2)
              : token.text.length === 3
                ? shortWeekdayNames[weekdayIndex]!
                : fullWeekdayNames[weekdayIndex]!;
        case "hour": {
          const normalizedHour = hasAmPm ? ((hour24 + 11) % 12) + 1 : hour24;
          return token.text.length >= 2 ? zeroPadText(normalizedHour, 2) : String(normalizedHour);
        }
        case "second":
          return token.text.length >= 2 ? zeroPadText(second, 2) : String(second);
        case "ampm":
          return formatAmPmToken(token.text, hour24);
      }
    })
    .join("");
}

function trimOptionalFractionDigits(fraction: string, minDigits: number): string {
  let trimmed = fraction;
  while (trimmed.length > minDigits && trimmed.endsWith("0")) {
    trimmed = trimmed.slice(0, -1);
  }
  return trimmed;
}

function formatScientificSection(value: number, core: string): string {
  const exponentIndex = core.search(/[Ee][+-]/);
  const mantissaPattern = core.slice(0, exponentIndex);
  const exponentPattern = core.slice(exponentIndex + 2);
  const mantissaParts = mantissaPattern.split(".");
  const fractionPattern = mantissaParts[1] ?? "";
  const maxFractionDigits = (fractionPattern.match(/[0#?]/g) ?? []).length;
  const minFractionDigits = (fractionPattern.match(/0/g) ?? []).length;
  const [mantissaRaw = "0", exponentRaw] = value.toExponential(maxFractionDigits).split("e");
  let [integerPart = "0", fractionPart = ""] = mantissaRaw.split(".");
  fractionPart = trimOptionalFractionDigits(fractionPart, minFractionDigits);
  const exponentValue = Number(exponentRaw ?? 0);
  const exponentText = String(Math.abs(exponentValue)).padStart(exponentPattern.length, "0");
  return `${integerPart}${fractionPart === "" ? "" : `.${fractionPart}`}E${
    exponentValue < 0 ? "-" : "+"
  }${exponentText}`;
}

function formatNumericSectionValue(value: number, section: string, autoNegative: boolean): string {
  const cleaned = stripFormatDecorations(section);
  if (!/[0#?]/.test(cleaned)) {
    return autoNegative && !cleaned.startsWith("-") ? `-${cleaned}` : cleaned;
  }
  const firstPlaceholder = cleaned.search(/[0#?]/);
  let lastPlaceholder = -1;
  for (let index = cleaned.length - 1; index >= 0; index -= 1) {
    if (/[0#?]/.test(cleaned[index]!)) {
      lastPlaceholder = index;
      break;
    }
  }
  const prefix = cleaned.slice(0, firstPlaceholder);
  const core = cleaned.slice(firstPlaceholder, lastPlaceholder + 1);
  const suffix = cleaned.slice(lastPlaceholder + 1);
  const percentCount = (cleaned.match(/%/g) ?? []).length;
  const scaledValue = Math.abs(value) * 100 ** percentCount;
  let numericText = "";
  if (/[Ee][+-]/.test(core)) {
    numericText = formatScientificSection(scaledValue, core);
  } else {
    const decimalIndex = core.indexOf(".");
    const integerPattern = (decimalIndex === -1 ? core : core.slice(0, decimalIndex)).replaceAll(
      ",",
      "",
    );
    const fractionPattern = decimalIndex === -1 ? "" : core.slice(decimalIndex + 1);
    const maxFractionDigits = (fractionPattern.match(/[0#?]/g) ?? []).length;
    const minFractionDigits = (fractionPattern.match(/0/g) ?? []).length;
    const minIntegerDigits = (integerPattern.match(/0/g) ?? []).length;
    const roundedValue = roundToDigits(scaledValue, maxFractionDigits);
    const fixed = roundedValue.toFixed(maxFractionDigits);
    let [integerPart = "0", fractionPart = ""] = fixed.split(".");
    if (integerPart.length < minIntegerDigits) {
      integerPart = integerPart.padStart(minIntegerDigits, "0");
    }
    if (core.includes(",")) {
      integerPart = formatThousandsText(integerPart);
    }
    fractionPart = trimOptionalFractionDigits(fractionPart, minFractionDigits);
    numericText = `${integerPart}${fractionPart === "" ? "" : `.${fractionPart}`}`;
  }
  const combined = `${prefix}${numericText}${suffix}`;
  return autoNegative && !combined.startsWith("-") ? `-${combined}` : combined;
}

function formatTextBuiltinValue(value: CellValue, formatText: string): CellValue {
  const chosen = chooseFormatSection(value, formatText);
  if ("tag" in chosen) {
    return chosen;
  }
  const { section, numeric, autoNegative } = chosen;
  if (value.tag === ValueTag.String) {
    const cleaned = stripFormatDecorations(section);
    if (isTextFormat(section) || !/[0#?YMDHS]/i.test(cleaned)) {
      return stringResult(formatTextSectionValue(value.value, section));
    }
    return error(ErrorCode.Value);
  }
  if (numeric === undefined) {
    return error(ErrorCode.Value);
  }
  if (isDateTimeFormat(section)) {
    const formatted = formatDateTimeSectionValue(numeric, section);
    return formatted === undefined ? error(ErrorCode.Value) : stringResult(formatted);
  }
  return stringResult(formatNumericSectionValue(numeric, section, autoNegative));
}

function parseNumberValueText(
  input: string,
  decimalSeparator: string,
  groupSeparator: string,
): number | undefined {
  const compact = input.replaceAll(/\s+/g, "");
  if (compact === "") {
    return 0;
  }

  const percentMatch = compact.match(/%+$/);
  const percentCount = percentMatch?.[0].length ?? 0;
  const core = percentCount === 0 ? compact : compact.slice(0, -percentCount);
  if (core.includes("%")) {
    return undefined;
  }
  if (decimalSeparator !== "" && groupSeparator !== "" && decimalSeparator === groupSeparator) {
    return undefined;
  }

  const decimal = decimalSeparator === "" ? "." : decimalSeparator[0]!;
  const group = groupSeparator === "" ? "" : groupSeparator[0]!;

  const decimalIndex = decimal === "" ? -1 : core.indexOf(decimal);
  if (decimalIndex !== -1 && core.indexOf(decimal, decimalIndex + 1) !== -1) {
    return undefined;
  }

  let normalized = core;
  if (group !== "") {
    const groupAfterDecimal =
      decimalIndex === -1 ? -1 : normalized.indexOf(group, decimalIndex + decimal.length);
    if (groupAfterDecimal !== -1) {
      return undefined;
    }
    normalized = normalized.replaceAll(group, "");
  }
  if (decimal !== "." && decimal !== "") {
    normalized = normalized.replace(decimal, ".");
  }
  if (normalized === "" || normalized === "." || normalized === "+" || normalized === "-") {
    return undefined;
  }
  const parsed = Number(normalized);
  if (!Number.isFinite(parsed)) {
    return undefined;
  }
  return parsed / 100 ** percentCount;
}

const bahtDigitWords = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"] as const;
const bahtScaleWords = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"] as const;
const maxBahtTextSatang = Number.MAX_SAFE_INTEGER;

function bahtSegmentText(digits: string): string {
  const normalized = digits.replace(/^0+(?=\d)/u, "");
  if (normalized === "" || /^0+$/u.test(normalized)) {
    return "";
  }

  let output = "";
  let hasHigherNonZero = false;
  const length = normalized.length;
  for (let index = 0; index < length; index += 1) {
    const digit = normalized.charCodeAt(index) - 48;
    if (digit === 0) {
      continue;
    }

    const position = length - index - 1;
    if (position === 0) {
      output += digit === 1 && hasHigherNonZero ? "เอ็ด" : bahtDigitWords[digit]!;
    } else if (position === 1) {
      output += digit === 1 ? "สิบ" : digit === 2 ? "ยี่สิบ" : `${bahtDigitWords[digit]!}สิบ`;
    } else {
      output += `${bahtDigitWords[digit]!}${bahtScaleWords[position]!}`;
    }
    hasHigherNonZero = true;
  }
  return output;
}

function bahtIntegerText(digits: string): string {
  const normalized = digits.replace(/^0+(?=\d)/u, "");
  if (normalized === "" || /^0+$/u.test(normalized)) {
    return bahtDigitWords[0];
  }
  if (normalized.length > 6) {
    const head = bahtIntegerText(normalized.slice(0, -6));
    const tail = bahtSegmentText(normalized.slice(-6));
    return `${head}ล้าน${tail}`;
  }
  return bahtSegmentText(normalized) || bahtDigitWords[0];
}

function bahtTextFromNumber(value: number): CellValue {
  if (!Number.isFinite(value)) {
    return error(ErrorCode.Value);
  }

  const absolute = Math.abs(value);
  const scaled = Math.round(absolute * 100);
  if (!Number.isSafeInteger(scaled) || scaled > maxBahtTextSatang) {
    return error(ErrorCode.Value);
  }

  const baht = Math.trunc(scaled / 100);
  const satang = scaled % 100;
  const prefix = value < 0 ? "ลบ" : "";
  const bahtText = bahtIntegerText(String(baht));
  if (satang === 0) {
    return stringResult(`${prefix}${bahtText}บาทถ้วน`);
  }
  return stringResult(`${prefix}${bahtText}บาท${bahtSegmentText(String(satang))}สตางค์`);
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

const halfWidthKanaToFullWidthMap = new Map<string, string>([
  ["｡", "。"],
  ["｢", "「"],
  ["｣", "」"],
  ["､", "、"],
  ["･", "・"],
  ["ｦ", "ヲ"],
  ["ｧ", "ァ"],
  ["ｨ", "ィ"],
  ["ｩ", "ゥ"],
  ["ｪ", "ェ"],
  ["ｫ", "ォ"],
  ["ｬ", "ャ"],
  ["ｭ", "ュ"],
  ["ｮ", "ョ"],
  ["ｯ", "ッ"],
  ["ｰ", "ー"],
  ["ｱ", "ア"],
  ["ｲ", "イ"],
  ["ｳ", "ウ"],
  ["ｴ", "エ"],
  ["ｵ", "オ"],
  ["ｶ", "カ"],
  ["ｷ", "キ"],
  ["ｸ", "ク"],
  ["ｹ", "ケ"],
  ["ｺ", "コ"],
  ["ｻ", "サ"],
  ["ｼ", "シ"],
  ["ｽ", "ス"],
  ["ｾ", "セ"],
  ["ｿ", "ソ"],
  ["ﾀ", "タ"],
  ["ﾁ", "チ"],
  ["ﾂ", "ツ"],
  ["ﾃ", "テ"],
  ["ﾄ", "ト"],
  ["ﾅ", "ナ"],
  ["ﾆ", "ニ"],
  ["ﾇ", "ヌ"],
  ["ﾈ", "ネ"],
  ["ﾉ", "ノ"],
  ["ﾊ", "ハ"],
  ["ﾋ", "ヒ"],
  ["ﾌ", "フ"],
  ["ﾍ", "ヘ"],
  ["ﾎ", "ホ"],
  ["ﾏ", "マ"],
  ["ﾐ", "ミ"],
  ["ﾑ", "ム"],
  ["ﾒ", "メ"],
  ["ﾓ", "モ"],
  ["ﾔ", "ヤ"],
  ["ﾕ", "ユ"],
  ["ﾖ", "ヨ"],
  ["ﾗ", "ラ"],
  ["ﾘ", "リ"],
  ["ﾙ", "ル"],
  ["ﾚ", "レ"],
  ["ﾛ", "ロ"],
  ["ﾜ", "ワ"],
  ["ﾝ", "ン"],
]);

const halfWidthVoicedKanaToFullWidthMap = new Map<string, string>([
  ["ｳﾞ", "ヴ"],
  ["ｶﾞ", "ガ"],
  ["ｷﾞ", "ギ"],
  ["ｸﾞ", "グ"],
  ["ｹﾞ", "ゲ"],
  ["ｺﾞ", "ゴ"],
  ["ｻﾞ", "ザ"],
  ["ｼﾞ", "ジ"],
  ["ｽﾞ", "ズ"],
  ["ｾﾞ", "ゼ"],
  ["ｿﾞ", "ゾ"],
  ["ﾀﾞ", "ダ"],
  ["ﾁﾞ", "ヂ"],
  ["ﾂﾞ", "ヅ"],
  ["ﾃﾞ", "デ"],
  ["ﾄﾞ", "ド"],
  ["ﾊﾞ", "バ"],
  ["ﾋﾞ", "ビ"],
  ["ﾌﾞ", "ブ"],
  ["ﾍﾞ", "ベ"],
  ["ﾎﾞ", "ボ"],
  ["ﾊﾟ", "パ"],
  ["ﾋﾟ", "ピ"],
  ["ﾌﾟ", "プ"],
  ["ﾍﾟ", "ペ"],
  ["ﾎﾟ", "ポ"],
]);

const fullWidthKanaToHalfWidthMap = new Map<string, string>([
  ...[...halfWidthKanaToFullWidthMap.entries()].map(([half, full]) => [full, half] as const),
  ...[...halfWidthVoicedKanaToFullWidthMap.entries()].map(([half, full]) => [full, half] as const),
]);

function isLeadingSurrogate(codeUnit: number): boolean {
  return codeUnit >= 0xd800 && codeUnit <= 0xdbff;
}

function isTrailingSurrogate(codeUnit: number): boolean {
  return codeUnit >= 0xdc00 && codeUnit <= 0xdfff;
}

function toJapaneseFullWidth(input: string): string {
  let output = "";
  for (let index = 0; index < input.length; index += 1) {
    const char = input[index]!;
    const code = input.charCodeAt(index);

    if (isLeadingSurrogate(code) && index + 1 < input.length) {
      const nextCode = input.charCodeAt(index + 1);
      if (isTrailingSurrogate(nextCode)) {
        output += input.slice(index, index + 2);
        index += 1;
        continue;
      }
    }

    if (code === 0x20) {
      output += "\u3000";
      continue;
    }
    if (code >= 0x21 && code <= 0x7e) {
      output += String.fromCharCode(code + 0xfee0);
      continue;
    }

    const next = input[index + 1];
    if (next !== undefined) {
      const voiced = halfWidthVoicedKanaToFullWidthMap.get(char + next);
      if (voiced !== undefined) {
        output += voiced;
        index += 1;
        continue;
      }
    }

    const mapped = halfWidthKanaToFullWidthMap.get(char);
    output += mapped ?? char;
  }
  return output;
}

function toJapaneseHalfWidth(input: string): string {
  let output = "";
  for (let index = 0; index < input.length; index += 1) {
    const char = input[index]!;
    const code = input.charCodeAt(index);

    if (isLeadingSurrogate(code) && index + 1 < input.length) {
      const nextCode = input.charCodeAt(index + 1);
      if (isTrailingSurrogate(nextCode)) {
        output += input.slice(index, index + 2);
        index += 1;
        continue;
      }
    }

    if (code === 0x3000) {
      output += " ";
      continue;
    }
    if (code >= 0xff01 && code <= 0xff5e) {
      output += String.fromCharCode(code - 0xfee0);
      continue;
    }

    const mapped = fullWidthKanaToHalfWidthMap.get(char);
    output += mapped ?? char;
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
  LENB: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    return numberResult(utf8Bytes(coerceText(value)).length);
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
  ASC: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(toJapaneseHalfWidth(coerceText(textValue)));
  },
  JIS: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(toJapaneseFullWidth(coerceText(textValue)));
  },
  DBCS: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(toJapaneseFullWidth(coerceText(textValue)));
  },
  BAHTTEXT: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    const numeric = coerceNumber(value);
    return numeric === undefined ? error(ErrorCode.Value) : bahtTextFromNumber(numeric);
  },
  TEXT: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value, formatValue] = args;
    if (value === undefined || formatValue === undefined) {
      return error(ErrorCode.Value);
    }
    return formatTextBuiltinValue(value, coerceText(formatValue));
  },
  PHONETIC: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [value] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    return stringResult(coerceText(value));
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
  SEARCHB: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [findTextValue, withinTextValue, startValue] = args;
    if (findTextValue === undefined || withinTextValue === undefined) {
      return error(ErrorCode.Value);
    }
    const text = coerceText(withinTextValue);
    const start = coercePositiveStart(startValue, 1);
    if (isErrorValue(start)) {
      return start;
    }
    if (start > utf8Bytes(text).length + 1) {
      return error(ErrorCode.Value);
    }
    const found = findPosition(
      coerceText(findTextValue),
      text,
      bytePositionToCharPosition(text, start),
      false,
      true,
    );
    return isErrorValue(found) ? found : numberResult(charPositionToBytePosition(text, found));
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
  NUMBERVALUE: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, decimalSeparatorValue, groupSeparatorValue] = args;
    if (textValue === undefined) {
      return error(ErrorCode.Value);
    }
    const text = coerceText(textValue);
    const decimalSeparator =
      decimalSeparatorValue === undefined ? "." : coerceText(decimalSeparatorValue);
    const groupSeparator =
      groupSeparatorValue === undefined ? "," : coerceText(groupSeparatorValue);
    const parsed = parseNumberValueText(text, decimalSeparator, groupSeparator);
    return parsed === undefined ? error(ErrorCode.Value) : numberResult(parsed);
  },
  VALUETOTEXT: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return valueToTextResult(existingError, 0);
    }
    const [value, formatValue] = args;
    if (value === undefined) {
      return error(ErrorCode.Value);
    }
    const format = coerceInteger(formatValue, 0);
    if (isErrorValue(format)) {
      return format;
    }
    return valueToTextResult(value, format);
  },
  REGEXTEST: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, patternValue, caseSensitivityValue] = args;
    if (textValue === undefined || patternValue === undefined) {
      return error(ErrorCode.Value);
    }
    const caseSensitivity = coerceInteger(caseSensitivityValue, 0);
    if (isErrorValue(caseSensitivity) || (caseSensitivity !== 0 && caseSensitivity !== 1)) {
      return error(ErrorCode.Value);
    }
    const pattern = coerceText(patternValue);
    const regex = compileRegex(pattern, caseSensitivity);
    if (isRegexError(regex)) {
      return regex;
    }
    return booleanResult(regex.test(coerceText(textValue)));
  },
  REGEXREPLACE: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, patternValue, replacementValue, occurrenceValue, caseSensitivityValue] = args;
    if (textValue === undefined || patternValue === undefined || replacementValue === undefined) {
      return error(ErrorCode.Value);
    }
    const occurrence = coerceInteger(occurrenceValue, 0);
    const caseSensitivity = coerceInteger(caseSensitivityValue, 0);
    if (
      isErrorValue(occurrence) ||
      isErrorValue(caseSensitivity) ||
      (caseSensitivity !== 0 && caseSensitivity !== 1)
    ) {
      return error(ErrorCode.Value);
    }
    const text = coerceText(textValue);
    const replacement = coerceText(replacementValue);
    const regex = compileRegex(coerceText(patternValue), caseSensitivity, true);
    if (isRegexError(regex)) {
      return regex;
    }
    if (occurrence === 0) {
      return stringResult(text.replace(regex, replacement));
    }
    const matches = [...text.matchAll(regex)];
    if (matches.length === 0) {
      return stringResult(text);
    }
    const targetIndex = occurrence > 0 ? occurrence - 1 : matches.length + occurrence;
    if (targetIndex < 0 || targetIndex >= matches.length) {
      return stringResult(text);
    }
    let currentIndex = -1;
    return stringResult(
      text.replace(regex, (match, ...rest) => {
        currentIndex += 1;
        if (currentIndex !== targetIndex) {
          return match;
        }
        const captures = rest
          .slice(0, -2)
          .map((value) => (typeof value === "string" ? value : undefined));
        return applyReplacementTemplate(replacement, match, captures);
      }),
    );
  },
  REGEXEXTRACT: (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const [textValue, patternValue, returnModeValue, caseSensitivityValue] = args;
    if (textValue === undefined || patternValue === undefined) {
      return error(ErrorCode.Value);
    }
    const returnMode = coerceInteger(returnModeValue, 0);
    const caseSensitivity = coerceInteger(caseSensitivityValue, 0);
    if (
      isErrorValue(returnMode) ||
      isErrorValue(caseSensitivity) ||
      ![0, 1, 2].includes(returnMode) ||
      (caseSensitivity !== 0 && caseSensitivity !== 1)
    ) {
      return error(ErrorCode.Value);
    }

    const text = coerceText(textValue);
    const pattern = coerceText(patternValue);
    if (returnMode === 1) {
      const regex = compileRegex(pattern, caseSensitivity, true);
      if (isRegexError(regex)) {
        return regex;
      }
      const matches = [...text.matchAll(regex)].map((entry) => entry[0]);
      if (matches.length === 0) {
        return error(ErrorCode.NA);
      }
      return {
        kind: "array",
        rows: matches.length,
        cols: 1,
        values: matches.map((match) => stringResult(match)),
      };
    }

    const regex = compileRegex(pattern, caseSensitivity, false);
    if (isRegexError(regex)) {
      return regex;
    }
    const match = text.match(regex);
    if (!match) {
      return error(ErrorCode.NA);
    }
    if (returnMode === 0) {
      return stringResult(match[0]);
    }
    const groups = match.slice(1);
    if (groups.length === 0) {
      return error(ErrorCode.NA);
    }
    return {
      kind: "array",
      rows: 1,
      cols: groups.length,
      values: groups.map((group) => stringResult(group ?? "")),
    };
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
    if (typeof ignoreEmpty !== "boolean") {
      return ignoreEmpty;
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
  REPLACEB: (...args) => {
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
    return stringResult(
      replaceBytes(coerceText(textValue), start, count, coerceText(replacementValue)),
    );
  },
  SUBSTITUTE: createSubstituteBuiltin(),
  REPT: createReptBuiltin(),
  ...textPlaceholderBuiltins,
};

export function getTextBuiltin(name: string): TextBuiltin | undefined {
  return textBuiltins[name.toUpperCase()];
}
