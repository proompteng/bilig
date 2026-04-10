import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { excelSerialToDateParts } from "./datetime.js";
import type { TextBuiltin } from "./text.js";

interface TextFormatBuiltinDeps {
  error: (code: ErrorCode) => CellValue;
  stringResult: (value: string) => CellValue;
  numberResult: (value: number) => CellValue;
  firstError: (args: readonly (CellValue | undefined)[]) => CellValue | undefined;
  coerceText: (value: CellValue) => string;
  coerceNumber: (value: CellValue) => number | undefined;
  coerceInteger: (value: CellValue | undefined, defaultValue: number) => number | CellValue;
  isErrorValue: (value: number | CellValue) => value is CellValue;
}

function valueToTextResult(
  deps: TextFormatBuiltinDeps,
  value: CellValue,
  format: number,
): CellValue {
  if (format !== 0 && format !== 1) {
    return deps.error(ErrorCode.Value);
  }
  switch (value.tag) {
    case ValueTag.Empty:
      return deps.stringResult("");
    case ValueTag.Number:
      return deps.stringResult(String(value.value));
    case ValueTag.Boolean:
      return deps.stringResult(value.value ? "TRUE" : "FALSE");
    case ValueTag.String:
      return deps.stringResult(format === 1 ? JSON.stringify(value.value) : value.value);
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
      return deps.stringResult(label);
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
  deps: TextFormatBuiltinDeps,
  value: CellValue,
  formatText: string,
): { section: string; numeric?: number; autoNegative: boolean } | CellValue {
  const sections = splitFormatSections(formatText);
  if (value.tag === ValueTag.String) {
    return { section: sections[3] ?? sections[0] ?? "", autoNegative: false };
  }
  const numeric = deps.coerceNumber(value);
  if (numeric === undefined) {
    return deps.error(ErrorCode.Value);
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

function formatTextBuiltinValue(
  deps: TextFormatBuiltinDeps,
  value: CellValue,
  formatText: string,
): CellValue {
  const chosen = chooseFormatSection(deps, value, formatText);
  if ("tag" in chosen) {
    return chosen;
  }
  const { section, numeric, autoNegative } = chosen;
  if (value.tag === ValueTag.String) {
    const cleaned = stripFormatDecorations(section);
    if (isTextFormat(section) || !/[0#?YMDHS]/i.test(cleaned)) {
      return deps.stringResult(formatTextSectionValue(value.value, section));
    }
    return deps.error(ErrorCode.Value);
  }
  if (numeric === undefined) {
    return deps.error(ErrorCode.Value);
  }
  if (isDateTimeFormat(section)) {
    const formatted = formatDateTimeSectionValue(numeric, section);
    return formatted === undefined ? deps.error(ErrorCode.Value) : deps.stringResult(formatted);
  }
  return deps.stringResult(formatNumericSectionValue(numeric, section, autoNegative));
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

export function createTextFormatBuiltins(
  deps: TextFormatBuiltinDeps,
): Record<string, TextBuiltin> {
  return {
    TEXT: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [value, formatValue] = args;
      if (value === undefined || formatValue === undefined) {
        return deps.error(ErrorCode.Value);
      }
      return formatTextBuiltinValue(deps, value, deps.coerceText(formatValue));
    },
    VALUE: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [value] = args;
      if (value === undefined) {
        return deps.error(ErrorCode.Value);
      }
      const coerced = deps.coerceNumber(value);
      return coerced === undefined ? deps.error(ErrorCode.Value) : deps.numberResult(coerced);
    },
    NUMBERVALUE: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return existingError;
      }
      const [textValue, decimalSeparatorValue, groupSeparatorValue] = args;
      if (textValue === undefined) {
        return deps.error(ErrorCode.Value);
      }
      const text = deps.coerceText(textValue);
      const decimalSeparator =
        decimalSeparatorValue === undefined ? "." : deps.coerceText(decimalSeparatorValue);
      const groupSeparator =
        groupSeparatorValue === undefined ? "," : deps.coerceText(groupSeparatorValue);
      const parsed = parseNumberValueText(text, decimalSeparator, groupSeparator);
      return parsed === undefined ? deps.error(ErrorCode.Value) : deps.numberResult(parsed);
    },
    VALUETOTEXT: (...args) => {
      const existingError = deps.firstError(args);
      if (existingError) {
        return valueToTextResult(deps, existingError, 0);
      }
      const [value, formatValue] = args;
      if (value === undefined) {
        return deps.error(ErrorCode.Value);
      }
      const format = deps.coerceInteger(formatValue, 0);
      if (deps.isErrorValue(format)) {
        return format;
      }
      return valueToTextResult(deps, value, format);
    },
  };
}
