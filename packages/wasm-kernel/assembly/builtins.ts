import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import {
  allocateOutputString,
  allocateSpillArrayResult,
  encodeOutputStringId,
  nextVolatileRandomValue,
  readSpillArrayTag,
  readSpillArrayLength,
  readSpillArrayNumber,
  readVolatileNowSerial,
  writeOutputStringData,
  writeSpillArrayNumber,
  writeSpillArrayValue,
} from "./vm";

export const STACK_KIND_SCALAR: u8 = 0;
export const STACK_KIND_RANGE: u8 = 1;
export const STACK_KIND_ARRAY: u8 = 2;
const UNRESOLVED_WASM_OPERAND: u32 = 0x00ffffff;
const DEFAULT_DYNAMIC_ARRAY_SHAPE_CAPACITY = 16;
let dynamicArrayShapeRows = new Uint32Array(DEFAULT_DYNAMIC_ARRAY_SHAPE_CAPACITY);
let dynamicArrayShapeCols = new Uint32Array(DEFAULT_DYNAMIC_ARRAY_SHAPE_CAPACITY);
let dynamicArrayShapeValid = new Uint8Array(DEFAULT_DYNAMIC_ARRAY_SHAPE_CAPACITY);

function ensureDynamicArrayShapeCapacity(required: i32): void {
  if (required <= dynamicArrayShapeRows.length) {
    return;
  }
  let nextCapacity = dynamicArrayShapeRows.length;
  while (nextCapacity < required) {
    nextCapacity *= 2;
  }
  const nextRows = new Uint32Array(nextCapacity);
  const nextCols = new Uint32Array(nextCapacity);
  const nextValid = new Uint8Array(nextCapacity);
  for (let index = 0; index < dynamicArrayShapeRows.length; index++) {
    nextRows[index] = dynamicArrayShapeRows[index];
    nextCols[index] = dynamicArrayShapeCols[index];
    nextValid[index] = dynamicArrayShapeValid[index];
  }
  dynamicArrayShapeRows = nextRows;
  dynamicArrayShapeCols = nextCols;
  dynamicArrayShapeValid = nextValid;
}

function registerDynamicArrayShape(arrayIndex: u32, rows: i32, cols: i32): void {
  const index = <i32>arrayIndex;
  ensureDynamicArrayShapeCapacity(index + 1);
  dynamicArrayShapeRows[index] = <u32>rows;
  dynamicArrayShapeCols[index] = <u32>cols;
  dynamicArrayShapeValid[index] = 1;
}

function getDynamicArrayRows(arrayIndex: u32): i32 {
  const index = <i32>arrayIndex;
  if (index < 0 || index >= dynamicArrayShapeValid.length || dynamicArrayShapeValid[index] == 0) {
    return i32.MIN_VALUE;
  }
  return <i32>dynamicArrayShapeRows[index];
}

function getDynamicArrayCols(arrayIndex: u32): i32 {
  const index = <i32>arrayIndex;
  if (index < 0 || index >= dynamicArrayShapeValid.length || dynamicArrayShapeValid[index] == 0) {
    return i32.MIN_VALUE;
  }
  return <i32>dynamicArrayShapeCols[index];
}

function toNumberOrNaN(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value;
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

function toNumberOrZero(tag: u8, value: f64): f64 {
  const numeric = toNumberOrNaN(tag, value);
  return isNaN(numeric) ? 0 : numeric;
}

function toNumberExact(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value;
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

const OUTPUT_STRING_BASE: f64 = 2147483648.0;

function volatileNowResult(): f64 {
  return readVolatileNowSerial();
}

function outputStringIndex(value: f64): i32 {
  if (value < OUTPUT_STRING_BASE) {
    return -1;
  }
  return <i32>(value - OUTPUT_STRING_BASE);
}

function textLength(
  tag: u8,
  value: f64,
  stringLengths: Uint32Array,
  outputStringLengths: Uint32Array,
): i32 {
  if (tag == ValueTag.Empty) {
    return 0;
  }
  if (tag == ValueTag.Boolean) {
    return value != 0 ? 4 : 5;
  }
  if (tag == ValueTag.Number) {
    return value.toString().length;
  }
  if (tag == ValueTag.String) {
    const outputIndex = outputStringIndex(value);
    if (outputIndex >= 0) {
      const index = outputIndex;
      if (index < 0 || index >= outputStringLengths.length) {
        return -1;
      }
      return <i32>outputStringLengths[index];
    }
    const stringId = <i32>value;
    if (stringId < 0 || stringId >= stringLengths.length) {
      return -1;
    }
    return <i32>stringLengths[stringId];
  }
  return -1;
}

function poolString(
  stringId: i32,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
): string | null {
  if (stringId < 0 || stringId >= stringLengths.length) {
    return null;
  }
  const offset = <i32>stringOffsets[stringId];
  const length = <i32>stringLengths[stringId];
  let text = "";
  for (let index = 0; index < length; index++) {
    text += String.fromCharCode(stringData[offset + index]);
  }
  return text;
}

function scalarText(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  if (tag == ValueTag.Empty) {
    return "";
  }
  if (tag == ValueTag.Boolean) {
    return value != 0 ? "TRUE" : "FALSE";
  }
  if (tag == ValueTag.Number) {
    return value.toString();
  }
  if (tag == ValueTag.String) {
    const outputIndex = outputStringIndex(value);
    if (outputIndex >= 0) {
      const index = outputIndex;
      if (index < 0 || index >= outputStringLengths.length) return null;
      const offset = <i32>outputStringOffsets[index];
      const length = <i32>outputStringLengths[index];
      let text = "";
      for (let i = 0; i < length; i++) {
        text += String.fromCharCode(outputStringData[offset + i]);
      }
      return text;
    }
    const stringId = <i32>value;
    return poolString(stringId, stringOffsets, stringLengths, stringData);
  }
  return null;
}

function arrayToTextCell(
  tag: u8,
  value: f64,
  strict: bool,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  if (tag == ValueTag.Number) {
    return value == Math.trunc(value) ? (<i64>value).toString() : value.toString();
  }
  const text = scalarText(
    tag,
    value,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (text == null) {
    return null;
  }
  if (!strict || tag != ValueTag.String) {
    return text;
  }
  return '"' + substituteText(text, '"', '""') + '"';
}

function truncToInt(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  return isNaN(numeric) ? i32.MIN_VALUE : <i32>numeric;
}

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  return -1;
}

function floorDiv(a: i32, b: i32): i32 {
  let quotient = a / b;
  const remainder = a % b;
  if (remainder != 0 && remainder > 0 != b > 0) {
    quotient -= 1;
  }
  return quotient;
}

function daysFromCivil(year: i32, month: i32, day: i32): i32 {
  let adjustedYear = year;
  if (month <= 2) {
    adjustedYear -= 1;
  }
  const era = adjustedYear >= 0 ? adjustedYear / 400 : (adjustedYear - 399) / 400;
  const yearOfEra = adjustedYear - era * 400;
  const shiftedMonth = month + (month > 2 ? -3 : 9);
  const dayOfYear = (153 * shiftedMonth + 2) / 5 + day - 1;
  const dayOfEra = yearOfEra * 365 + yearOfEra / 4 - yearOfEra / 100 + dayOfYear;
  return era * 146097 + dayOfEra - 719468;
}

function civilYear(days: i32): i32 {
  const shifted = days + 719468;
  const era = shifted >= 0 ? shifted / 146097 : (shifted - 146096) / 146097;
  const dayOfEra = shifted - era * 146097;
  const yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365;
  const dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100);
  const monthPrime = (5 * dayOfYear + 2) / 153;
  const month = monthPrime + (monthPrime < 10 ? 3 : -9);
  return yearOfEra + era * 400 + (month <= 2 ? 1 : 0);
}

function civilMonth(days: i32): i32 {
  const shifted = days + 719468;
  const era = shifted >= 0 ? shifted / 146097 : (shifted - 146096) / 146097;
  const dayOfEra = shifted - era * 146097;
  const yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365;
  const dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100);
  const monthPrime = (5 * dayOfYear + 2) / 153;
  return monthPrime + (monthPrime < 10 ? 3 : -9);
}

function civilDay(days: i32): i32 {
  const shifted = days + 719468;
  const era = shifted >= 0 ? shifted / 146097 : (shifted - 146096) / 146097;
  const dayOfEra = shifted - era * 146097;
  const yearOfEra = (dayOfEra - dayOfEra / 1460 + dayOfEra / 36524 - dayOfEra / 146096) / 365;
  const dayOfYear = dayOfEra - (365 * yearOfEra + yearOfEra / 4 - yearOfEra / 100);
  const monthPrime = (5 * dayOfYear + 2) / 153;
  return dayOfYear - (153 * monthPrime + 2) / 5 + 1;
}

const EXCEL_EPOCH_DAYS: i32 = -25568;
const EXCEL_SECONDS_PER_DAY: i32 = 86400;

function excelSerialWhole(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  return isNaN(numeric) ? i32.MIN_VALUE : <i32>Math.floor(numeric);
}

function daysInExcelMonth(year: i32, month: i32): i32 {
  if (year == 1900 && month == 2) {
    return 29;
  }
  const start = daysFromCivil(year, month, 1);
  const nextMonth = month == 12 ? 1 : month + 1;
  const nextYear = month == 12 ? year + 1 : year;
  const end = daysFromCivil(nextYear, nextMonth, 1);
  return end - start;
}

function excelYearPartFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value);
  if (whole == i32.MIN_VALUE) {
    return i32.MIN_VALUE;
  }
  if (whole == 60) {
    return 1900;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  return civilYear(EXCEL_EPOCH_DAYS + adjustedWhole);
}

function excelMonthPartFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value);
  if (whole == i32.MIN_VALUE) {
    return i32.MIN_VALUE;
  }
  if (whole == 60) {
    return 2;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  return civilMonth(EXCEL_EPOCH_DAYS + adjustedWhole);
}

function excelDayPartFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value);
  if (whole == i32.MIN_VALUE) {
    return i32.MIN_VALUE;
  }
  if (whole == 60) {
    return 29;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  return civilDay(EXCEL_EPOCH_DAYS + adjustedWhole);
}

function excelTimeSerial(
  hourTag: u8,
  hourValue: f64,
  minuteTag: u8,
  minuteValue: f64,
  secondTag: u8,
  secondValue: f64,
): f64 {
  const hourNumeric = toNumberExact(hourTag, hourValue);
  const minuteNumeric = toNumberExact(minuteTag, minuteValue);
  const secondNumeric = toNumberExact(secondTag, secondValue);
  if (isNaN(hourNumeric) || isNaN(minuteNumeric) || isNaN(secondNumeric)) {
    return NaN;
  }
  const hour = <f64>(<i32>hourNumeric);
  const minute = <f64>(<i32>minuteNumeric);
  const second = <f64>(<i32>secondNumeric);
  if (hour < 0 || minute < 0 || second < 0 || hour > 32767 || minute > 32767 || second > 32767) {
    return NaN;
  }
  let totalSeconds = hour * 3600.0 + minute * 60.0 + second;
  totalSeconds %= <f64>EXCEL_SECONDS_PER_DAY;
  if (totalSeconds < 0) {
    totalSeconds += <f64>EXCEL_SECONDS_PER_DAY;
  }
  return totalSeconds / <f64>EXCEL_SECONDS_PER_DAY;
}

function excelSecondOfDay(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric) || numeric < 0) {
    return i32.MIN_VALUE;
  }
  const whole = Math.floor(numeric);
  let fraction = numeric - whole;
  if (fraction < 0) {
    fraction += 1.0;
  }
  let seconds = <i32>Math.floor(fraction * <f64>EXCEL_SECONDS_PER_DAY + 1e-9);
  if (seconds >= EXCEL_SECONDS_PER_DAY) {
    seconds = 0;
  }
  return seconds;
}

function excelWeekdayFromSerial(tag: u8, value: f64, returnType: i32): i32 {
  const whole = excelSerialWhole(tag, value);
  if (whole == i32.MIN_VALUE || whole < 0) {
    return i32.MIN_VALUE;
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  const sundayOne = (((adjustedWhole % 7) + 7) % 7) + 1;
  if (returnType == 1) {
    return sundayOne;
  }
  if (returnType == 3) {
    return sundayOne == 1 ? 6 : sundayOne - 2;
  }

  let startDay = 0;
  if (returnType == 2 || returnType == 11) {
    startDay = 2;
  } else if (returnType == 12) {
    startDay = 3;
  } else if (returnType == 13) {
    startDay = 4;
  } else if (returnType == 14) {
    startDay = 5;
  } else if (returnType == 15) {
    startDay = 6;
  } else if (returnType == 16) {
    startDay = 7;
  } else if (returnType == 17) {
    startDay = 1;
  } else {
    return i32.MIN_VALUE;
  }

  return ((sundayOne - startDay + 7) % 7) + 1;
}

function excelWeeknumFromSerial(tag: u8, value: f64, returnType: i32): i32 {
  const year = excelYearPartFromSerial(tag, value);
  const month = excelMonthPartFromSerial(tag, value);
  const day = excelDayPartFromSerial(tag, value);
  if (year == i32.MIN_VALUE || month == i32.MIN_VALUE || day == i32.MIN_VALUE) {
    return i32.MIN_VALUE;
  }

  let weekStartDay = 0;
  if (returnType == 1 || returnType == 17) {
    weekStartDay = 0;
  } else if (returnType == 2 || returnType == 11) {
    weekStartDay = 1;
  } else if (returnType == 12) {
    weekStartDay = 2;
  } else if (returnType == 13) {
    weekStartDay = 3;
  } else if (returnType == 14) {
    weekStartDay = 4;
  } else if (returnType == 15) {
    weekStartDay = 5;
  } else if (returnType == 16) {
    weekStartDay = 6;
  } else {
    return i32.MIN_VALUE;
  }

  const jan1Serial = excelDateSerial(
    <u8>ValueTag.Number,
    <f64>year,
    <u8>ValueTag.Number,
    1.0,
    <u8>ValueTag.Number,
    1.0,
  );
  if (isNaN(jan1Serial)) {
    return i32.MIN_VALUE;
  }

  let adjustedJan1 = <i32>Math.floor(jan1Serial);
  adjustedJan1 = adjustedJan1 < 60 ? adjustedJan1 : adjustedJan1 - 1;
  const jan1Weekday = ((adjustedJan1 % 7) + 7) % 7;
  const shift = (jan1Weekday - weekStartDay + 7) % 7;

  let dayOfYear = day;
  for (let currentMonth = 1; currentMonth < month; currentMonth += 1) {
    dayOfYear += daysInExcelMonth(year, currentMonth);
  }

  return <i32>Math.floor(<f64>(dayOfYear - 1 + shift) / 7.0) + 1;
}

function excelDateSerial(
  yearTag: u8,
  yearValue: f64,
  monthTag: u8,
  monthValue: f64,
  dayTag: u8,
  dayValue: f64,
): f64 {
  let year = truncToInt(yearTag, yearValue);
  const month = truncToInt(monthTag, monthValue);
  const day = truncToInt(dayTag, dayValue);
  if (year == i32.MIN_VALUE || month == i32.MIN_VALUE || day == i32.MIN_VALUE) {
    return NaN;
  }
  if (year >= 0 && year <= 1899) {
    year += 1900;
  }
  if (year < 0 || year > 9999) {
    return NaN;
  }
  if (year == 1900 && month == 2 && day == 29) {
    return 60;
  }

  const zeroBasedMonth = month - 1;
  const monthQuotient = floorDiv(zeroBasedMonth, 12);
  const normalizedYear = year + monthQuotient;
  const normalizedMonthZero = zeroBasedMonth - monthQuotient * 12;
  if (normalizedYear < 0 || normalizedYear > 9999) {
    return NaN;
  }
  const days = daysFromCivil(normalizedYear, normalizedMonthZero + 1, 1) + (day - 1);
  let serial = days - EXCEL_EPOCH_DAYS;
  if (days >= daysFromCivil(1900, 3, 1)) {
    serial += 1;
  }
  return <f64>serial;
}

function addMonthsExcelSerial(
  tag: u8,
  value: f64,
  offsetTag: u8,
  offsetValue: f64,
  endOfMonth: bool,
): f64 {
  const startYear = excelYearPartFromSerial(tag, value);
  const startMonth = excelMonthPartFromSerial(tag, value);
  const startDay = excelDayPartFromSerial(tag, value);
  const offset = truncToInt(offsetTag, offsetValue);
  if (
    startYear == i32.MIN_VALUE ||
    startMonth == i32.MIN_VALUE ||
    startDay == i32.MIN_VALUE ||
    offset == i32.MIN_VALUE
  ) {
    return NaN;
  }

  const totalMonths = startYear * 12 + (startMonth - 1) + offset;
  const shiftedYear = floorDiv(totalMonths, 12);
  const shiftedMonth = totalMonths - shiftedYear * 12 + 1;
  if (shiftedYear < 0 || shiftedYear > 9999) {
    return NaN;
  }

  const targetDay = endOfMonth
    ? daysInExcelMonth(shiftedYear, shiftedMonth)
    : min<i32>(startDay, daysInExcelMonth(shiftedYear, shiftedMonth));
  if (shiftedYear == 1900 && shiftedMonth == 2 && targetDay == 29) {
    return 60;
  }
  return excelDateSerial(
    <u8>ValueTag.Number,
    <f64>shiftedYear,
    <u8>ValueTag.Number,
    <f64>shiftedMonth,
    <u8>ValueTag.Number,
    <f64>targetDay,
  );
}

function roundToDigits(value: f64, digits: i32): f64 {
  if (digits >= 0) {
    const factor = Math.pow(10.0, <f64>digits);
    return Math.round(value * factor) / factor;
  }
  const factor = Math.pow(10.0, <f64>-digits);
  return Math.round(value / factor) * factor;
}

function writeStringResult(
  base: i32,
  text: string,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  const outputStringId = allocateOutputString(text.length);
  for (let index = 0; index < text.length; index++) {
    writeOutputStringData(outputStringId, index, <u16>text.charCodeAt(index));
  }
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    <u8>ValueTag.String,
    encodeOutputStringId(outputStringId),
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
}

function coerceLength(tag: u8, value: f64, defaultValue: i32): i32 {
  if (tag == ValueTag.Empty) {
    return defaultValue;
  }
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 0 ? truncated : i32.MIN_VALUE;
}

function coercePositiveStart(tag: u8, value: f64, defaultValue: i32): i32 {
  if (tag == ValueTag.Empty) {
    return defaultValue;
  }
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 1 ? truncated : i32.MIN_VALUE;
}

function excelTrim(input: string): string {
  let start = 0;
  let end = input.length;
  while (start < end && input.charCodeAt(start) == 32) {
    start += 1;
  }
  while (end > start && input.charCodeAt(end - 1) == 32) {
    end -= 1;
  }
  let result = "";
  let previousSpace = false;
  for (let index = start; index < end; index++) {
    const char = input.charCodeAt(index);
    if (char == 32) {
      if (!previousSpace) {
        result += " ";
      }
      previousSpace = true;
      continue;
    }
    previousSpace = false;
    result += String.fromCharCode(char);
  }
  return result;
}

function hasSearchSyntax(pattern: string): bool {
  for (let index = 0; index < pattern.length; index++) {
    const char = pattern.charCodeAt(index);
    if (char == 126 || char == 42 || char == 63) {
      return true;
    }
  }
  return false;
}

function wildcardMatchAt(
  pattern: string,
  haystack: string,
  patternIndex: i32,
  haystackIndex: i32,
): bool {
  let p = patternIndex;
  let h = haystackIndex;
  while (p < pattern.length) {
    const char = pattern.charCodeAt(p);
    if (char == 126) {
      const nextIndex = p + 1;
      const nextChar = nextIndex < pattern.length ? pattern.charCodeAt(nextIndex) : 126;
      if (h >= haystack.length || haystack.charCodeAt(h) != nextChar) {
        return false;
      }
      p = nextIndex < pattern.length ? nextIndex + 1 : nextIndex;
      h += 1;
      continue;
    }
    if (char == 42) {
      let nextPatternIndex = p + 1;
      while (nextPatternIndex < pattern.length && pattern.charCodeAt(nextPatternIndex) == 42) {
        nextPatternIndex += 1;
      }
      if (nextPatternIndex >= pattern.length) {
        return true;
      }
      for (let scan = h; scan <= haystack.length; scan++) {
        if (wildcardMatchAt(pattern, haystack, nextPatternIndex, scan)) {
          return true;
        }
      }
      return false;
    }
    if (h >= haystack.length) {
      return false;
    }
    if (char == 63) {
      p += 1;
      h += 1;
      continue;
    }
    if (haystack.charCodeAt(h) != char) {
      return false;
    }
    p += 1;
    h += 1;
  }
  return true;
}

function findPosition(
  needle: string,
  haystack: string,
  start: i32,
  caseSensitive: bool,
  wildcardAware: bool,
): i32 {
  const startIndex = start - 1;
  if (needle.length == 0) {
    return start;
  }
  if (startIndex > haystack.length) {
    return i32.MIN_VALUE;
  }
  const normalizedHaystack = caseSensitive ? haystack : haystack.toLowerCase();
  const normalizedNeedle = caseSensitive ? needle : needle.toLowerCase();
  if (wildcardAware && hasSearchSyntax(normalizedNeedle)) {
    for (let index = startIndex; index <= normalizedHaystack.length; index++) {
      if (wildcardMatchAt(normalizedNeedle, normalizedHaystack, 0, index)) {
        return index + 1;
      }
    }
    return i32.MIN_VALUE;
  }
  const found = normalizedHaystack.indexOf(normalizedNeedle, startIndex);
  return found < 0 ? i32.MIN_VALUE : found + 1;
}

function trimAsciiWhitespace(input: string): string {
  let start = 0;
  let end = input.length;
  while (start < end && input.charCodeAt(start) <= 32) {
    start += 1;
  }
  while (end > start && input.charCodeAt(end - 1) <= 32) {
    end -= 1;
  }
  return input.slice(start, end);
}

function parseNumericText(input: string): f64 {
  const text = trimAsciiWhitespace(input);
  if (text.length == 0) {
    return 0;
  }

  let index = 0;
  let sign = 1.0;
  const first = text.charCodeAt(index);
  if (first == 43) {
    index += 1;
  } else if (first == 45) {
    sign = -1.0;
    index += 1;
  }

  let value = 0.0;
  let digitCount = 0;
  while (index < text.length) {
    const char = text.charCodeAt(index);
    if (char < 48 || char > 57) {
      break;
    }
    value = value * 10.0 + <f64>(char - 48);
    digitCount += 1;
    index += 1;
  }

  if (index < text.length && text.charCodeAt(index) == 46) {
    index += 1;
    let factor = 0.1;
    while (index < text.length) {
      const char = text.charCodeAt(index);
      if (char < 48 || char > 57) {
        break;
      }
      value += <f64>(char - 48) * factor;
      factor *= 0.1;
      digitCount += 1;
      index += 1;
    }
  }

  if (digitCount == 0) {
    return NaN;
  }

  if (index < text.length) {
    const exponentMarker = text.charCodeAt(index);
    if (exponentMarker == 69 || exponentMarker == 101) {
      index += 1;
      let exponentSign = 1;
      if (index < text.length) {
        const exponentPrefix = text.charCodeAt(index);
        if (exponentPrefix == 43) {
          index += 1;
        } else if (exponentPrefix == 45) {
          exponentSign = -1;
          index += 1;
        }
      }

      let exponent = 0;
      let exponentDigits = 0;
      while (index < text.length) {
        const char = text.charCodeAt(index);
        if (char < 48 || char > 57) {
          break;
        }
        exponent = exponent * 10 + (char - 48);
        exponentDigits += 1;
        index += 1;
      }
      if (exponentDigits == 0) {
        return NaN;
      }
      value *= Math.pow(10.0, <f64>(exponentSign * exponent));
    }
  }

  if (index != text.length) {
    return NaN;
  }

  const parsed = sign * value;
  if (parsed == Infinity || parsed == -Infinity) {
    return NaN;
  }
  return parsed;
}

function isWeekendSerial(whole: i32): bool {
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  const dayOfWeek = ((adjustedWhole % 7) + 7) % 7;
  return dayOfWeek == 0 || dayOfWeek == 6;
}

function isHolidaySerial(
  serial: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndex: u32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  if (kind == STACK_KIND_SCALAR) {
    if (tag == ValueTag.Error) {
      return -1;
    }
    const holiday = truncToInt(tag, value);
    return holiday == i32.MIN_VALUE ? -1 : holiday == serial ? 1 : 0;
  }
  if (kind != STACK_KIND_RANGE) {
    return 0;
  }
  const start = <i32>rangeOffsets[rangeIndex];
  const length = <i32>rangeLengths[rangeIndex];
  for (let index = 0; index < length; index += 1) {
    const memberIndex = rangeMembers[start + index];
    if (cellTags[memberIndex] == ValueTag.Error) {
      return -1;
    }
    const serialCandidate = truncToInt(
      cellTags[memberIndex],
      memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
    );
    if (serialCandidate == i32.MIN_VALUE) {
      return -1;
    }
    if (serialCandidate == serial) {
      return 1;
    }
  }
  return 0;
}

function isWorkdaySerial(
  serial: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndex: u32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  if (isWeekendSerial(serial)) {
    return 0;
  }
  if (kind == STACK_KIND_SCALAR && tag == ValueTag.Empty) {
    return 1;
  }
  if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE) {
    return 1;
  }
  const holiday = isHolidaySerial(
    serial,
    kind,
    tag,
    value,
    rangeIndex,
    rangeOffsets,
    rangeLengths,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  );
  if (holiday < 0) {
    return -1;
  }
  return holiday == 1 ? 0 : 1;
}

function coerceNonNegativeLength(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 0 ? truncated : i32.MIN_VALUE;
}

function coerceInteger(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE;
  }
  return <i32>numeric;
}

function clipIndex(index: i32, length: i32): i32 {
  if (length <= 0) {
    return i32.MIN_VALUE;
  }
  if (index == 0) {
    return i32.MIN_VALUE;
  }
  return index < 0 ? max(index, -length) : min(index, length);
}

function inputRowsFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeRowCounts: Uint32Array,
): i32 {
  const kind = kindStack[slot];
  if (kind == STACK_KIND_SCALAR) {
    return 1;
  }
  if (kind == STACK_KIND_RANGE) {
    return <i32>rangeRowCounts[rangeIndexStack[slot]];
  }
  if (kind == STACK_KIND_ARRAY) {
    return getDynamicArrayRows(rangeIndexStack[slot]);
  }
  return i32.MIN_VALUE;
}

function inputColsFromSlot(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  const kind = kindStack[slot];
  if (kind == STACK_KIND_SCALAR) {
    return 1;
  }
  if (kind == STACK_KIND_RANGE) {
    return <i32>rangeColCounts[rangeIndexStack[slot]];
  }
  if (kind == STACK_KIND_ARRAY) {
    return getDynamicArrayCols(rangeIndexStack[slot]);
  }
  return i32.MIN_VALUE;
}

function inputCellTag(
  slot: i32,
  row: i32,
  col: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
): u8 {
  const kind = kindStack[slot];
  if (row < 0 || col < 0) {
    return <u8>ValueTag.Error;
  }
  if (kind == STACK_KIND_SCALAR) {
    if (row != 0 || col != 0) {
      return <u8>ValueTag.Error;
    }
    return tagStack[slot];
  }
  if (kind == STACK_KIND_RANGE) {
    const rangeIndex = rangeIndexStack[slot];
    const memberIndex = rangeMemberAt(
      rangeIndex,
      row,
      col,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (memberIndex == 0xffffffff) {
      return <u8>ValueTag.Error;
    }
    return cellTags[memberIndex];
  }
  if (kind == STACK_KIND_ARRAY) {
    const arrayIndex = rangeIndexStack[slot];
    const arrayRows = getDynamicArrayRows(arrayIndex);
    const arrayCols = getDynamicArrayCols(arrayIndex);
    if (arrayRows < 1 || arrayCols < 1 || row >= arrayRows || col >= arrayCols) {
      return <u8>ValueTag.Error;
    }
    const arrayOffset = row * arrayCols + col;
    if (arrayOffset >= readSpillArrayLength(arrayIndex)) {
      return <u8>ValueTag.Error;
    }
    return readSpillArrayTag(arrayIndex, arrayOffset);
  }
  return <u8>ValueTag.Error;
}

function inputCellScalarValue(
  slot: i32,
  row: i32,
  col: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): f64 {
  const kind = kindStack[slot];
  if (row < 0 || col < 0) {
    return NaN;
  }
  if (kind == STACK_KIND_SCALAR) {
    return row == 0 && col == 0 ? valueStack[slot] : NaN;
  }
  if (kind == STACK_KIND_RANGE) {
    const rangeIndex = rangeIndexStack[slot];
    const memberIndex = rangeMemberAt(
      rangeIndex,
      row,
      col,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (memberIndex == 0xffffffff) {
      return NaN;
    }
    return memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors);
  }
  if (kind == STACK_KIND_ARRAY) {
    const arrayIndex = rangeIndexStack[slot];
    const arrayRows = getDynamicArrayRows(arrayIndex);
    const arrayCols = getDynamicArrayCols(arrayIndex);
    if (arrayRows < 1 || arrayCols < 1 || row >= arrayRows || col >= arrayCols) {
      return NaN;
    }
    const arrayOffset = row * arrayCols + col;
    if (arrayOffset >= readSpillArrayLength(arrayIndex)) {
      return NaN;
    }
    return readSpillArrayNumber(arrayIndex, arrayOffset);
  }
  return NaN;
}

function inputCellNumeric(
  slot: i32,
  row: i32,
  col: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
): f64 {
  const kind = kindStack[slot];
  if (row < 0 || col < 0) {
    return NaN;
  }
  if (kind == STACK_KIND_SCALAR) {
    return row == 0 && col == 0 ? toNumberOrNaN(tagStack[slot], valueStack[slot]) : NaN;
  }
  if (kind == STACK_KIND_RANGE) {
    const rangeIndex = rangeIndexStack[slot];
    const memberIndex = rangeMemberAt(
      rangeIndex,
      row,
      col,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (memberIndex == 0xffffffff) {
      return NaN;
    }
    return toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
  }
  if (kind == STACK_KIND_ARRAY) {
    const arrayIndex = rangeIndexStack[slot];
    const arrayRows = getDynamicArrayRows(arrayIndex);
    const arrayCols = getDynamicArrayCols(arrayIndex);
    if (arrayRows < 1 || arrayCols < 1 || row >= arrayRows || col >= arrayCols) {
      return NaN;
    }
    const arrayOffset = row * arrayCols + col;
    if (arrayOffset >= readSpillArrayLength(arrayIndex)) {
      return NaN;
    }
    return toNumberOrNaN(
      readSpillArrayTag(arrayIndex, arrayOffset),
      readSpillArrayNumber(arrayIndex, arrayOffset),
    );
  }
  return NaN;
}

function copyInputCellToSpill(
  arrayIndex: u32,
  outputOffset: i32,
  slot: i32,
  row: i32,
  col: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  const tag = inputCellTag(
    slot,
    row,
    col,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
  );
  const value = inputCellScalarValue(
    slot,
    row,
    col,
    kindStack,
    valueStack,
    tagStack,
    rangeIndexStack,
    rangeOffsets,
    rangeLengths,
    rangeRowCounts,
    rangeColCounts,
    rangeMembers,
    cellTags,
    cellNumbers,
    cellStringIds,
    cellErrors,
  );
  if (tag == ValueTag.Error && isNaN(value)) {
    return ErrorCode.Value;
  }
  writeSpillArrayValue(arrayIndex, outputOffset, tag, value);
  return ErrorCode.None;
}

function replaceText(text: string, start: i32, count: i32, replacement: string): string {
  const startIndex = start - 1;
  if (startIndex >= text.length) {
    return text;
  }
  return text.slice(0, startIndex) + replacement + text.slice(startIndex + count);
}

function substituteText(text: string, oldText: string, newText: string): string {
  if (oldText.length == 0) {
    return text;
  }
  let result = "";
  let searchIndex = 0;
  let found = text.indexOf(oldText, searchIndex);
  if (found < 0) {
    return text;
  }
  while (found >= 0) {
    result += text.slice(searchIndex, found) + newText;
    searchIndex = found + oldText.length;
    found = text.indexOf(oldText, searchIndex);
  }
  return result + text.slice(searchIndex);
}

function substituteNthText(text: string, oldText: string, newText: string, instance: i32): string {
  let count = 0;
  let searchIndex = 0;
  while (searchIndex <= text.length) {
    const found = text.indexOf(oldText, searchIndex);
    if (found < 0) {
      return text;
    }
    count += 1;
    if (count == instance) {
      return text.slice(0, found) + newText + text.slice(found + oldText.length);
    }
    searchIndex = found + oldText.length;
  }
  return text;
}

function repeatText(text: string, count: i32): string {
  let result = "";
  for (let index = 0; index < count; index += 1) {
    result += text;
  }
  return result;
}

function valueNumber(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) {
    return value;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  if (tag != ValueTag.String) {
    return NaN;
  }
  const text = scalarText(
    tag,
    value,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  return text == null ? NaN : parseNumericText(text);
}

const CRITERIA_OP_EQ: i32 = 0;
const CRITERIA_OP_NE: i32 = 1;
const CRITERIA_OP_GT: i32 = 2;
const CRITERIA_OP_GTE: i32 = 3;
const CRITERIA_OP_LT: i32 = 4;
const CRITERIA_OP_LTE: i32 = 5;

function compareScalarValues(
  leftTag: u8,
  leftValue: f64,
  rightTag: u8,
  rightValue: f64,
  rightText: string | null,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  const leftTextlike = leftTag == ValueTag.String || leftTag == ValueTag.Empty;
  const rightTextlike = rightTag == ValueTag.String || rightTag == ValueTag.Empty;
  if (leftTextlike && rightTextlike) {
    const leftText = scalarText(
      leftTag,
      leftValue,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const resolvedRightText =
      rightText != null
        ? rightText
        : scalarText(
            rightTag,
            rightValue,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
    if (leftText == null || resolvedRightText == null) {
      return i32.MIN_VALUE;
    }
    const normalizedLeft = leftText.toUpperCase();
    const normalizedRight = resolvedRightText.toUpperCase();
    if (normalizedLeft == normalizedRight) {
      return 0;
    }
    return normalizedLeft < normalizedRight ? -1 : 1;
  }

  const leftNumeric = toNumberOrNaN(leftTag, leftValue);
  const rightNumeric = toNumberOrNaN(rightTag, rightValue);
  if (isNaN(leftNumeric) || isNaN(rightNumeric)) {
    return i32.MIN_VALUE;
  }
  if (leftNumeric == rightNumeric) {
    return 0;
  }
  return leftNumeric < rightNumeric ? -1 : 1;
}

function matchesCriteriaValue(
  valueTag: u8,
  valueValue: f64,
  criteriaTag: u8,
  criteriaValue: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): bool {
  if (valueTag == ValueTag.Error) {
    return false;
  }

  let operator = CRITERIA_OP_EQ;
  let operandTag = criteriaTag;
  let operandValue = criteriaValue;
  let operandText: string | null = null;

  if (criteriaTag == ValueTag.String) {
    const criteriaText = scalarText(
      criteriaTag,
      criteriaValue,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (criteriaText == null) {
      return false;
    }

    let rawOperand = criteriaText;
    let parsedOperator = false;
    if (criteriaText.length >= 2) {
      const prefix = criteriaText.slice(0, 2);
      if (prefix == "<=") {
        operator = CRITERIA_OP_LTE;
        rawOperand = criteriaText.slice(2);
        parsedOperator = true;
      } else if (prefix == ">=") {
        operator = CRITERIA_OP_GTE;
        rawOperand = criteriaText.slice(2);
        parsedOperator = true;
      } else if (prefix == "<>") {
        operator = CRITERIA_OP_NE;
        rawOperand = criteriaText.slice(2);
        parsedOperator = true;
      }
    }
    if (!parsedOperator && criteriaText.length >= 1) {
      const first = criteriaText.charCodeAt(0);
      if (first == 61) {
        operator = CRITERIA_OP_EQ;
        rawOperand = criteriaText.slice(1);
        parsedOperator = true;
      } else if (first == 62) {
        operator = CRITERIA_OP_GT;
        rawOperand = criteriaText.slice(1);
        parsedOperator = true;
      } else if (first == 60) {
        operator = CRITERIA_OP_LT;
        rawOperand = criteriaText.slice(1);
        parsedOperator = true;
      }
    }

    if (!parsedOperator) {
      operandText = criteriaText;
    } else {
      const trimmed = trimAsciiWhitespace(rawOperand);
      if (trimmed.length == 0) {
        operandTag = <u8>ValueTag.String;
        operandValue = 0;
        operandText = "";
      } else {
        const upper = trimmed.toUpperCase();
        if (upper == "TRUE" || upper == "FALSE") {
          operandTag = <u8>ValueTag.Boolean;
          operandValue = upper == "TRUE" ? 1 : 0;
        } else {
          const numeric = parseNumericText(trimmed);
          if (!isNaN(numeric)) {
            operandTag = <u8>ValueTag.Number;
            operandValue = numeric;
          } else {
            operandTag = <u8>ValueTag.String;
            operandValue = 0;
            operandText = trimmed;
          }
        }
      }
    }
  }

  const comparison = compareScalarValues(
    valueTag,
    valueValue,
    operandTag,
    operandValue,
    operandText,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  if (comparison == i32.MIN_VALUE) {
    return false;
  }
  if (operator == CRITERIA_OP_EQ) {
    return comparison == 0;
  }
  if (operator == CRITERIA_OP_NE) {
    return comparison != 0;
  }
  if (operator == CRITERIA_OP_GT) {
    return comparison > 0;
  }
  if (operator == CRITERIA_OP_GTE) {
    return comparison >= 0;
  }
  if (operator == CRITERIA_OP_LT) {
    return comparison < 0;
  }
  return comparison <= 0;
}

function memberScalarValue(
  memberIndex: u32,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): f64 {
  const tag = cellTags[memberIndex];
  if (tag == ValueTag.String) {
    return <f64>cellStringIds[memberIndex];
  }
  if (tag == ValueTag.Error) {
    return <f64>cellErrors[memberIndex];
  }
  return cellNumbers[memberIndex];
}

function rangeMemberAt(
  rangeIndex: u32,
  row: i32,
  col: i32,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
): u32 {
  if (rangeIndex == UNRESOLVED_WASM_OPERAND) {
    return 0xffffffff;
  }
  const rowCount = <i32>rangeRowCounts[rangeIndex];
  const colCount = <i32>rangeColCounts[rangeIndex];
  const length = <i32>rangeLengths[rangeIndex];
  if (
    rowCount <= 0 ||
    colCount <= 0 ||
    row < 0 ||
    col < 0 ||
    row >= rowCount ||
    col >= colCount ||
    row * colCount + col >= length
  ) {
    return 0xffffffff;
  }
  return rangeMembers[rangeOffsets[rangeIndex] + row * colCount + col];
}

function unresolvedRangeOperandError(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] == STACK_KIND_RANGE && rangeIndexStack[slot] == UNRESOLVED_WASM_OPERAND) {
      return ErrorCode.Ref;
    }
  }
  return -1;
}

function writeMemberResult(
  base: i32,
  memberIndex: u32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
): i32 {
  return writeResult(
    base,
    STACK_KIND_SCALAR,
    cellTags[memberIndex],
    memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
}

function writeResult(
  base: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  rangeIndexStack[base] = 0;
  valueStack[base] = value;
  tagStack[base] = tag;
  kindStack[base] = kind;
  return base + 1;
}

function writeArrayResult(
  base: i32,
  arrayIndex: u32,
  rows: i32,
  cols: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  registerDynamicArrayShape(arrayIndex, rows, cols);
  rangeIndexStack[base] = arrayIndex;
  valueStack[base] = 0;
  tagStack[base] = ValueTag.Empty;
  kindStack[base] = STACK_KIND_ARRAY;
  return base + 1;
}

function vectorSlotLength(
  slot: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
): i32 {
  if (kindStack[slot] == STACK_KIND_SCALAR) {
    return 1;
  }
  if (kindStack[slot] != STACK_KIND_RANGE) {
    return i32.MIN_VALUE;
  }
  const rangeIndex = rangeIndexStack[slot];
  const rowCount = <i32>rangeRowCounts[rangeIndex];
  const colCount = <i32>rangeColCounts[rangeIndex];
  if (rowCount <= 0 || colCount <= 0 || (rowCount != 1 && colCount != 1)) {
    return i32.MIN_VALUE;
  }
  return <i32>rangeLengths[rangeIndex];
}

function coercePositiveIntegerArg(tag: u8, value: f64, hasValue: bool, fallback: i32): i32 {
  if (!hasValue) {
    return fallback;
  }
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 1 ? truncated : i32.MIN_VALUE;
}

function coerceNumberArg(tag: u8, value: f64, hasValue: bool, fallback: f64): f64 {
  if (!hasValue) {
    return fallback;
  }
  const numeric = toNumberExact(tag, value);
  return isFinite(numeric) ? numeric : NaN;
}

function scalarErrorAt(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  tagStack: Uint8Array,
  valueStack: Float64Array,
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] == STACK_KIND_SCALAR && tagStack[slot] == ValueTag.Error) {
      return valueStack[slot];
    }
  }
  return -1;
}

function rangeErrorAt(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellErrors: Uint16Array,
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] != STACK_KIND_RANGE) {
      continue;
    }
    const rangeIndex = rangeIndexStack[slot];
    if (rangeIndex == UNRESOLVED_WASM_OPERAND) {
      return ErrorCode.Ref;
    }
    const start = rangeOffsets[rangeIndex];
    const length = <i32>rangeLengths[rangeIndex];
    for (let cursor = 0; cursor < length; cursor++) {
      const memberIndex = rangeMembers[start + cursor];
      if (cellTags[memberIndex] == ValueTag.Error) {
        return cellErrors[memberIndex];
      }
    }
  }
  return -1;
}

function rangeSupportedScalarOnly(base: i32, argc: i32, kindStack: Uint8Array): bool {
  for (let index = 0; index < argc; index++) {
    if (kindStack[base + index] == STACK_KIND_RANGE) {
      return false;
    }
  }
  return true;
}

function coerceLogical(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  if (tag == ValueTag.Error) {
    return -(<i32>value) - 1;
  }
  return -(<i32>ErrorCode.Value) - 1;
}

function scalarArgsOnly(base: i32, argc: i32, kindStack: Uint8Array): bool {
  for (let index = 0; index < argc; index++) {
    if (kindStack[base + index] != STACK_KIND_SCALAR) {
      return false;
    }
  }
  return true;
}

function statScalarValue(tag: u8, value: f64, includeStringsAsZero: bool): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) {
    return value;
  }
  if (tag == ValueTag.String) {
    return includeStringsAsZero ? 0 : NaN;
  }
  return NaN;
}

function collectScalarStatValues(
  base: i32,
  argc: i32,
  tagStack: Uint8Array,
  valueStack: Float64Array,
  includeStringsAsZero: bool,
): Array<f64> {
  const values = new Array<f64>();
  for (let index = 0; index < argc; index++) {
    const numeric = statScalarValue(
      tagStack[base + index],
      valueStack[base + index],
      includeStringsAsZero,
    );
    if (!isNaN(numeric)) {
      values.push(numeric);
    }
  }
  return values;
}

function meanOf(values: Array<f64>): f64 {
  let sum = 0.0;
  for (let index = 0; index < values.length; index++) {
    sum += unchecked(values[index]);
  }
  return values.length == 0 ? NaN : sum / <f64>values.length;
}

function sampleVarianceOf(values: Array<f64>): f64 {
  if (values.length < 2) {
    return NaN;
  }
  const mean = meanOf(values);
  let squared = 0.0;
  for (let index = 0; index < values.length; index++) {
    const deviation = unchecked(values[index]) - mean;
    squared += deviation * deviation;
  }
  return squared / <f64>(values.length - 1);
}

function populationVarianceOf(values: Array<f64>): f64 {
  if (values.length == 0) {
    return NaN;
  }
  const mean = meanOf(values);
  let squared = 0.0;
  for (let index = 0; index < values.length; index++) {
    const deviation = unchecked(values[index]) - mean;
    squared += deviation * deviation;
  }
  return squared / <f64>values.length;
}

function modeSingleOf(values: Array<f64>): f64 {
  let bestValue = 0.0;
  let bestCount = 1;
  let found = false;
  const unique = new Array<f64>();
  const counts = new Array<i32>();
  for (let index = 0; index < values.length; index++) {
    const value = unchecked(values[index]);
    let match = -1;
    for (let cursor = 0; cursor < unique.length; cursor++) {
      if (unchecked(unique[cursor]) == value) {
        match = cursor;
        break;
      }
    }
    if (match >= 0) {
      counts[match] = unchecked(counts[match]) + 1;
    } else {
      unique.push(value);
      counts.push(1);
      match = unique.length - 1;
    }
    const count = unchecked(counts[match]);
    if (count > bestCount || (count == bestCount && (!found || value < bestValue))) {
      bestCount = count;
      bestValue = value;
      found = true;
    }
  }
  return bestCount >= 2 && found ? bestValue : NaN;
}

function erfApprox(value: f64): f64 {
  const sign = value < 0 ? -1.0 : 1.0;
  const absolute = Math.abs(value);
  const t = 1.0 / (1.0 + 0.3275911 * absolute);
  const y =
    1.0 -
    ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) *
      t *
      Math.exp(-absolute * absolute);
  return sign * y;
}

function standardNormalPdf(value: f64): f64 {
  return Math.exp(-(value * value) / 2.0) / Math.sqrt(2.0 * Math.PI);
}

function standardNormalCdf(value: f64): f64 {
  return 0.5 * (1.0 + erfApprox(value / Math.sqrt(2.0)));
}

function inverseStandardNormal(value: f64): f64 {
  if (!(value > 0.0 && value < 1.0)) {
    return NaN;
  }
  const lower = 0.02425;
  const upper = 1.0 - lower;

  if (value < lower) {
    const q = Math.sqrt(-2.0 * Math.log(value));
    return (
      (((((-0.007784894002430293 * q - 0.3223964580411365) * q - 2.400758277161838) * q -
        2.549732539343734) *
        q +
        4.374664141464968) *
        q +
        2.938163982698783) /
      ((((0.007784695709041462 * q + 0.3224671290700398) * q + 2.445134137142996) * q +
        3.754408661907416) *
        q +
        1.0)
    );
  }

  if (value > upper) {
    const q = Math.sqrt(-2.0 * Math.log(1.0 - value));
    return -(
      (((((-0.007784894002430293 * q - 0.3223964580411365) * q - 2.400758277161838) * q -
        2.549732539343734) *
        q +
        4.374664141464968) *
        q +
        2.938163982698783) /
      ((((0.007784695709041462 * q + 0.3224671290700398) * q + 2.445134137142996) * q +
        3.754408661907416) *
        q +
        1.0)
    );
  }

  const q = value - 0.5;
  const r = q * q;
  return (
    ((((((-39.69683028665376 * r + 220.9460984245205) * r - 275.9285104469687) * r +
      138.357751867269) *
      r -
      30.66479806614716) *
      r +
      2.506628277459239) *
      q) /
    (((((-54.47609879822406 * r + 161.5858368580409) * r - 155.6989798598866) * r +
      66.80131188771972) *
      r -
      13.28068155288572) *
      r +
      1.0)
  );
}

function skewSampleOf(values: Array<f64>): f64 {
  if (values.length < 3) {
    return NaN;
  }
  const mean = meanOf(values);
  const stddev = Math.sqrt(sampleVarianceOf(values));
  if (!(stddev > 0.0)) {
    return NaN;
  }
  let moment3 = 0.0;
  for (let index = 0; index < values.length; index++) {
    const deviation = unchecked(values[index]) - mean;
    moment3 += deviation * deviation * deviation;
  }
  const n = <f64>values.length;
  return (n * moment3) / ((n - 1.0) * (n - 2.0) * stddev * stddev * stddev);
}

function skewPopulationOf(values: Array<f64>): f64 {
  if (values.length == 0) {
    return NaN;
  }
  const mean = meanOf(values);
  const stddev = Math.sqrt(populationVarianceOf(values));
  if (!(stddev > 0.0)) {
    return NaN;
  }
  let moment3 = 0.0;
  for (let index = 0; index < values.length; index++) {
    const deviation = unchecked(values[index]) - mean;
    moment3 += deviation * deviation * deviation;
  }
  return moment3 / <f64>values.length / (stddev * stddev * stddev);
}

function kurtosisOf(values: Array<f64>): f64 {
  if (values.length < 4) {
    return NaN;
  }
  const mean = meanOf(values);
  const stddev = Math.sqrt(sampleVarianceOf(values));
  if (!(stddev > 0.0)) {
    return NaN;
  }
  let sum4 = 0.0;
  for (let index = 0; index < values.length; index++) {
    const standardized = (unchecked(values[index]) - mean) / stddev;
    sum4 += standardized * standardized * standardized * standardized;
  }
  const n = <f64>values.length;
  return (
    (n * (n + 1.0) * sum4) / ((n - 1.0) * (n - 2.0) * (n - 3.0)) -
    (3.0 * (n - 1.0) * (n - 1.0)) / ((n - 2.0) * (n - 3.0))
  );
}

function paymentType(tag: u8, value: f64, hasValue: bool): i32 {
  if (!hasValue) {
    return 0;
  }
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return -1;
  }
  const truncated = <i32>numeric;
  return truncated == 0 || truncated == 1 ? truncated : -1;
}

function futureValueCalc(
  rate: f64,
  periods: f64,
  payment: f64,
  present: f64,
  paymentTypeValue: i32,
): f64 {
  if (rate == 0.0) {
    return -(present + payment * periods);
  }
  const growth = Math.pow(1.0 + rate, periods);
  return -(
    present * growth +
    payment * (1.0 + rate * <f64>paymentTypeValue) * ((growth - 1.0) / rate)
  );
}

function presentValueCalc(
  rate: f64,
  periods: f64,
  payment: f64,
  future: f64,
  paymentTypeValue: i32,
): f64 {
  if (rate == 0.0) {
    return -(future + payment * periods);
  }
  const growth = Math.pow(1.0 + rate, periods);
  return (
    -(future + payment * (1.0 + rate * <f64>paymentTypeValue) * ((growth - 1.0) / rate)) / growth
  );
}

function periodicPaymentCalc(
  rate: f64,
  periods: f64,
  present: f64,
  future: f64,
  paymentTypeValue: i32,
): f64 {
  if (periods <= 0.0) {
    return NaN;
  }
  if (rate == 0.0) {
    return -(future + present) / periods;
  }
  const growth = Math.pow(1.0 + rate, periods);
  const denominator = (1.0 + rate * <f64>paymentTypeValue) * (growth - 1.0);
  if (denominator == 0.0) {
    return NaN;
  }
  return (-rate * (future + present * growth)) / denominator;
}

function totalPeriodsCalc(
  rate: f64,
  payment: f64,
  present: f64,
  future: f64,
  paymentTypeValue: i32,
): f64 {
  if (payment == 0.0 && rate == 0.0) {
    return NaN;
  }
  if (rate == 0.0) {
    return payment == 0.0 ? NaN : -(future + present) / payment;
  }
  const adjustedPayment = payment * (1.0 + rate * <f64>paymentTypeValue);
  const numerator = adjustedPayment - future * rate;
  const denominator = adjustedPayment + present * rate;
  if (numerator <= 0.0 || denominator <= 0.0) {
    return NaN;
  }
  return Math.log(numerator / denominator) / Math.log(1.0 + rate);
}

function interestPaymentCalc(
  rate: f64,
  period: f64,
  periods: f64,
  present: f64,
  future: f64,
  paymentTypeValue: i32,
): f64 {
  if (period < 1.0 || period > periods) {
    return NaN;
  }
  const payment = periodicPaymentCalc(rate, periods, present, future, paymentTypeValue);
  if (isNaN(payment)) {
    return NaN;
  }
  if (paymentTypeValue == 1 && period == 1.0) {
    return 0.0;
  }
  const balance = futureValueCalc(
    rate,
    paymentTypeValue == 1 ? period - 2.0 : period - 1.0,
    payment,
    present,
    paymentTypeValue,
  );
  return balance * rate;
}

function principalPaymentCalc(
  rate: f64,
  period: f64,
  periods: f64,
  present: f64,
  future: f64,
  paymentTypeValue: i32,
): f64 {
  const payment = periodicPaymentCalc(rate, periods, present, future, paymentTypeValue);
  const interest = interestPaymentCalc(rate, period, periods, present, future, paymentTypeValue);
  return isNaN(payment) || isNaN(interest) ? NaN : payment - interest;
}

export function applyBuiltin(
  builtinId: i32,
  argc: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
  sp: i32,
): i32 {
  const base = sp - argc;
  const unresolvedRangeError = unresolvedRangeOperandError(base, argc, kindStack, rangeIndexStack);
  if (unresolvedRangeError >= 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Error,
      unresolvedRangeError,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Today) {
    if (argc != 0)
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    const nowSerial = volatileNowResult();
    if (isNaN(nowSerial))
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.floor(nowSerial),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Now) {
    if (argc != 0)
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    const nowSerial = volatileNowResult();
    if (isNaN(nowSerial))
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      nowSerial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rand) {
    if (argc != 0)
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    const next = nextVolatileRandomValue();
    if (!isFinite(next))
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    const bounded = Math.min(Math.max(next, 0), 1 - f64.EPSILON);
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      bounded,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sequence) {
    if (argc < 1 || argc > 4) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rows = coercePositiveIntegerArg(tagStack[base], valueStack[base], argc >= 1, 1);
    const cols = coercePositiveIntegerArg(tagStack[base + 1], valueStack[base + 1], argc >= 2, 1);
    const start = coerceNumberArg(tagStack[base + 2], valueStack[base + 2], argc >= 3, 1);
    const step = coerceNumberArg(tagStack[base + 3], valueStack[base + 3], argc >= 4, 1);
    if (rows == i32.MIN_VALUE || cols == i32.MIN_VALUE || isNaN(start) || isNaN(step)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const arrayIndex = allocateSpillArrayResult(rows, cols);
    const length = rows * cols;
    for (let index = 0; index < length; index++) {
      writeSpillArrayNumber(arrayIndex, index, start + <f64>index * step);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rows,
      cols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Offset && argc >= 3 && argc <= 5) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rowOffset = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    const colOffset = truncToInt(tagStack[base + 2], valueStack[base + 2]);
    const height =
      argc >= 4
        ? coercePositiveIntegerArg(tagStack[base + 3], valueStack[base + 3], true, sourceRows)
        : sourceRows;
    const width =
      argc >= 5
        ? coercePositiveIntegerArg(tagStack[base + 4], valueStack[base + 4], true, sourceCols)
        : sourceCols;
    const areaNumber = argc >= 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 1;
    if (
      rowOffset == i32.MIN_VALUE ||
      colOffset == i32.MIN_VALUE ||
      height == i32.MIN_VALUE ||
      width == i32.MIN_VALUE ||
      areaNumber == i32.MIN_VALUE ||
      areaNumber != 1 ||
      height < 1 ||
      width < 1
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rowStart = rowOffset < 0 ? sourceRows + rowOffset : rowOffset;
    const colStart = colOffset < 0 ? sourceCols + colOffset : colOffset;
    if (rowStart < 0 || colStart < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (rowStart + height > sourceRows || colStart + width > sourceCols) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (height == 1 && width == 1) {
      const result = inputCellNumeric(
        base,
        rowStart,
        colStart,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
      );
      if (isNaN(result)) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        result,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const arrayIndex = allocateSpillArrayResult(height, width);
    for (let row = 0; row < height; row++) {
      for (let col = 0; col < width; col++) {
        const sourceValue = inputCellNumeric(
          base,
          rowStart + row,
          colStart + col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(sourceValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        writeSpillArrayNumber(arrayIndex, row * width + col, sourceValue);
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      height,
      width,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Take && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const requestedRows =
      argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : sourceRows;
    const requestedCols =
      argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : sourceCols;
    if (
      (argc >= 2 && requestedRows == i32.MIN_VALUE) ||
      (argc >= 3 && requestedCols == i32.MIN_VALUE)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const clippedRows = argc >= 2 ? clipIndex(requestedRows, sourceRows) : sourceRows;
    const clippedCols = argc >= 3 ? clipIndex(requestedCols, sourceCols) : sourceCols;
    if (clippedRows == i32.MIN_VALUE || clippedCols == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rowCount =
      clippedRows > 0 ? min<i32>(clippedRows, sourceRows) : min<i32>(-clippedRows, sourceRows);
    const colCount =
      clippedCols > 0 ? min<i32>(clippedCols, sourceCols) : min<i32>(-clippedCols, sourceCols);
    if (rowCount == 0 || colCount == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rowOffset = clippedRows > 0 ? 0 : max<i32>(sourceRows - rowCount, 0);
    const colOffset = clippedCols > 0 ? 0 : max<i32>(sourceCols - colCount, 0);

    const arrayIndex = allocateSpillArrayResult(rowCount, colCount);
    let outputOffset = 0;
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < colCount; col++) {
        const sourceValue = inputCellNumeric(
          base,
          rowOffset + row,
          colOffset + col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(sourceValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rowCount,
      colCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Drop && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const requestedRows = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    const requestedCols = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (
      (argc >= 2 && requestedRows == i32.MIN_VALUE) ||
      (argc >= 3 && requestedCols == i32.MIN_VALUE)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const clippedRows =
      argc >= 2 ? (requestedRows == 0 ? 0 : clipIndex(requestedRows, sourceRows)) : 0;
    const clippedCols =
      argc >= 3 ? (requestedCols == 0 ? 0 : clipIndex(requestedCols, sourceCols)) : 0;
    if (clippedRows == i32.MIN_VALUE || clippedCols == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rowDrop =
      clippedRows >= 0 ? min<i32>(clippedRows, sourceRows) : min<i32>(-clippedRows, sourceRows);
    const colDrop =
      clippedCols >= 0 ? min<i32>(clippedCols, sourceCols) : min<i32>(-clippedCols, sourceCols);
    const rowCount = sourceRows - rowDrop;
    const colCount = sourceCols - colDrop;
    if (rowCount <= 0 || colCount <= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rowOffset = clippedRows > 0 ? rowDrop : 0;
    const colOffset = clippedCols > 0 ? colDrop : 0;

    const arrayIndex = allocateSpillArrayResult(rowCount, colCount);
    let outputOffset = 0;
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < colCount; col++) {
        const sourceValue = inputCellNumeric(
          base,
          rowOffset + row,
          colOffset + col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(sourceValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rowCount,
      colCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Choosecols && argc >= 2) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, 1, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const selectedCols = new Array<i32>();
    for (let arg = 1; arg < argc; arg++) {
      const argumentSlot = base + arg;
      if (kindStack[argumentSlot] != STACK_KIND_SCALAR) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const selectedCol = coerceInteger(tagStack[argumentSlot], valueStack[argumentSlot]) - 1;
      if (selectedCol < 0 || selectedCol >= sourceCols || selectedCol == i32.MIN_VALUE - 1) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      selectedCols.push(selectedCol);
    }
    if (selectedCols.length == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const outputCols = <i32>selectedCols.length;
    const arrayIndex = allocateSpillArrayResult(sourceRows, outputCols);
    let outputOffset = 0;
    for (let row = 0; row < sourceRows; row++) {
      for (let selectedColIndex = 0; selectedColIndex < outputCols; selectedColIndex++) {
        const selectedCol = selectedCols[selectedColIndex];
        const sourceValue = inputCellNumeric(
          base,
          row,
          selectedCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(sourceValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      sourceRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Chooserows && argc >= 2) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, 1, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const selectedRows = new Array<i32>();
    for (let arg = 1; arg < argc; arg++) {
      const argumentSlot = base + arg;
      if (kindStack[argumentSlot] != STACK_KIND_SCALAR) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const selectedRow = coerceInteger(tagStack[argumentSlot], valueStack[argumentSlot]) - 1;
      if (selectedRow < 0 || selectedRow >= sourceRows || selectedRow == i32.MIN_VALUE - 1) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      selectedRows.push(selectedRow);
    }
    if (selectedRows.length == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const outputRows = <i32>selectedRows.length;
    const arrayIndex = allocateSpillArrayResult(outputRows, sourceCols);
    let outputOffset = 0;
    for (let selectedRowIndex = 0; selectedRowIndex < outputRows; selectedRowIndex++) {
      const selectedRow = selectedRows[selectedRowIndex];
      for (let col = 0; col < sourceCols; col++) {
        const sourceValue = inputCellNumeric(
          base,
          selectedRow,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(sourceValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      sourceCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sort && argc >= 1 && argc <= 4) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sortIndex = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 1;
    const sortOrder = argc >= 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 1;
    const sortByColBoolean =
      argc >= 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 0;
    if (
      sortIndex == i32.MIN_VALUE ||
      sortIndex < 1 ||
      (sortOrder != 1 && sortOrder != -1) ||
      sortByColBoolean < 0
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sortByCol = sortByColBoolean != 0;

    if (sourceRows == 1 || sourceCols == 1) {
      const length = sourceRows * sourceCols;
      const order = new Array<i32>(length);
      for (let index = 0; index < length; index++) {
        order.push(index);
      }
      for (let index = 1; index < length; index++) {
        const current = order[index];
        const currentRow = current / sourceCols;
        const currentCol = current - currentRow * sourceCols;
        const currentValue = inputCellNumeric(
          base,
          currentRow,
          currentCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(currentValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        let cursor = index;
        while (cursor > 0) {
          const previous = order[cursor - 1];
          const previousRow = previous / sourceCols;
          const previousCol = previous - previousRow * sourceCols;
          const previousValue = inputCellNumeric(
            base,
            previousRow,
            previousCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (isNaN(previousValue)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          const comparison =
            currentValue == previousValue ? 0 : currentValue < previousValue ? -1 : 1;
          if (comparison * sortOrder < 0) {
            order[cursor] = previous;
            cursor -= 1;
            continue;
          }
          break;
        }
        order[cursor] = current;
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
      for (let index = 0; index < length; index++) {
        const sourceOffset = order[index];
        const sourceRow = sourceOffset / sourceCols;
        const sourceCol = sourceOffset - sourceRow * sourceCols;
        const sourceValue = inputCellNumeric(
          base,
          sourceRow,
          sourceCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(sourceValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        writeSpillArrayNumber(arrayIndex, index, sourceValue);
      }
      return writeArrayResult(
        base,
        arrayIndex,
        sourceRows,
        sourceCols,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (sortByCol) {
      if (sortIndex > sourceRows) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const rowSort = new Array<i32>(sourceRows);
      for (let row = 0; row < sourceRows; row++) {
        rowSort.push(row);
      }
      const sortCol = sortIndex - 1;
      for (let cursor = 1; cursor < sourceRows; cursor++) {
        const currentRow = rowSort[cursor];
        const currentValue = inputCellNumeric(
          base,
          currentRow,
          sortCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(currentValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        let position = cursor;
        while (position > 0) {
          const previousRow = rowSort[position - 1];
          const previousValue = inputCellNumeric(
            base,
            previousRow,
            sortCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (isNaN(previousValue)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          const comparison =
            currentValue == previousValue ? 0 : currentValue < previousValue ? -1 : 1;
          if (comparison * sortOrder < 0) {
            rowSort[position] = previousRow;
            position -= 1;
            continue;
          }
          break;
        }
        rowSort[position] = currentRow;
      }
      const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
      let outputOffset = 0;
      for (let sortedRow = 0; sortedRow < sourceRows; sortedRow++) {
        const sourceRow = rowSort[sortedRow];
        for (let col = 0; col < sourceCols; col++) {
          const value = inputCellNumeric(
            base,
            sourceRow,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (isNaN(value)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          writeSpillArrayNumber(arrayIndex, outputOffset, value);
          outputOffset += 1;
        }
      }
      return writeArrayResult(
        base,
        arrayIndex,
        sourceRows,
        sourceCols,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (sortIndex > sourceCols) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const colSort = new Array<i32>(sourceCols);
    for (let col = 0; col < sourceCols; col++) {
      colSort.push(col);
    }
    const sortRow = sortIndex - 1;
    for (let cursor = 1; cursor < sourceCols; cursor++) {
      const currentCol = colSort[cursor];
      const currentValue = inputCellNumeric(
        base,
        currentCol,
        sortRow,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
      );
      if (isNaN(currentValue)) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      let position = cursor;
      while (position > 0) {
        const previousCol = colSort[position - 1];
        const previousValue = inputCellNumeric(
          base,
          previousCol,
          sortRow,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(previousValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        const comparison =
          currentValue == previousValue ? 0 : currentValue < previousValue ? -1 : 1;
        if (comparison * sortOrder < 0) {
          colSort[position] = previousCol;
          position -= 1;
          continue;
        }
        break;
      }
      colSort[position] = currentCol;
    }
    const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
    let outputOffset = 0;
    for (let row = 0; row < sourceRows; row++) {
      for (let sortedCol = 0; sortedCol < sourceCols; sortedCol++) {
        const sourceCol = colSort[sortedCol];
        const value = inputCellNumeric(
          base,
          row,
          sourceCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(value)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        writeSpillArrayNumber(arrayIndex, outputOffset, value);
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      sourceRows,
      sourceCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sortby && argc >= 2) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceLength = sourceRows * sourceCols;
    if (sourceRows > 1 && sourceCols > 1) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sortBySlots = new Array<i32>();
    const sortByLengths = new Array<i32>();
    const sortByCols = new Array<i32>();
    const sortByOrders = new Array<i32>();
    let arg = 1;
    while (arg < argc) {
      const criterionSlot = base + arg;
      const criterionKind = kindStack[criterionSlot];
      if (
        criterionKind != STACK_KIND_SCALAR &&
        criterionKind != STACK_KIND_RANGE &&
        criterionKind != STACK_KIND_ARRAY
      ) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }

      const criterionRows = inputRowsFromSlot(
        criterionSlot,
        kindStack,
        rangeIndexStack,
        rangeRowCounts,
      );
      const criterionCols = inputColsFromSlot(
        criterionSlot,
        kindStack,
        rangeIndexStack,
        rangeColCounts,
      );
      if (
        criterionRows <= 0 ||
        criterionCols <= 0 ||
        criterionRows == i32.MIN_VALUE ||
        criterionCols == i32.MIN_VALUE
      ) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const criterionLength = criterionRows * criterionCols;
      if (criterionLength != 1 && criterionLength != sourceLength) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }

      sortBySlots.push(criterionSlot);
      sortByLengths.push(criterionLength);
      sortByCols.push(criterionCols);

      const nextSlot = criterionSlot + 1;
      if (nextSlot < base + argc) {
        if (kindStack[nextSlot] == STACK_KIND_RANGE || kindStack[nextSlot] == STACK_KIND_ARRAY) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        if (kindStack[nextSlot] == STACK_KIND_SCALAR) {
          const requestedOrder = coerceInteger(tagStack[nextSlot], valueStack[nextSlot]);
          if (requestedOrder != 1 && requestedOrder != -1) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          sortByOrders.push(requestedOrder);
          arg += 2;
          continue;
        }
      }

      sortByOrders.push(1);
      arg += 1;
    }

    if (sortBySlots.length == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sourceIndexes = new Array<i32>(sourceLength);
    for (let offset = 0; offset < sourceLength; offset++) {
      sourceIndexes.push(offset);
    }

    for (let cursor = 1; cursor < sourceLength; cursor++) {
      const currentOffset = sourceIndexes[cursor];
      const currentSourceRow = currentOffset / sourceCols;
      const currentSourceCol = currentOffset - currentSourceRow * sourceCols;
      const currentSourceValue = inputCellNumeric(
        base,
        currentSourceRow,
        currentSourceCol,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
      );
      if (isNaN(currentSourceValue)) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }

      let position = cursor;
      while (position > 0) {
        const previousOffset = sourceIndexes[position - 1];
        const previousSourceRow = previousOffset / sourceCols;
        const previousSourceCol = previousOffset - previousSourceRow * sourceCols;
        const previousSourceValue = inputCellNumeric(
          base,
          previousSourceRow,
          previousSourceCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        if (isNaN(previousSourceValue)) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.Value,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }

        let comparison = 0;
        for (let sortByIndex = 0; sortByIndex < sortBySlots.length; sortByIndex++) {
          const slot = sortBySlots[sortByIndex];
          const slotLength = sortByLengths[sortByIndex];
          const slotCols = sortByCols[sortByIndex];
          const slotOrder = sortByOrders[sortByIndex];

          const currentCriterionOffset = slotLength == 1 ? 0 : currentOffset;
          const previousCriterionOffset = slotLength == 1 ? 0 : previousOffset;
          const currentCriterionRow = slotLength == 1 ? 0 : currentCriterionOffset / slotCols;
          const currentCriterionCol =
            slotLength == 1 ? 0 : currentCriterionOffset - currentCriterionRow * slotCols;
          const previousCriterionRow = slotLength == 1 ? 0 : previousCriterionOffset / slotCols;
          const previousCriterionCol =
            slotLength == 1 ? 0 : previousCriterionOffset - previousCriterionRow * slotCols;

          const currentCriterionTag = inputCellTag(
            slot,
            currentCriterionRow,
            currentCriterionCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          const currentCriterionValue = inputCellNumeric(
            slot,
            currentCriterionRow,
            currentCriterionCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          const previousCriterionTag = inputCellTag(
            slot,
            previousCriterionRow,
            previousCriterionCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          const previousCriterionValue = inputCellNumeric(
            slot,
            previousCriterionRow,
            previousCriterionCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          const criterionComparison = compareScalarValues(
            currentCriterionTag,
            currentCriterionValue,
            previousCriterionTag,
            previousCriterionValue,
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (criterionComparison == i32.MIN_VALUE) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          if (criterionComparison != 0) {
            comparison = criterionComparison * slotOrder;
            break;
          }
        }
        if (comparison < 0) {
          sourceIndexes[position] = previousOffset;
          position -= 1;
          continue;
        }
        if (comparison > 0) {
          break;
        }
        break;
      }
      sourceIndexes[position] = currentOffset;
    }

    const arrayIndex = allocateSpillArrayResult(sourceRows, sourceCols);
    let outputOffset = 0;
    for (let index = 0; index < sourceLength; index++) {
      const sortedOffset = sourceIndexes[index];
      const sourceRow = sortedOffset / sourceCols;
      const sourceCol = sortedOffset - sourceRow * sourceCols;
      const sourceValue = inputCellNumeric(
        base,
        sourceRow,
        sourceCol,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
      );
      if (isNaN(sourceValue)) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
      outputOffset += 1;
    }
    return writeArrayResult(
      base,
      arrayIndex,
      sourceRows,
      sourceCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Tocol && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const ignoreValue = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    const scanByCol = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 1;
    if (ignoreValue < 0 || (ignoreValue != 0 && ignoreValue != 1) || scanByCol < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const ignoreEmpty = ignoreValue == 1;

    const values = new Array<f64>();
    if (scanByCol == 0) {
      for (let row = 0; row < sourceRows; row++) {
        for (let col = 0; col < sourceCols; col++) {
          const sourceTag = inputCellTag(
            base,
            row,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
            base,
            row,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (isNaN(sourceValue)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          values.push(sourceValue);
        }
      }
    } else {
      for (let col = 0; col < sourceCols; col++) {
        for (let row = 0; row < sourceRows; row++) {
          const sourceTag = inputCellTag(
            base,
            row,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
            base,
            row,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (isNaN(sourceValue)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          values.push(sourceValue);
        }
      }
    }

    const outputRows = values.length;
    const arrayIndex = allocateSpillArrayResult(outputRows, 1);
    for (let offset = 0; offset < outputRows; offset++) {
      writeSpillArrayNumber(arrayIndex, offset, values[offset]);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Torow && argc >= 1 && argc <= 3) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const ignoreValue = argc >= 2 ? coerceInteger(tagStack[base + 1], valueStack[base + 1]) : 0;
    const scanByCol = argc >= 3 ? coerceBoolean(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (ignoreValue < 0 || (ignoreValue != 0 && ignoreValue != 1) || scanByCol < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const ignoreEmpty = ignoreValue == 1;

    const values = new Array<f64>();
    if (scanByCol == 0) {
      for (let row = 0; row < sourceRows; row++) {
        for (let col = 0; col < sourceCols; col++) {
          const sourceTag = inputCellTag(
            base,
            row,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
            base,
            row,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (isNaN(sourceValue)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          values.push(sourceValue);
        }
      }
    } else {
      for (let col = 0; col < sourceCols; col++) {
        for (let row = 0; row < sourceRows; row++) {
          const sourceTag = inputCellTag(
            base,
            row,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (ignoreEmpty && sourceTag == ValueTag.Empty) {
            continue;
          }
          const sourceValue = inputCellNumeric(
            base,
            row,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
          );
          if (isNaN(sourceValue)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          values.push(sourceValue);
        }
      }
    }

    const outputCols = values.length;
    const arrayIndex = allocateSpillArrayResult(1, outputCols);
    for (let offset = 0; offset < outputCols; offset++) {
      writeSpillArrayNumber(arrayIndex, offset, values[offset]);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      1,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Wraprows && argc >= 2 && argc <= 4) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const wrapCount = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
    if (wrapCount == i32.MIN_VALUE || wrapCount < 1) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const needsPad = (sourceRows * sourceCols) % wrapCount != 0;
    if (needsPad && argc < 3) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (argc >= 4) {
      const padByColBoolean = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (padByColBoolean < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    const defaultPadValue = argc >= 3 ? toNumberOrNaN(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (argc >= 3 && isNaN(defaultPadValue)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sourceLength = sourceRows * sourceCols;
    const outputCols = wrapCount;
    const outputRows = (sourceLength + wrapCount - 1) / wrapCount;
    const outputLength = outputRows * outputCols;

    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
    for (let outputOffset = 0; outputOffset < sourceLength; outputOffset++) {
      const sourceRow = outputOffset / sourceCols;
      const sourceCol = outputOffset - sourceRow * sourceCols;
      const sourceValue = inputCellNumeric(
        base,
        sourceRow,
        sourceCol,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
      );
      if (isNaN(sourceValue)) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
    }
    const padValue = needsPad && argc >= 3 ? defaultPadValue : 0;
    for (let outputOffset = sourceLength; outputOffset < outputLength; outputOffset++) {
      writeSpillArrayNumber(arrayIndex, outputOffset, padValue);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Wrapcols && argc >= 2 && argc <= 4) {
    const sourceKind = kindStack[base];
    if (
      sourceKind != STACK_KIND_SCALAR &&
      sourceKind != STACK_KIND_RANGE &&
      sourceKind != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const wrapCount = coerceInteger(tagStack[base + 1], valueStack[base + 1]);
    if (wrapCount == i32.MIN_VALUE || wrapCount < 1) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (argc >= 4) {
      const padByColBoolean = coerceBoolean(tagStack[base + 3], valueStack[base + 3]);
      if (padByColBoolean < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    const needsPad = (sourceRows * sourceCols) % wrapCount != 0;
    if (needsPad && argc < 3) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const defaultPadValue = argc >= 3 ? toNumberOrNaN(tagStack[base + 2], valueStack[base + 2]) : 0;
    if (argc >= 3 && isNaN(defaultPadValue)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sourceLength = sourceRows * sourceCols;
    const outputRows = wrapCount;
    const outputCols = (sourceLength + wrapCount - 1) / wrapCount;
    const outputLength = outputRows * outputCols;

    const arrayIndex = allocateSpillArrayResult(outputRows, outputCols);
    for (let outputOffset = 0; outputOffset < outputLength; outputOffset++) {
      const sourceOffset = (outputOffset / outputRows) * outputRows + (outputOffset % outputRows);
      let sourceValue = defaultPadValue;
      if (sourceOffset < sourceLength) {
        const sourceRow = sourceOffset / sourceCols;
        const sourceCol = sourceOffset - sourceRow * sourceCols;
        sourceValue = inputCellNumeric(
          base,
          sourceRow,
          sourceCol,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
      }
      if (isNaN(sourceValue)) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      writeSpillArrayNumber(arrayIndex, outputOffset, sourceValue);
    }
    return writeArrayResult(
      base,
      arrayIndex,
      outputRows,
      outputCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Index && (argc == 2 || argc == 3)) {
    if (kindStack[base] != STACK_KIND_RANGE || kindStack[base + 1] != STACK_KIND_SCALAR) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 1],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 3 &&
      (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[base];
    const rowCount = <i32>rangeRowCounts[rangeIndex];
    const colCount = <i32>rangeColCounts[rangeIndex];
    const rawRowNum = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    const rawColNum = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 1;
    if (
      rowCount <= 0 ||
      colCount <= 0 ||
      rawRowNum == i32.MIN_VALUE ||
      rawColNum == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let rowNum = rawRowNum;
    let colNum = rawColNum;
    if (rowCount == 1 && rawColNum == 1) {
      rowNum = 1;
      colNum = rawRowNum;
    }
    if (rowNum < 1 || colNum < 1 || rowNum > rowCount || colNum > colCount) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const memberIndex = rangeMemberAt(
      rangeIndex,
      rowNum - 1,
      colNum - 1,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (memberIndex == 0xffffffff) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Ref,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeMemberResult(
      base,
      memberIndex,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    );
  }

  if (builtinId == BuiltinId.Vlookup && (argc == 3 || argc == 4)) {
    if (
      kindStack[base] != STACK_KIND_SCALAR ||
      kindStack[base + 1] != STACK_KIND_RANGE ||
      kindStack[base + 2] != STACK_KIND_SCALAR
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 4 &&
      (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[base + 1];
    const rowCount = <i32>rangeRowCounts[rangeIndex];
    const colCount = <i32>rangeColCounts[rangeIndex];
    const colIndex = truncToInt(tagStack[base + 2], valueStack[base + 2]);
    const rangeLookup = argc == 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1;
    if (
      rowCount <= 0 ||
      colCount <= 0 ||
      colIndex == i32.MIN_VALUE ||
      colIndex < 1 ||
      colIndex > colCount ||
      rangeLookup < 0
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let matchedRow = -1;
    for (let row = 0; row < rowCount; row++) {
      const memberIndex = rangeMemberAt(
        rangeIndex,
        row,
        0,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
      );
      if (memberIndex == 0xffffffff) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (comparison == i32.MIN_VALUE) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (comparison == 0) {
        matchedRow = row;
        break;
      }
      if (rangeLookup == 1 && comparison < 0) {
        matchedRow = row;
        continue;
      }
      if (rangeLookup == 1 && comparison > 0) {
        break;
      }
    }

    if (matchedRow < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.NA,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const resultMemberIndex = rangeMemberAt(
      rangeIndex,
      matchedRow,
      colIndex - 1,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (resultMemberIndex == 0xffffffff) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeMemberResult(
      base,
      resultMemberIndex,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    );
  }

  if (builtinId == BuiltinId.Hlookup && (argc == 3 || argc == 4)) {
    if (
      kindStack[base] != STACK_KIND_SCALAR ||
      kindStack[base + 1] != STACK_KIND_RANGE ||
      kindStack[base + 2] != STACK_KIND_SCALAR
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base + 2] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 4 &&
      (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[base + 1];
    const rowCount = <i32>rangeRowCounts[rangeIndex];
    const colCount = <i32>rangeColCounts[rangeIndex];
    const rowIndex = truncToInt(tagStack[base + 2], valueStack[base + 2]);
    const rangeLookup = argc == 4 ? coerceBoolean(tagStack[base + 3], valueStack[base + 3]) : 1;
    if (
      rowCount <= 0 ||
      colCount <= 0 ||
      rowIndex == i32.MIN_VALUE ||
      rowIndex < 1 ||
      rowIndex > rowCount ||
      rangeLookup < 0
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let matchedCol = -1;
    for (let col = 0; col < colCount; col++) {
      const memberIndex = rangeMemberAt(
        rangeIndex,
        0,
        col,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
      );
      if (memberIndex == 0xffffffff) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (comparison == i32.MIN_VALUE) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (comparison == 0) {
        matchedCol = col;
        break;
      }
      if (rangeLookup == 1 && comparison < 0) {
        matchedCol = col;
        continue;
      }
      if (rangeLookup == 1 && comparison > 0) {
        break;
      }
    }

    if (matchedCol < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.NA,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const resultMemberIndex = rangeMemberAt(
      rangeIndex,
      rowIndex - 1,
      matchedCol,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
    );
    if (resultMemberIndex == 0xffffffff) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeMemberResult(
      base,
      resultMemberIndex,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    );
  }

  if (builtinId == BuiltinId.Sum) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rangeError = rangeErrorAt(
      base,
      argc,
      kindStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeMembers,
      cellTags,
      cellErrors,
    );
    if (rangeError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        rangeError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let sum = 0.0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric)) {
            sum += numeric;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor++) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor);
          if (!isNaN(numeric)) {
            sum += numeric;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric)) {
        sum += numeric;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Avg) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const rangeError = rangeErrorAt(
      base,
      argc,
      kindStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeMembers,
      cellTags,
      cellErrors,
    );
    if (rangeError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        rangeError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let sum = 0.0;
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric)) {
            sum += numeric;
            count += 1;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor++) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor);
          if (!isNaN(numeric)) {
            sum += numeric;
            count += 1;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric)) {
        sum += numeric;
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count == 0 ? 0 : sum / count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Min) {
    let min = Infinity;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric) && numeric < min) {
            min = numeric;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor++) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor);
          if (!isNaN(numeric) && numeric < min) {
            min = numeric;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric) && numeric < min) {
        min = numeric;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      min,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Max) {
    let max = -Infinity;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric) && numeric > max) {
            max = numeric;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        const arrayIndex = rangeIndexStack[slot];
        const length = readSpillArrayLength(arrayIndex);
        for (let cursor = 0; cursor < length; cursor++) {
          const numeric = readSpillArrayNumber(arrayIndex, cursor);
          if (!isNaN(numeric) && numeric > max) {
            max = numeric;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric) && numeric > max) {
        max = numeric;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      max,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Count) {
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          if (!isNaN(toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]))) {
            count += 1;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        count += readSpillArrayLength(rangeIndexStack[slot]);
        continue;
      }
      if (!isNaN(toNumberOrNaN(tagStack[slot], valueStack[slot]))) {
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.CountA) {
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          if (cellTags[memberIndex] != ValueTag.Empty) {
            count += 1;
          }
        }
        continue;
      }
      if (kindStack[slot] == STACK_KIND_ARRAY) {
        count += readSpillArrayLength(rangeIndexStack[slot]);
        continue;
      }
      if (tagStack[slot] != ValueTag.Empty) {
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Countif && argc == 2) {
    if (kindStack[base] != STACK_KIND_RANGE || kindStack[base + 1] != STACK_KIND_SCALAR) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 1],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[base];
    const start = rangeOffsets[rangeIndex];
    const length = <i32>rangeLengths[rangeIndex];
    let count = 0;
    for (let cursor = 0; cursor < length; cursor++) {
      const memberIndex = rangeMembers[start + cursor];
      const memberTag = cellTags[memberIndex];
      const memberValue =
        memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
      if (
        matchesCriteriaValue(
          memberTag,
          memberValue,
          tagStack[base + 1],
          valueStack[base + 1],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
      ) {
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Countifs) {
    if (argc == 0 || argc % 2 != 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const firstRangeIndex = rangeIndexStack[base];
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const expectedLength = <i32>rangeLengths[firstRangeIndex];
    for (let index = 0; index < argc; index += 2) {
      const rangeSlot = base + index;
      const criteriaSlot = rangeSlot + 1;
      if (
        kindStack[rangeSlot] != STACK_KIND_RANGE ||
        kindStack[criteriaSlot] != STACK_KIND_SCALAR
      ) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[criteriaSlot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    let count = 0;
    for (let row = 0; row < expectedLength; row++) {
      let matchesAll = true;
      for (let index = 0; index < argc; index += 2) {
        const rangeSlot = base + index;
        const criteriaSlot = rangeSlot + 1;
        const rangeIndex = rangeIndexStack[rangeSlot];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        const memberTag = cellTags[memberIndex];
        const memberValue =
          memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
        if (
          !matchesCriteriaValue(
            memberTag,
            memberValue,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        ) {
          matchesAll = false;
          break;
        }
      }
      if (matchesAll) {
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sumif && (argc == 2 || argc == 3)) {
    const rangeSlot = base;
    const criteriaSlot = base + 1;
    const sumRangeSlot = argc == 3 ? base + 2 : base;
    if (
      kindStack[rangeSlot] != STACK_KIND_RANGE ||
      kindStack[criteriaSlot] != STACK_KIND_SCALAR ||
      kindStack[sumRangeSlot] != STACK_KIND_RANGE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[criteriaSlot] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[criteriaSlot],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[rangeSlot];
    const sumRangeIndex = rangeIndexStack[sumRangeSlot];
    const length = <i32>rangeLengths[rangeIndex];
    if (<i32>rangeLengths[sumRangeIndex] != length) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let sum = 0.0;
    for (let cursor = 0; cursor < length; cursor++) {
      const criteriaMemberIndex = rangeMembers[rangeOffsets[rangeIndex] + cursor];
      const criteriaTag = cellTags[criteriaMemberIndex];
      const criteriaValue =
        criteriaTag == ValueTag.String
          ? <f64>cellStringIds[criteriaMemberIndex]
          : cellNumbers[criteriaMemberIndex];
      if (
        !matchesCriteriaValue(
          criteriaTag,
          criteriaValue,
          tagStack[criteriaSlot],
          valueStack[criteriaSlot],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
      ) {
        continue;
      }
      const sumMemberIndex = rangeMembers[rangeOffsets[sumRangeIndex] + cursor];
      sum += toNumberOrZero(cellTags[sumMemberIndex], cellNumbers[sumMemberIndex]);
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sumifs) {
    if (argc < 3 || argc % 2 == 0 || kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const sumRangeIndex = rangeIndexStack[base];
    const expectedLength = <i32>rangeLengths[sumRangeIndex];
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index;
      const criteriaSlot = rangeSlot + 1;
      if (
        kindStack[rangeSlot] != STACK_KIND_RANGE ||
        kindStack[criteriaSlot] != STACK_KIND_SCALAR
      ) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[criteriaSlot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    let sum = 0.0;
    for (let row = 0; row < expectedLength; row++) {
      let matchesAll = true;
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index;
        const criteriaSlot = rangeSlot + 1;
        const rangeIndex = rangeIndexStack[rangeSlot];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        const memberTag = cellTags[memberIndex];
        const memberValue =
          memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
        if (
          !matchesCriteriaValue(
            memberTag,
            memberValue,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        ) {
          matchesAll = false;
          break;
        }
      }
      if (!matchesAll) {
        continue;
      }
      const sumMemberIndex = rangeMembers[rangeOffsets[sumRangeIndex] + row];
      sum += toNumberOrZero(cellTags[sumMemberIndex], cellNumbers[sumMemberIndex]);
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Averageif && (argc == 2 || argc == 3)) {
    const rangeSlot = base;
    const criteriaSlot = base + 1;
    const averageRangeSlot = argc == 3 ? base + 2 : base;
    if (
      kindStack[rangeSlot] != STACK_KIND_RANGE ||
      kindStack[criteriaSlot] != STACK_KIND_SCALAR ||
      kindStack[averageRangeSlot] != STACK_KIND_RANGE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[criteriaSlot] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[criteriaSlot],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[rangeSlot];
    const averageRangeIndex = rangeIndexStack[averageRangeSlot];
    const length = <i32>rangeLengths[rangeIndex];
    if (<i32>rangeLengths[averageRangeIndex] != length) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let count = 0;
    let sum = 0.0;
    for (let cursor = 0; cursor < length; cursor++) {
      const criteriaMemberIndex = rangeMembers[rangeOffsets[rangeIndex] + cursor];
      const criteriaTag = cellTags[criteriaMemberIndex];
      const criteriaValue =
        criteriaTag == ValueTag.String
          ? <f64>cellStringIds[criteriaMemberIndex]
          : cellNumbers[criteriaMemberIndex];
      if (
        !matchesCriteriaValue(
          criteriaTag,
          criteriaValue,
          tagStack[criteriaSlot],
          valueStack[criteriaSlot],
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        )
      ) {
        continue;
      }
      const averageMemberIndex = rangeMembers[rangeOffsets[averageRangeIndex] + cursor];
      const numeric = toNumberOrNaN(cellTags[averageMemberIndex], cellNumbers[averageMemberIndex]);
      if (isNaN(numeric)) {
        continue;
      }
      count += 1;
      sum += numeric;
    }
    if (count == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum / count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Averageifs) {
    if (argc < 3 || argc % 2 == 0 || kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const averageRangeIndex = rangeIndexStack[base];
    const expectedLength = <i32>rangeLengths[averageRangeIndex];
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index;
      const criteriaSlot = rangeSlot + 1;
      if (
        kindStack[rangeSlot] != STACK_KIND_RANGE ||
        kindStack[criteriaSlot] != STACK_KIND_SCALAR
      ) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[criteriaSlot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    let count = 0;
    let sum = 0.0;
    for (let row = 0; row < expectedLength; row++) {
      let matchesAll = true;
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index;
        const criteriaSlot = rangeSlot + 1;
        const rangeIndex = rangeIndexStack[rangeSlot];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        const memberTag = cellTags[memberIndex];
        const memberValue =
          memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
        if (
          !matchesCriteriaValue(
            memberTag,
            memberValue,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        ) {
          matchesAll = false;
          break;
        }
      }
      if (!matchesAll) {
        continue;
      }
      const averageMemberIndex = rangeMembers[rangeOffsets[averageRangeIndex] + row];
      const numeric = toNumberOrNaN(cellTags[averageMemberIndex], cellNumbers[averageMemberIndex]);
      if (isNaN(numeric)) {
        continue;
      }
      count += 1;
      sum += numeric;
    }
    if (count == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum / count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Minifs || builtinId == BuiltinId.Maxifs) {
    if (argc < 3 || argc % 2 == 0 || kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const targetRangeIndex = rangeIndexStack[base];
    const expectedLength = <i32>rangeLengths[targetRangeIndex];
    for (let index = 1; index < argc; index += 2) {
      const rangeSlot = base + index;
      const criteriaSlot = rangeSlot + 1;
      if (
        kindStack[rangeSlot] != STACK_KIND_RANGE ||
        kindStack[criteriaSlot] != STACK_KIND_SCALAR
      ) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (tagStack[criteriaSlot] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[criteriaSlot],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (<i32>rangeLengths[rangeIndexStack[rangeSlot]] != expectedLength) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    let found = false;
    let result = builtinId == BuiltinId.Minifs ? Infinity : -Infinity;
    for (let row = 0; row < expectedLength; row++) {
      let matchesAll = true;
      for (let index = 1; index < argc; index += 2) {
        const rangeSlot = base + index;
        const criteriaSlot = rangeSlot + 1;
        const rangeIndex = rangeIndexStack[rangeSlot];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        const memberTag = cellTags[memberIndex];
        const memberValue =
          memberTag == ValueTag.String ? <f64>cellStringIds[memberIndex] : cellNumbers[memberIndex];
        if (
          !matchesCriteriaValue(
            memberTag,
            memberValue,
            tagStack[criteriaSlot],
            valueStack[criteriaSlot],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        ) {
          matchesAll = false;
          break;
        }
      }
      if (!matchesAll) {
        continue;
      }
      const targetMemberIndex = rangeMembers[rangeOffsets[targetRangeIndex] + row];
      if (cellTags[targetMemberIndex] != ValueTag.Number) {
        continue;
      }
      const numeric = cellNumbers[targetMemberIndex];
      result = builtinId == BuiltinId.Minifs ? min(result, numeric) : max(result, numeric);
      found = true;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      found ? result : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sumproduct) {
    if (argc == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const firstRangeIndex = rangeIndexStack[base];
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const expectedLength = <i32>rangeLengths[firstRangeIndex];
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (
        kindStack[slot] != STACK_KIND_RANGE ||
        <i32>rangeLengths[rangeIndexStack[slot]] != expectedLength
      ) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    let sum = 0.0;
    for (let row = 0; row < expectedLength; row++) {
      let product = 1.0;
      for (let index = 0; index < argc; index++) {
        const rangeIndex = rangeIndexStack[base + index];
        const memberIndex = rangeMembers[rangeOffsets[rangeIndex] + row];
        product *= toNumberOrZero(cellTags[memberIndex], cellNumbers[memberIndex]);
      }
      sum += product;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Match && (argc == 2 || argc == 3)) {
    if (kindStack[base] != STACK_KIND_SCALAR || kindStack[base + 1] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 3 &&
      (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const matchType = argc == 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 1;
    if (!(matchType == -1 || matchType == 0 || matchType == 1)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[base + 1];
    const start = rangeOffsets[rangeIndex];
    const length = <i32>rangeLengths[rangeIndex];
    let best = -1;
    for (let index = 0; index < length; index++) {
      const memberIndex = rangeMembers[start + index];
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (matchType == 0) {
        if (comparison == 0) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Number,
            index + 1,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        continue;
      }
      if (comparison == i32.MIN_VALUE) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.NA,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (matchType == 1) {
        if (comparison <= 0) {
          best = index + 1;
        } else {
          break;
        }
      } else if (comparison >= 0) {
        best = index + 1;
      } else {
        break;
      }
    }
    return best < 0
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.NA,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          best,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Xmatch && argc >= 2 && argc <= 4) {
    if (kindStack[base] != STACK_KIND_SCALAR || kindStack[base + 1] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc >= 3 &&
      (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc >= 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 4 &&
      (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const matchMode = argc >= 3 ? truncToInt(tagStack[base + 2], valueStack[base + 2]) : 0;
    const searchMode = argc == 4 ? truncToInt(tagStack[base + 3], valueStack[base + 3]) : 1;
    if (
      !(matchMode == -1 || matchMode == 0 || matchMode == 1) ||
      !(searchMode == -1 || searchMode == 1)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const rangeIndex = rangeIndexStack[base + 1];
    const start = rangeOffsets[rangeIndex];
    const length = <i32>rangeLengths[rangeIndex];
    if (searchMode == -1) {
      if (matchMode == 0) {
        for (let index = length - 1; index >= 0; index--) {
          const memberIndex = rangeMembers[start + index];
          const comparison = compareScalarValues(
            cellTags[memberIndex],
            memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
            tagStack[base],
            valueStack[base],
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (comparison == 0) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Number,
              index + 1,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
        }
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.NA,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }

      let bestReversed = -1;
      let reversedPosition = 0;
      for (let index = length - 1; index >= 0; index--) {
        reversedPosition += 1;
        const memberIndex = rangeMembers[start + index];
        const comparison = compareScalarValues(
          cellTags[memberIndex],
          memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == i32.MIN_VALUE) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.NA,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        if (matchMode == 1) {
          if (comparison <= 0) {
            bestReversed = reversedPosition;
          } else {
            break;
          }
        } else if (comparison >= 0) {
          bestReversed = reversedPosition;
        } else {
          break;
        }
      }
      return bestReversed < 0
        ? writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            ErrorCode.NA,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          )
        : writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Number,
            length - bestReversed + 1,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
    }

    if (matchMode == 0) {
      for (let index = 0; index < length; index++) {
        const memberIndex = rangeMembers[start + index];
        const comparison = compareScalarValues(
          cellTags[memberIndex],
          memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == 0) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Number,
            index + 1,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
      }
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.NA,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let best = -1;
    for (let index = 0; index < length; index++) {
      const memberIndex = rangeMembers[start + index];
      const comparison = compareScalarValues(
        cellTags[memberIndex],
        memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (comparison == i32.MIN_VALUE) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.NA,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (matchMode == 1) {
        if (comparison <= 0) {
          best = index + 1;
        } else {
          break;
        }
      } else if (comparison >= 0) {
        best = index + 1;
      } else {
        break;
      }
    }
    return best < 0
      ? writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.NA,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      : writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          best,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
  }

  if (builtinId == BuiltinId.Xlookup && argc >= 3 && argc <= 6) {
    if (
      kindStack[base] != STACK_KIND_SCALAR ||
      kindStack[base + 1] != STACK_KIND_RANGE ||
      kindStack[base + 2] != STACK_KIND_RANGE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc >= 4 &&
      (kindStack[base + 3] != STACK_KIND_SCALAR || tagStack[base + 3] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc >= 4 && tagStack[base + 3] == ValueTag.Error ? valueStack[base + 3] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc >= 5 &&
      (kindStack[base + 4] != STACK_KIND_SCALAR || tagStack[base + 4] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc >= 5 && tagStack[base + 4] == ValueTag.Error ? valueStack[base + 4] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 6 &&
      (kindStack[base + 5] != STACK_KIND_SCALAR || tagStack[base + 5] == ValueTag.Error)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 6 && tagStack[base + 5] == ValueTag.Error ? valueStack[base + 5] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const lookupRangeIndex = rangeIndexStack[base + 1];
    const returnRangeIndex = rangeIndexStack[base + 2];
    const length = <i32>rangeLengths[lookupRangeIndex];
    if (<i32>rangeLengths[returnRangeIndex] != length) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const matchMode = argc >= 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0;
    const searchMode = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 1;
    if (matchMode != 0 || !(searchMode == -1 || searchMode == 1)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const lookupStart = rangeOffsets[lookupRangeIndex];
    const returnStart = rangeOffsets[returnRangeIndex];
    if (searchMode == -1) {
      for (let index = length - 1; index >= 0; index--) {
        const lookupMemberIndex = rangeMembers[lookupStart + index];
        const comparison = compareScalarValues(
          cellTags[lookupMemberIndex],
          memberScalarValue(lookupMemberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == 0) {
          const returnMemberIndex = rangeMembers[returnStart + index];
          return writeMemberResult(
            base,
            returnMemberIndex,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          );
        }
      }
    } else {
      for (let index = 0; index < length; index++) {
        const lookupMemberIndex = rangeMembers[lookupStart + index];
        const comparison = compareScalarValues(
          cellTags[lookupMemberIndex],
          memberScalarValue(lookupMemberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == 0) {
          const returnMemberIndex = rangeMembers[returnStart + index];
          return writeMemberResult(
            base,
            returnMemberIndex,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          );
        }
      }
    }

    if (argc >= 4) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        tagStack[base + 3],
        valueStack[base + 3],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Error,
      ErrorCode.NA,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Lookup && (argc == 2 || argc == 3)) {
    if (kindStack[base] != STACK_KIND_SCALAR) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      (kindStack[base + 1] != STACK_KIND_SCALAR && kindStack[base + 1] != STACK_KIND_RANGE) ||
      (argc == 3 &&
        kindStack[base + 2] != STACK_KIND_SCALAR &&
        kindStack[base + 2] != STACK_KIND_RANGE)
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (kindStack[base + 1] == STACK_KIND_SCALAR && tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 1],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (
      argc == 3 &&
      kindStack[base + 2] == STACK_KIND_SCALAR &&
      tagStack[base + 2] == ValueTag.Error
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 2],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const lookupLength = vectorSlotLength(
      base + 1,
      kindStack,
      rangeIndexStack,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
    );
    const resultSlot = argc == 3 ? base + 2 : base + 1;
    const resultLength = vectorSlotLength(
      resultSlot,
      kindStack,
      rangeIndexStack,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
    );
    if (
      lookupLength == i32.MIN_VALUE ||
      resultLength == i32.MIN_VALUE ||
      lookupLength != resultLength
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    let position = -1;
    if (kindStack[base + 1] == STACK_KIND_SCALAR) {
      const exactComparison = compareScalarValues(
        tagStack[base + 1],
        valueStack[base + 1],
        tagStack[base],
        valueStack[base],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (exactComparison == 0) {
        position = 1;
      } else if (
        position < 0 &&
        tagStack[base] == ValueTag.Number &&
        exactComparison != i32.MIN_VALUE &&
        exactComparison <= 0
      ) {
        position = 1;
      }
    } else {
      const lookupRangeIndex = rangeIndexStack[base + 1];
      const lookupStart = rangeOffsets[lookupRangeIndex];
      for (let index = 0; index < lookupLength; index++) {
        const memberIndex = rangeMembers[lookupStart + index];
        const comparison = compareScalarValues(
          cellTags[memberIndex],
          memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
          tagStack[base],
          valueStack[base],
          null,
          stringOffsets,
          stringLengths,
          stringData,
          outputStringOffsets,
          outputStringLengths,
          outputStringData,
        );
        if (comparison == 0) {
          position = index + 1;
          break;
        }
      }
      if (position < 0 && tagStack[base] == ValueTag.Number) {
        let best = -1;
        for (let index = 0; index < lookupLength; index++) {
          const memberIndex = rangeMembers[lookupStart + index];
          const comparison = compareScalarValues(
            cellTags[memberIndex],
            memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
            tagStack[base],
            valueStack[base],
            null,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (comparison == i32.MIN_VALUE) {
            best = -1;
            break;
          }
          if (comparison <= 0) {
            best = index + 1;
            continue;
          }
          break;
        }
        position = best;
      }
    }

    if (position < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.NA,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    if (kindStack[resultSlot] == STACK_KIND_SCALAR) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        tagStack[resultSlot],
        valueStack[resultSlot],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const resultRangeIndex = rangeIndexStack[resultSlot];
    const resultMemberIndex = rangeMembers[rangeOffsets[resultRangeIndex] + position - 1];
    return writeMemberResult(
      base,
      resultMemberIndex,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    );
  }

  if (builtinId == BuiltinId.Areas && argc == 1) {
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Arraytotext && (argc == 1 || argc == 2)) {
    if (kindStack[base] != STACK_KIND_SCALAR && kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    let format = 0;
    if (argc == 2) {
      if (kindStack[base + 1] != STACK_KIND_SCALAR) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      format = truncToInt(tagStack[base + 1], valueStack[base + 1]);
      if (!(format == 0 || format == 1)) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    const strict = format == 1;
    let text = strict ? "{" : "";
    if (kindStack[base] == STACK_KIND_SCALAR) {
      const cellText = arrayToTextCell(
        tagStack[base],
        valueStack[base],
        strict,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (cellText == null) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      text += cellText;
    } else {
      const rangeIndex = rangeIndexStack[base];
      const rowCount = <i32>rangeRowCounts[rangeIndex];
      const colCount = <i32>rangeColCounts[rangeIndex];
      for (let row = 0; row < rowCount; row++) {
        if (row > 0) {
          text += ";";
        }
        for (let col = 0; col < colCount; col++) {
          if (col > 0) {
            text += strict ? ", " : "\t";
          }
          const memberIndex = rangeMemberAt(
            rangeIndex,
            row,
            col,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
          );
          if (memberIndex == 0xffffffff) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          const cellText = arrayToTextCell(
            cellTags[memberIndex],
            memberScalarValue(memberIndex, cellTags, cellNumbers, cellStringIds, cellErrors),
            strict,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
          if (cellText == null) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          text += cellText;
        }
      }
    }
    if (strict) {
      text += "}";
    }
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Columns && argc == 1) {
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      rangeColCounts[rangeIndexStack[base]],
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rows && argc == 1) {
    if (kindStack[base] != STACK_KIND_RANGE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      rangeRowCounts[rangeIndexStack[base]],
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Transpose && argc == 1) {
    if (
      kindStack[base] != STACK_KIND_SCALAR &&
      kindStack[base] != STACK_KIND_RANGE &&
      kindStack[base] != STACK_KIND_ARRAY
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const sourceRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts);
    const sourceCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts);
    if (
      sourceRows <= 0 ||
      sourceCols <= 0 ||
      sourceRows == i32.MIN_VALUE ||
      sourceCols == i32.MIN_VALUE
    ) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (sourceRows == 1 && sourceCols == 1) {
      const tag = inputCellTag(
        base,
        0,
        0,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
      );
      const value = inputCellScalarValue(
        base,
        0,
        0,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (tag == ValueTag.Error && isNaN(value)) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        tag,
        value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const arrayIndex = allocateSpillArrayResult(sourceCols, sourceRows);
    let outputOffset = 0;
    for (let col = 0; col < sourceCols; col++) {
      for (let row = 0; row < sourceRows; row++) {
        const copyError = copyInputCellToSpill(
          arrayIndex,
          outputOffset,
          base,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        );
        if (copyError != ErrorCode.None) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            copyError,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          );
        }
        outputOffset += 1;
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      sourceCols,
      sourceRows,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Hstack && argc >= 1) {
    let rowCount = 0;
    let totalCols = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      const kind = kindStack[slot];
      if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE && kind != STACK_KIND_ARRAY) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
      if (rows <= 0 || cols <= 0 || rows == i32.MIN_VALUE || cols == i32.MIN_VALUE) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      rowCount = max(rowCount, rows);
      totalCols += cols;
    }
    for (let index = 0; index < argc; index++) {
      const rows = inputRowsFromSlot(base + index, kindStack, rangeIndexStack, rangeRowCounts);
      if (rows != 1 && rows != rowCount) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    const arrayIndex = allocateSpillArrayResult(rowCount, totalCols);
    let outputOffset = 0;
    for (let row = 0; row < rowCount; row++) {
      for (let index = 0; index < argc; index++) {
        const slot = base + index;
        const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
        const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
        const sourceRow = rows == 1 ? 0 : row;
        for (let col = 0; col < cols; col++) {
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            slot,
            sourceRow,
            col,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          );
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      rowCount,
      totalCols,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Vstack && argc >= 1) {
    let totalRows = 0;
    let colCount = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      const kind = kindStack[slot];
      if (kind != STACK_KIND_SCALAR && kind != STACK_KIND_RANGE && kind != STACK_KIND_ARRAY) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
      if (rows <= 0 || cols <= 0 || rows == i32.MIN_VALUE || cols == i32.MIN_VALUE) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      totalRows += rows;
      colCount = max(colCount, cols);
    }
    for (let index = 0; index < argc; index++) {
      const cols = inputColsFromSlot(base + index, kindStack, rangeIndexStack, rangeColCounts);
      if (cols != 1 && cols != colCount) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    const arrayIndex = allocateSpillArrayResult(totalRows, colCount);
    let outputOffset = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
      const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
      for (let row = 0; row < rows; row++) {
        for (let col = 0; col < colCount; col++) {
          const sourceCol = cols == 1 ? 0 : col;
          const copyError = copyInputCellToSpill(
            arrayIndex,
            outputOffset,
            slot,
            row,
            sourceCol,
            kindStack,
            valueStack,
            tagStack,
            rangeIndexStack,
            rangeOffsets,
            rangeLengths,
            rangeRowCounts,
            rangeColCounts,
            rangeMembers,
            cellTags,
            cellNumbers,
            cellStringIds,
            cellErrors,
          );
          if (copyError != ErrorCode.None) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              copyError,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            );
          }
          outputOffset += 1;
        }
      }
    }
    return writeArrayResult(
      base,
      arrayIndex,
      totalRows,
      colCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Error,
      ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Na && argc == 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Error,
      ErrorCode.NA,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Iferror && argc == 2) {
    if (tagStack[base] == ValueTag.Error) {
      kindStack[base] = kindStack[base + 1];
      tagStack[base] = tagStack[base + 1];
      valueStack[base] = valueStack[base + 1];
      rangeIndexStack[base] = rangeIndexStack[base + 1];
    }
    return base + 1;
  }

  if (builtinId == BuiltinId.Ifna && argc == 2) {
    if (tagStack[base] == ValueTag.Error && <i32>valueStack[base] == ErrorCode.NA) {
      kindStack[base] = kindStack[base + 1];
      tagStack[base] = tagStack[base + 1];
      valueStack[base] = valueStack[base + 1];
      rangeIndexStack[base] = rangeIndexStack[base + 1];
    }
    return base + 1;
  }

  if (builtinId == BuiltinId.T && (argc == 0 || argc == 1)) {
    if (argc == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Empty,
        0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (!scalarArgsOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base] == ValueTag.Error) {
      return base + 1;
    }
    if (tagStack[base] == ValueTag.String) {
      return base + 1;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Empty,
      0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.N && (argc == 0 || argc == 1)) {
    if (argc == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (!scalarArgsOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (tagStack[base] == ValueTag.Error) {
      return base + 1;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      toNumberOrZero(tagStack[base], valueStack[base]),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Type && (argc == 0 || argc == 1)) {
    if (argc == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        1,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const typeCode =
      kindStack[base] == STACK_KIND_ARRAY || kindStack[base] == STACK_KIND_RANGE
        ? 64
        : tagStack[base] == ValueTag.Number || tagStack[base] == ValueTag.Empty
          ? 1
          : tagStack[base] == ValueTag.String
            ? 2
            : tagStack[base] == ValueTag.Boolean
              ? 4
              : 16;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>typeCode,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Delta && (argc == 1 || argc == 2)) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const left = valueNumber(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const right =
      argc == 2
        ? valueNumber(
            tagStack[base + 1],
            valueStack[base + 1],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0;
    if (isNaN(left) || isNaN(right)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      left == right ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Gestep && (argc == 1 || argc == 2)) {
    if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const numberValue = valueNumber(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const stepValue =
      argc == 2
        ? valueNumber(
            tagStack[base + 1],
            valueStack[base + 1],
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          )
        : 0.0;
    if (isNaN(numberValue) || isNaN(stepValue)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      numberValue >= stepValue ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
  if (scalarError >= 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Error,
      scalarError,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Abs && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.abs(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Round && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    if (isNaN(numeric)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      roundToDigits(numeric, 0),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Round && argc == 2) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const digits = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (isNaN(numeric) || isNaN(digits)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      roundToDigits(numeric, <i32>digits),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Floor && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    if (isNaN(numeric)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.floor(numeric),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Floor && argc == 2) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const significance = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (isNaN(numeric) || isNaN(significance)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (significance == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.floor(numeric / significance) * significance,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Ceiling && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    if (isNaN(numeric)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.ceil(numeric),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Ceiling && argc == 2) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const significance = toNumberExact(tagStack[base + 1], valueStack[base + 1]);
    if (isNaN(numeric) || isNaN(significance)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (significance == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.ceil(numeric / significance) * significance,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Mod && argc == 2) {
    const divisor = toNumberOrZero(tagStack[base + 1], valueStack[base + 1]);
    if (divisor == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Div0,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      toNumberOrZero(tagStack[base], valueStack[base]) % divisor,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.And) {
    if (argc == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    for (let index = 0; index < argc; index++) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index]);
      if (coerced < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          -coerced - 1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (coerced == 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Boolean,
          0,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Or) {
    if (argc == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    for (let index = 0; index < argc; index++) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index]);
      if (coerced < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          -coerced - 1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (coerced != 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Boolean,
          1,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Not && argc == 1) {
    const coerced = coerceLogical(tagStack[base], valueStack[base]);
    if (coerced < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        -coerced - 1,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      coerced == 0 ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Concat) {
    let scalarError = -1;
    for (let index = 0; index < argc; index++) {
      if (tagStack[base + index] == ValueTag.Error) {
        scalarError = <i32>valueStack[base + index];
        break;
      }
    }
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    for (let index = 0; index < argc; index++) {
      const len = textLength(
        tagStack[base + index],
        valueStack[base + index],
        stringLengths,
        outputStringLengths,
      );
      if (len < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
    }

    let text = "";
    for (let index = 0; index < argc; index++) {
      const part = scalarText(
        tagStack[base + index],
        valueStack[base + index],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      );
      if (part != null) {
        text += part;
      }
    }
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Len && argc == 1) {
    const length = textLength(tagStack[base], valueStack[base], stringLengths, outputStringLengths);
    if (length < 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>length,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Exact && argc == 2) {
    const left = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const right = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (left === null || right === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      left == right ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Left || builtinId == BuiltinId.Right) && (argc == 1 || argc == 2)) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const count = argc == 2 ? coerceLength(tagStack[base + 1], valueStack[base + 1], 1) : 1;
    if (text == null || count == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const result =
      builtinId == BuiltinId.Left
        ? text.slice(0, count)
        : count == 0
          ? ""
          : count >= text.length
            ? text
            : text.slice(text.length - count);
    return writeStringResult(base, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Mid && argc == 3) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1);
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0);
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      text.slice(start - 1, start - 1 + count),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Trim && argc == 1) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      excelTrim(text),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if ((builtinId == BuiltinId.Upper || builtinId == BuiltinId.Lower) && argc == 1) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      builtinId == BuiltinId.Upper ? text.toUpperCase() : text.toLowerCase(),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Find && (argc == 2 || argc == 3)) {
    const needle = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const haystack = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const start = argc == 3 ? coercePositiveStart(tagStack[base + 2], valueStack[base + 2], 1) : 1;
    if (needle == null || haystack == null || start == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const found = findPosition(needle, haystack, start, true, false);
    if (found == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>found,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Search && (argc == 2 || argc == 3)) {
    const needle = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const haystack = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const start = argc == 3 ? coercePositiveStart(tagStack[base + 2], valueStack[base + 2], 1) : 1;
    if (needle == null || haystack == null || start == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const found = findPosition(needle, haystack, start, false, true);
    if (found == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>found,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Value && argc == 1) {
    const numeric = valueNumber(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (isNaN(numeric)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      numeric,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsBlank && argc == 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      1,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsBlank && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      tagStack[base] == ValueTag.Empty ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsNumber && argc == 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsNumber && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      tagStack[base] == ValueTag.Number ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsText && argc == 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.IsText && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      tagStack[base] == ValueTag.String ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Date && argc == 3) {
    const serial = excelDateSerial(
      tagStack[base],
      valueStack[base],
      tagStack[base + 1],
      valueStack[base + 1],
      tagStack[base + 2],
      valueStack[base + 2],
    );
    if (isNaN(serial)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      serial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Year && argc == 1) {
    const year = excelYearPartFromSerial(tagStack[base], valueStack[base]);
    if (year == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>year,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Month && argc == 1) {
    const month = excelMonthPartFromSerial(tagStack[base], valueStack[base]);
    if (month == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>month,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Day && argc == 1) {
    const day = excelDayPartFromSerial(tagStack[base], valueStack[base]);
    if (day == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>day,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Time && argc == 3) {
    const serial = excelTimeSerial(
      tagStack[base],
      valueStack[base],
      tagStack[base + 1],
      valueStack[base + 1],
      tagStack[base + 2],
      valueStack[base + 2],
    );
    if (isNaN(serial)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      serial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Hour && argc == 1) {
    const second = excelSecondOfDay(tagStack[base], valueStack[base]);
    if (second == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>(second / 3600),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Minute && argc == 1) {
    const second = excelSecondOfDay(tagStack[base], valueStack[base]);
    if (second == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>((second % 3600) / 60),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Second && argc == 1) {
    const second = excelSecondOfDay(tagStack[base], valueStack[base]);
    if (second == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>(second % 60),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Weekday && (argc == 1 || argc == 2)) {
    const returnType = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 1;
    const weekday =
      returnType == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : excelWeekdayFromSerial(tagStack[base], valueStack[base], returnType);
    if (weekday == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>weekday,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Edate && argc == 2) {
    const serial = addMonthsExcelSerial(
      tagStack[base],
      valueStack[base],
      tagStack[base + 1],
      valueStack[base + 1],
      false,
    );
    if (isNaN(serial)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      serial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Eomonth && argc == 2) {
    const serial = addMonthsExcelSerial(
      tagStack[base],
      valueStack[base],
      tagStack[base + 1],
      valueStack[base + 1],
      true,
    );
    if (isNaN(serial)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      serial,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Int && argc == 1) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    if (isNaN(numeric)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.floor(numeric),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (
    (builtinId == BuiltinId.RoundUp || builtinId == BuiltinId.RoundDown) &&
    (argc == 1 || argc == 2)
  ) {
    const numeric = toNumberExact(tagStack[base], valueStack[base]);
    const digits = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 0;
    if (isNaN(numeric) || digits == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    let result = 0.0;
    if (digits >= 0) {
      const factor = Math.pow(10.0, <f64>digits);
      const scaled = numeric * factor;
      result =
        (builtinId == BuiltinId.RoundUp
          ? numeric >= 0
            ? Math.ceil(scaled)
            : Math.floor(scaled)
          : numeric >= 0
            ? Math.floor(scaled)
            : Math.ceil(scaled)) / factor;
    } else {
      const factor = Math.pow(10.0, <f64>-digits);
      const scaled = numeric / factor;
      result =
        (builtinId == BuiltinId.RoundUp
          ? numeric >= 0
            ? Math.ceil(scaled)
            : Math.floor(scaled)
          : numeric >= 0
            ? Math.floor(scaled)
            : Math.ceil(scaled)) * factor;
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      result,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Sin && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.sin(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Cos && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.cos(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Tan && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.tan(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Asin && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.asin(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Acos && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.acos(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Atan && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.atan(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Atan2 && argc == 2) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.atan2(
        toNumberOrZero(tagStack[base], valueStack[base]),
        toNumberOrZero(tagStack[base + 1], valueStack[base + 1]),
      ),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Degrees && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      (toNumberOrZero(tagStack[base], valueStack[base]) * 180.0) / Math.PI,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Radians && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      (toNumberOrZero(tagStack[base], valueStack[base]) * Math.PI) / 180.0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Exp && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.exp(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Ln && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.log(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Log10 && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.log10(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Log && (argc == 1 || argc == 2)) {
    const num = toNumberOrZero(tagStack[base], valueStack[base]);
    const baseVal = argc == 2 ? toNumberOrZero(tagStack[base + 1], valueStack[base + 1]) : 10.0;
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.log(num) / Math.log(baseVal),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Power && argc == 2) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.pow(
        toNumberOrZero(tagStack[base], valueStack[base]),
        toNumberOrZero(tagStack[base + 1], valueStack[base + 1]),
      ),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Sqrt && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.sqrt(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }
  if (builtinId == BuiltinId.Pi && argc == 0) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.PI,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Days && argc == 2) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const end = excelSerialWhole(tagStack[base], valueStack[base]);
    const start = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    if (end == i32.MIN_VALUE || start == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>(end - start),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Weeknum && (argc == 1 || argc == 2)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const returnType = argc == 2 ? truncToInt(tagStack[base + 1], valueStack[base + 1]) : 1;
    const weeknum =
      returnType == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : excelWeeknumFromSerial(tagStack[base], valueStack[base], returnType);
    if (weeknum == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>weeknum,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Workday && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, min<i32>(argc, 2), kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const offset = truncToInt(tagStack[base + 1], valueStack[base + 1]);
    if (start == i32.MIN_VALUE || offset == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const holidayKind = argc == 3 ? kindStack[base + 2] : STACK_KIND_SCALAR;
    const holidayTag = argc == 3 ? tagStack[base + 2] : <u8>ValueTag.Empty;
    const holidayValue = argc == 3 ? valueStack[base + 2] : 0.0;
    const holidayRangeIndex = argc == 3 ? rangeIndexStack[base + 2] : 0;

    let cursor = start;
    const direction = offset >= 0 ? 1 : -1;
    while (true) {
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (workday == 1) {
        break;
      }
      cursor += direction;
    }

    let remaining = offset >= 0 ? offset : -offset;
    while (remaining > 0) {
      cursor += direction;
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (workday == 1) {
        remaining -= 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>cursor,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Networkdays && (argc == 2 || argc == 3)) {
    const scalarError = scalarErrorAt(base, min<i32>(argc, 2), kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const start = excelSerialWhole(tagStack[base], valueStack[base]);
    const end = excelSerialWhole(tagStack[base + 1], valueStack[base + 1]);
    if (start == i32.MIN_VALUE || end == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }

    const holidayKind = argc == 3 ? kindStack[base + 2] : STACK_KIND_SCALAR;
    const holidayTag = argc == 3 ? tagStack[base + 2] : <u8>ValueTag.Empty;
    const holidayValue = argc == 3 ? valueStack[base + 2] : 0.0;
    const holidayRangeIndex = argc == 3 ? rangeIndexStack[base + 2] : 0;

    const step = start <= end ? 1 : -1;
    let count = 0;
    for (let cursor = start; ; cursor += step) {
      const workday = isWorkdaySerial(
        cursor,
        holidayKind,
        holidayTag,
        holidayValue,
        holidayRangeIndex,
        rangeOffsets,
        rangeLengths,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      );
      if (workday < 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          ErrorCode.Value,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        );
      }
      if (workday == 1) {
        count += step;
      }
      if (cursor == end) {
        break;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Replace && argc == 4) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1);
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0);
    const replacement = scalarText(
      tagStack[base + 3],
      valueStack[base + 3],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE || replacement == null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      replaceText(text, start, count, replacement),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Substitute && (argc == 3 || argc == 4)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const oldText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const newText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || oldText == null || newText == null || oldText.length == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (argc == 3) {
      return writeStringResult(
        base,
        substituteText(text, oldText, newText),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const instance = coercePositiveStart(tagStack[base + 3], valueStack[base + 3], 1);
    if (instance == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      substituteNthText(text, oldText, newText, instance),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rept && argc == 2) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const count = coerceNonNegativeLength(tagStack[base + 1], valueStack[base + 1]);
    if (text == null || count == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      repeatText(text, count),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  return writeResult(
    base,
    STACK_KIND_SCALAR,
    <u8>ValueTag.Error,
    ErrorCode.Value,
    rangeIndexStack,
    valueStack,
    tagStack,
    kindStack,
  );
}
