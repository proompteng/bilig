import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import { createBlockedBuiltinMap, datetimePlaceholderBuiltinNames } from "./placeholder.js";

export type Builtin = (...args: CellValue[]) => CellValue;
export type DateTimeProvider = () => Date;
export type RandomProvider = () => number;

export interface ExcelDateParts {
  year: number;
  month: number;
  day: number;
}

const MS_PER_DAY = 86_400_000;
const SECONDS_PER_DAY = 86_400;
const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 31);
const EXCEL_LEAP_BUG_CUTOFF_UTC_MS = Date.UTC(1900, 2, 1);

function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value };
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function firstError(args: readonly CellValue[]): CellValue | undefined {
  return args.find((arg) => arg.tag === ValueTag.Error);
}

function coerceNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return Number.isFinite(value.value) ? value.value : undefined;
    case ValueTag.Boolean:
      return value.value ? 1 : 0;
    case ValueTag.Empty:
      return 0;
    case ValueTag.String:
    case ValueTag.Error:
      return undefined;
    default:
      return undefined;
  }
}

function truncArg(value: CellValue): number | CellValue {
  if (value.tag === ValueTag.Error) {
    return value;
  }
  const coerced = coerceNumber(value);
  if (coerced === undefined) {
    return valueError();
  }
  return Math.trunc(coerced);
}

function floorDateSerial(serial: number): number {
  return Math.floor(serial);
}

function isErrorValue(value: Set<number> | CellValue): value is CellValue {
  return !("size" in value);
}

function normalizeSecondOfDay(serial: number): number | undefined {
  if (!Number.isFinite(serial) || serial < 0) {
    return undefined;
  }
  const fraction = serial - floorDateSerial(serial);
  const normalizedFraction = fraction < 0 ? fraction + 1 : fraction;
  return Math.floor(normalizedFraction * SECONDS_PER_DAY + 1e-9) % SECONDS_PER_DAY;
}

function isExcelLeapBugDate(parts: ExcelDateParts): boolean {
  return parts.year === 1900 && parts.month === 2 && parts.day === 29;
}

function daysInExcelMonth(year: number, month: number): number {
  if (year === 1900 && month === 2) {
    return 29;
  }
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

function normalizeMonth(year: number, month: number): { year: number; month: number } {
  const zeroBased = year * 12 + (month - 1);
  const normalizedYear = Math.floor(zeroBased / 12);
  const normalizedMonth = zeroBased - normalizedYear * 12 + 1;
  return { year: normalizedYear, month: normalizedMonth };
}

export function excelSerialToDateParts(serial: number): ExcelDateParts | undefined {
  if (!Number.isFinite(serial)) {
    return undefined;
  }

  const whole = floorDateSerial(serial);
  if (whole === 60) {
    return { year: 1900, month: 2, day: 29 };
  }

  const adjustedWhole = whole < 60 ? whole : whole - 1;
  const date = new Date(EXCEL_EPOCH_UTC_MS + adjustedWhole * MS_PER_DAY);
  if (Number.isNaN(date.getTime())) {
    return undefined;
  }

  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate()
  };
}

export function excelDatePartsToSerial(year: number, month: number, day: number): number | undefined {
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return undefined;
  }

  let adjustedYear = Math.trunc(year);
  const adjustedMonth = Math.trunc(month);
  const adjustedDay = Math.trunc(day);

  if (adjustedYear >= 0 && adjustedYear <= 1899) {
    adjustedYear += 1900;
  }

  if (adjustedYear < 0 || adjustedYear > 9999) {
    return undefined;
  }

  if (adjustedYear === 1900 && adjustedMonth === 2 && adjustedDay === 29) {
    return 60;
  }

  const normalized = new Date(Date.UTC(adjustedYear, adjustedMonth - 1, adjustedDay));
  if (Number.isNaN(normalized.getTime())) {
    return undefined;
  }

  return utcDateToExcelSerial(normalized);
}

export function utcDateToExcelSerial(date: Date): number {
  const midnightUtcMs = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  let daySerial = (midnightUtcMs - EXCEL_EPOCH_UTC_MS) / MS_PER_DAY;
  if (midnightUtcMs >= EXCEL_LEAP_BUG_CUTOFF_UTC_MS) {
    daySerial += 1;
  }

  const dayFraction = (
    date.getUTCHours() * 3_600_000
    + date.getUTCMinutes() * 60_000
    + date.getUTCSeconds() * 1_000
    + date.getUTCMilliseconds()
  ) / MS_PER_DAY;

  return daySerial + dayFraction;
}

export function addMonthsToExcelDate(serial: number, offsetMonths: number): number | undefined {
  const start = excelSerialToDateParts(serial);
  if (!start || !Number.isFinite(offsetMonths)) {
    return undefined;
  }

  const shifted = normalizeMonth(start.year, start.month + Math.trunc(offsetMonths));
  const day = Math.min(start.day, daysInExcelMonth(shifted.year, shifted.month));

  if (shifted.year < 0 || shifted.year > 9999) {
    return undefined;
  }

  if (shifted.year === 1900 && shifted.month === 2 && day === 29) {
    return 60;
  }

  return excelDatePartsToSerial(shifted.year, shifted.month, day);
}

export function endOfMonthExcelDate(serial: number, offsetMonths: number): number | undefined {
  const start = excelSerialToDateParts(serial);
  if (!start || !Number.isFinite(offsetMonths)) {
    return undefined;
  }

  const shifted = normalizeMonth(start.year, start.month + Math.trunc(offsetMonths));
  const day = daysInExcelMonth(shifted.year, shifted.month);

  if (shifted.year < 0 || shifted.year > 9999) {
    return undefined;
  }

  if (isExcelLeapBugDate({ year: shifted.year, month: shifted.month, day })) {
    return 60;
  }

  return excelDatePartsToSerial(shifted.year, shifted.month, day);
}

function normalizeTimeSerial(hours: number, minutes: number, seconds: number): number | undefined {
  if (!Number.isFinite(hours) || !Number.isFinite(minutes) || !Number.isFinite(seconds)) {
    return undefined;
  }
  if (hours < 0 || minutes < 0 || seconds < 0) {
    return undefined;
  }
  if (hours > 32_767 || minutes > 32_767 || seconds > 32_767) {
    return undefined;
  }
  const totalSeconds = Math.trunc(hours) * 3600 + Math.trunc(minutes) * 60 + Math.trunc(seconds);
  return ((totalSeconds % SECONDS_PER_DAY) + SECONDS_PER_DAY) % SECONDS_PER_DAY / SECONDS_PER_DAY;
}

export function createDateBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length !== 3) {
      return valueError();
    }

    const year = truncArg(args[0]!);
    const month = truncArg(args[1]!);
    const day = truncArg(args[2]!);
    if (typeof year !== "number") return year;
    if (typeof month !== "number") return month;
    if (typeof day !== "number") return day;

    const serial = excelDatePartsToSerial(year, month, day);
    return serial === undefined ? valueError() : numberResult(serial);
  };
}

function createDatePartBuiltin(part: keyof ExcelDateParts): Builtin {
  return (value) => {
    const error = firstError([value]);
    if (error) {
      return error;
    }

    const serial = coerceNumber(value);
    if (serial === undefined) {
      return valueError();
    }

    const parts = excelSerialToDateParts(serial);
    return parts ? numberResult(parts[part]) : valueError();
  };
}

function createTimeBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length !== 3) {
      return valueError();
    }

    const hour = truncArg(args[0]!);
    const minute = truncArg(args[1]!);
    const second = truncArg(args[2]!);
    if (typeof hour !== "number") return hour;
    if (typeof minute !== "number") return minute;
    if (typeof second !== "number") return second;

    const serial = normalizeTimeSerial(hour, minute, second);
    return serial === undefined ? valueError() : numberResult(serial);
  };
}

function createTimePartBuiltin(part: "hour" | "minute" | "second"): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    const [value] = args;
    if (value === undefined) {
      return valueError();
    }

    const serial = coerceNumber(value);
    if (serial === undefined) {
      return valueError();
    }
    const seconds = normalizeSecondOfDay(serial);
    if (seconds === undefined) {
      return valueError();
    }

    switch (part) {
      case "hour":
        return numberResult(Math.floor(seconds / 3600));
      case "minute":
        return numberResult(Math.floor((seconds % 3600) / 60));
      case "second":
        return numberResult(seconds % 60);
    }
  };
}

function createWeekdayBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length < 1 || args.length > 2) {
      return valueError();
    }
    const serial = coerceNumber(args[0]!);
    if (serial === undefined || serial < 0) {
      return valueError();
    }

    const whole = floorDateSerial(serial);
    const adjustedWhole = whole < 60 ? whole : whole - 1;
    const sundayOne = ((adjustedWhole % 7) + 7) % 7 + 1;
    if (args.length === 1) {
      return numberResult(sundayOne);
    }

    const returnType = truncArg(args[1]!);
    if (typeof returnType !== "number") {
      return returnType;
    }
    if (returnType === 3) {
      return numberResult(sundayOne === 1 ? 6 : sundayOne - 2);
    }

    const startDayMap: Record<number, number> = {
      1: 1,
      2: 2,
      11: 2,
      12: 3,
      13: 4,
      14: 5,
      15: 6,
      16: 7,
      17: 1
    };
    const startDay = startDayMap[returnType];
    if (startDay === undefined) {
      return valueError();
    }
    return numberResult(((sundayOne - startDay + 7) % 7) + 1);
  };
}

function createDaysBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length !== 2) {
      return valueError();
    }
    const endSerial = truncArg(args[0]!);
    const startSerial = truncArg(args[1]!);
    if (typeof endSerial !== "number") {
      return endSerial;
    }
    if (typeof startSerial !== "number") {
      return startSerial;
    }
    return numberResult(endSerial - startSerial);
  };
}

function createWeeknumBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length < 1 || args.length > 2) {
      return valueError();
    }

    const serial = truncArg(args[0]!);
    if (typeof serial !== "number") {
      return serial;
    }

    const returnType = args[1] === undefined ? 1 : truncArg(args[1]);
    if (typeof returnType !== "number") {
      return returnType;
    }

    const dateParts = excelSerialToDateParts(serial);
    if (!dateParts) {
      return valueError();
    }

    let weekStartDay: number;
    if (returnType === 1 || returnType === 17) {
      weekStartDay = 0;
    } else if (returnType === 2 || returnType === 11) {
      weekStartDay = 1;
    } else if (returnType === 12) {
      weekStartDay = 2;
    } else if (returnType === 13) {
      weekStartDay = 3;
    } else if (returnType === 14) {
      weekStartDay = 4;
    } else if (returnType === 15) {
      weekStartDay = 5;
    } else if (returnType === 16) {
      weekStartDay = 6;
    } else {
      return valueError();
    }

    const serialJan1 = excelDatePartsToSerial(dateParts.year, 1, 1);
    if (serialJan1 === undefined) {
      return valueError();
    }

    const adjustedJan1 = serialJan1 < 60 ? Math.floor(serialJan1) : Math.floor(serialJan1) - 1;
    const jan1Weekday = ((adjustedJan1 % 7) + 7) % 7;
    const shift = ((jan1Weekday - weekStartDay) + 7) % 7;

    let dayOfYear = dateParts.day;
    for (let month = 1; month < dateParts.month; month += 1) {
      dayOfYear += daysInExcelMonth(dateParts.year, month);
    }

    return numberResult(Math.floor((dayOfYear - 1 + shift) / 7) + 1);
  };
}

function isWeekendSerial(serial: number): boolean {
  const whole = floorDateSerial(serial);
  const adjustedWhole = whole < 60 ? whole : whole - 1;
  const dow = ((adjustedWhole % 7) + 7) % 7;
  return dow === 0 || dow === 6;
}

function normalizeHolidayDateSet(holidays: readonly CellValue[] | undefined): Set<number> | CellValue {
  if (!holidays || holidays.length === 0) {
    return new Set<number>();
  }

  const set = new Set<number>();
  for (const holiday of holidays) {
    const raw = coerceNumber(holiday);
    if (raw === undefined) {
      return valueError();
    }
    set.add(Math.trunc(raw));
  }
  return set;
}

function createWorkdayBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length < 2) {
      return valueError();
    }

    const start = truncArg(args[0]!);
    const offset = truncArg(args[1]!);
    if (typeof start !== "number") {
      return start;
    }
    if (typeof offset !== "number") {
      return offset;
    }

    const holidays = normalizeHolidayDateSet(args.slice(2));
    if (isErrorValue(holidays)) {
      return holidays;
    }

    const isWorkday = (value: number): boolean => !isWeekendSerial(value) && !holidays.has(Math.trunc(value));
    let cursor = Math.trunc(start);
    const direction = offset >= 0 ? 1 : -1;

    while (!isWorkday(cursor)) {
      cursor += direction;
    }

    let remaining = Math.abs(offset);
    while (remaining > 0) {
      cursor += direction;
      if (isWorkday(cursor)) {
        remaining -= 1;
      }
    }
    return numberResult(cursor);
  };
}

function createNetworkdaysBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length < 2) {
      return valueError();
    }

    const start = truncArg(args[0]!);
    const end = truncArg(args[1]!);
    if (typeof start !== "number") {
      return start;
    }
    if (typeof end !== "number") {
      return end;
    }

    const holidays = normalizeHolidayDateSet(args.slice(2));
    if (isErrorValue(holidays)) {
      return holidays;
    }

    const isWorkday = (value: number): boolean => !isWeekendSerial(value) && !holidays.has(Math.trunc(value));
    const step = start <= end ? 1 : -1;
    let count = 0;
    for (let cursor = Math.trunc(start); ; cursor += step) {
      if (isWorkday(cursor)) {
        count += step;
      }
      if (cursor === Math.trunc(end)) {
        break;
      }
    }
    return numberResult(count);
  };
}

export function createTodayBuiltin(now: DateTimeProvider = () => new Date()): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length > 0) {
      return valueError();
    }
    return numberResult(Math.floor(utcDateToExcelSerial(now())));
  };
}

export function createNowBuiltin(now: DateTimeProvider = () => new Date()): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length > 0) {
      return valueError();
    }
    return numberResult(utcDateToExcelSerial(now()));
  };
}

export function createRandBuiltin(random: RandomProvider = () => Math.random()): Builtin {
  return (...args) => {
    const error = firstError(args);
    if (error) {
      return error;
    }
    if (args.length > 0) {
      return valueError();
    }

    const next = random();
    if (!Number.isFinite(next)) {
      return valueError();
    }

    const bounded = Math.min(Math.max(next, 0), 1 - Number.EPSILON);
    return numberResult(bounded);
  };
}

export function createEdateBuiltin(): Builtin {
  return (startDate, months) => {
    const error = firstError([startDate, months]);
    if (error) {
      return error;
    }

    const startSerial = coerceNumber(startDate);
    const monthOffset = truncArg(months);
    if (startSerial === undefined) {
      return valueError();
    }
    if (typeof monthOffset !== "number") {
      return monthOffset;
    }

    const serial = addMonthsToExcelDate(startSerial, monthOffset);
    return serial === undefined ? valueError() : numberResult(serial);
  };
}

export function createEomonthBuiltin(): Builtin {
  return (startDate, months) => {
    const error = firstError([startDate, months]);
    if (error) {
      return error;
    }

    const startSerial = coerceNumber(startDate);
    const monthOffset = truncArg(months);
    if (startSerial === undefined) {
      return valueError();
    }
    if (typeof monthOffset !== "number") {
      return monthOffset;
    }

    const serial = endOfMonthExcelDate(startSerial, monthOffset);
    return serial === undefined ? valueError() : numberResult(serial);
  };
}

const datetimePlaceholderBuiltins = createBlockedBuiltinMap(datetimePlaceholderBuiltinNames);

export const datetimeBuiltins: Record<string, Builtin> = {
  DATE: createDateBuiltin(),
  YEAR: createDatePartBuiltin("year"),
  MONTH: createDatePartBuiltin("month"),
  DAY: createDatePartBuiltin("day"),
  TIME: createTimeBuiltin(),
  HOUR: createTimePartBuiltin("hour"),
  MINUTE: createTimePartBuiltin("minute"),
  SECOND: createTimePartBuiltin("second"),
  WEEKDAY: createWeekdayBuiltin(),
  DAYS: createDaysBuiltin(),
  WEEKNUM: createWeeknumBuiltin(),
  WORKDAY: createWorkdayBuiltin(),
  NETWORKDAYS: createNetworkdaysBuiltin(),
  TODAY: createTodayBuiltin(),
  NOW: createNowBuiltin(),
  RAND: createRandBuiltin(),
  EDATE: createEdateBuiltin(),
  EOMONTH: createEomonthBuiltin(),
  ...datetimePlaceholderBuiltins
};
