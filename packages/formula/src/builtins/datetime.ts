import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";

export type Builtin = (...args: CellValue[]) => CellValue;
export type DateTimeProvider = () => Date;
export type RandomProvider = () => number;

export interface ExcelDateParts {
  year: number;
  month: number;
  day: number;
}

const MS_PER_DAY = 86_400_000;
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

export const datetimeBuiltins: Record<string, Builtin> = {
  DATE: createDateBuiltin(),
  YEAR: createDatePartBuiltin("year"),
  MONTH: createDatePartBuiltin("month"),
  DAY: createDatePartBuiltin("day"),
  TODAY: createTodayBuiltin(),
  NOW: createNowBuiltin(),
  RAND: createRandBuiltin(),
  EDATE: createEdateBuiltin(),
  EOMONTH: createEomonthBuiltin()
};
