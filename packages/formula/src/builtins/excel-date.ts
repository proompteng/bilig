export interface ExcelDateParts {
  year: number
  month: number
  day: number
}

export type ExcelDateSystem = '1900' | '1904'

export const MS_PER_DAY = 86_400_000

const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 31)
const EXCEL_LEAP_BUG_CUTOFF_UTC_MS = Date.UTC(1900, 2, 1)
const EXCEL_1904_EPOCH_UTC_MS = Date.UTC(1904, 0, 1)

export function isLeapYear(year: number): boolean {
  return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)
}

export function isValidYearfracBasis(basis: number): boolean {
  return basis === 0 || basis === 1 || basis === 2 || basis === 3 || basis === 4
}

export function floorDateSerial(serial: number): number {
  return Math.floor(serial)
}

function isExcelLeapBugDate(parts: ExcelDateParts): boolean {
  return parts.year === 1900 && parts.month === 2 && parts.day === 29
}

export function daysInExcelMonth(year: number, month: number): number {
  if (year === 1900 && month === 2) {
    return 29
  }
  return new Date(Date.UTC(year, month, 0)).getUTCDate()
}

function normalizeMonth(year: number, month: number): { year: number; month: number } {
  const zeroBased = year * 12 + (month - 1)
  const normalizedYear = Math.floor(zeroBased / 12)
  const normalizedMonth = zeroBased - normalizedYear * 12 + 1
  return { year: normalizedYear, month: normalizedMonth }
}

export function excelSerialToDateParts(serial: number, dateSystem: ExcelDateSystem = '1900'): ExcelDateParts | undefined {
  if (!Number.isFinite(serial)) {
    return undefined
  }

  const whole = floorDateSerial(serial)
  if (dateSystem === '1904') {
    const date = new Date(EXCEL_1904_EPOCH_UTC_MS + whole * MS_PER_DAY)
    if (Number.isNaN(date.getTime())) {
      return undefined
    }
    return {
      year: date.getUTCFullYear(),
      month: date.getUTCMonth() + 1,
      day: date.getUTCDate(),
    }
  }

  if (whole === 60) {
    return { year: 1900, month: 2, day: 29 }
  }

  const adjustedWhole = whole < 60 ? whole : whole - 1
  const date = new Date(EXCEL_EPOCH_UTC_MS + adjustedWhole * MS_PER_DAY)
  if (Number.isNaN(date.getTime())) {
    return undefined
  }

  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
  }
}

export function excelDatePartsToSerial(year: number, month: number, day: number, dateSystem: ExcelDateSystem = '1900'): number | undefined {
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return undefined
  }

  let adjustedYear = Math.trunc(year)
  const adjustedMonth = Math.trunc(month)
  const adjustedDay = Math.trunc(day)

  if (adjustedYear >= 0 && adjustedYear <= 1899) {
    adjustedYear += 1900
  }

  if (adjustedYear < 0 || adjustedYear > 9999) {
    return undefined
  }

  if (dateSystem === '1900' && adjustedYear === 1900 && adjustedMonth === 2 && adjustedDay === 29) {
    return 60
  }

  const normalized = new Date(Date.UTC(adjustedYear, adjustedMonth - 1, adjustedDay))
  if (Number.isNaN(normalized.getTime())) {
    return undefined
  }

  return utcDateToExcelSerial(normalized, dateSystem)
}

export function utcDateToExcelSerial(date: Date, dateSystem: ExcelDateSystem = '1900'): number {
  const midnightUtcMs = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate())
  if (dateSystem === '1904') {
    const daySerial = (midnightUtcMs - EXCEL_1904_EPOCH_UTC_MS) / MS_PER_DAY
    const dayFraction =
      (date.getUTCHours() * 3_600_000 + date.getUTCMinutes() * 60_000 + date.getUTCSeconds() * 1_000 + date.getUTCMilliseconds()) /
      MS_PER_DAY
    return daySerial + dayFraction
  }

  let daySerial = (midnightUtcMs - EXCEL_EPOCH_UTC_MS) / MS_PER_DAY
  if (midnightUtcMs >= EXCEL_LEAP_BUG_CUTOFF_UTC_MS) {
    daySerial += 1
  }

  const dayFraction =
    (date.getUTCHours() * 3_600_000 + date.getUTCMinutes() * 60_000 + date.getUTCSeconds() * 1_000 + date.getUTCMilliseconds()) / MS_PER_DAY

  return daySerial + dayFraction
}

export function excelSerialWeekdayIndex(serial: number, dateSystem: ExcelDateSystem = '1900'): number | undefined {
  if (!Number.isFinite(serial)) {
    return undefined
  }
  const whole = floorDateSerial(serial)
  if (whole < 0) {
    return undefined
  }
  if (dateSystem === '1900') {
    const adjustedWhole = whole < 60 ? whole : whole - 1
    return ((adjustedWhole % 7) + 7) % 7
  }
  const parts = excelSerialToDateParts(serial, dateSystem)
  if (!parts) {
    return undefined
  }
  return new Date(Date.UTC(parts.year, parts.month - 1, parts.day)).getUTCDay()
}

export function addMonthsToExcelDate(serial: number, offsetMonths: number, dateSystem: ExcelDateSystem = '1900'): number | undefined {
  const start = excelSerialToDateParts(serial, dateSystem)
  if (!start || !Number.isFinite(offsetMonths)) {
    return undefined
  }

  const shifted = normalizeMonth(start.year, start.month + Math.trunc(offsetMonths))
  const day = Math.min(start.day, daysInExcelMonth(shifted.year, shifted.month))

  if (shifted.year < 0 || shifted.year > 9999) {
    return undefined
  }

  if (dateSystem === '1900' && shifted.year === 1900 && shifted.month === 2 && day === 29) {
    return 60
  }

  return excelDatePartsToSerial(shifted.year, shifted.month, day, dateSystem)
}

export function endOfMonthExcelDate(serial: number, offsetMonths: number, dateSystem: ExcelDateSystem = '1900'): number | undefined {
  const start = excelSerialToDateParts(serial, dateSystem)
  if (!start || !Number.isFinite(offsetMonths)) {
    return undefined
  }

  const shifted = normalizeMonth(start.year, start.month + Math.trunc(offsetMonths))
  const day = daysInExcelMonth(shifted.year, shifted.month)

  if (shifted.year < 0 || shifted.year > 9999) {
    return undefined
  }

  if (dateSystem === '1900' && isExcelLeapBugDate({ year: shifted.year, month: shifted.month, day })) {
    return 60
  }

  return excelDatePartsToSerial(shifted.year, shifted.month, day, dateSystem)
}

export function yearFracByBasis(
  startSerial: number,
  endSerial: number,
  basis: number,
  dateSystem: ExcelDateSystem = '1900',
): number | undefined {
  if (!isValidYearfracBasis(basis)) {
    return undefined
  }

  let start = startSerial
  let end = endSerial
  if (start > end) {
    ;[start, end] = [end, start]
  }

  const startParts = excelSerialToDateParts(start, dateSystem)
  const endParts = excelSerialToDateParts(end, dateSystem)
  if (startParts === undefined || endParts === undefined) {
    return undefined
  }

  let startDay = startParts.day
  let startMonth = startParts.month
  let startYear = startParts.year
  let endDay = endParts.day
  let endMonth = endParts.month
  let endYear = endParts.year

  let totalDays: number
  switch (basis) {
    case 0:
      if (startDay === 31) {
        startDay -= 1
      }
      if (startDay === 30 && endDay === 31) {
        endDay -= 1
      } else if (startMonth === 2 && startDay === (isLeapYear(startYear) ? 29 : 28)) {
        startDay = 30
        if (endMonth === 2 && endDay === (isLeapYear(endYear) ? 29 : 28)) {
          endDay = 30
        }
      }
      totalDays = (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay)
      break
    case 1:
    case 2:
    case 3:
      totalDays = end - start
      break
    case 4:
      if (startDay === 31) {
        startDay -= 1
      }
      if (endDay === 31) {
        endDay -= 1
      }
      totalDays = (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay)
      break
    default:
      return undefined
  }

  let daysInYear: number
  switch (basis) {
    case 1: {
      const yearLength = (year: number) => (isLeapYear(year) ? 366 : 365)
      if (startYear === endYear) {
        daysInYear = yearLength(startYear)
        break
      }
      const crossesMultipleYears = endYear !== startYear + 1 || endMonth < startMonth || (endMonth === startMonth && endDay > startDay)
      if (crossesMultipleYears) {
        let total = 0
        for (let year = startYear; year <= endYear; year += 1) {
          total += yearLength(year)
        }
        daysInYear = total / (endYear - startYear + 1)
      } else {
        const startsInLeapYear = isLeapYear(startYear) && (startMonth < 2 || (startMonth === 2 && startDay <= 29))
        const endsInLeapYear = isLeapYear(endYear) && (endMonth > 2 || (endMonth === 2 && endDay === 29))
        daysInYear = startsInLeapYear || endsInLeapYear ? 366 : 365
      }
      break
    }
    case 3:
      daysInYear = 365
      break
    case 0:
    case 2:
    case 4:
      daysInYear = 360
      break
    default:
      return undefined
  }

  return totalDays / daysInYear
}
