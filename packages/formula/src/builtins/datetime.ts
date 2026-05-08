import { ValueTag } from '@bilig/protocol'
import type { CellValue } from '@bilig/protocol'
import { coerceNumber, coerceText, firstError, integerValue, isErrorValue, numberResult, truncArg, valueError } from './cell-value-utils.js'
import {
  MS_PER_DAY,
  addMonthsToExcelDate,
  daysInExcelMonth,
  endOfMonthExcelDate,
  excelDatePartsToSerial,
  excelSerialToDateParts,
  excelSerialWeekdayIndex,
  floorDateSerial,
  isValidYearfracBasis,
  utcDateToExcelSerial,
  yearFracByBasis,
  type ExcelDateSystem,
  type ExcelDateParts,
} from './excel-date.js'
import { createBlockedBuiltinMap, datetimePlaceholderBuiltinNames } from './placeholder.js'
export {
  addMonthsToExcelDate,
  endOfMonthExcelDate,
  excelDatePartsToSerial,
  excelSerialToDateParts,
  excelSerialWeekdayIndex,
  utcDateToExcelSerial,
} from './excel-date.js'
export type { ExcelDateParts, ExcelDateSystem } from './excel-date.js'

export type Builtin = (...args: CellValue[]) => CellValue
export type DateTimeProvider = () => Date
export type RandomProvider = () => number

const SECONDS_PER_DAY = 86_400

function parseDateValueFromText(raw: string, dateSystem: ExcelDateSystem): number | undefined {
  const trimmed = raw.trim()
  if (trimmed === '') {
    return undefined
  }
  const parsed = new Date(trimmed)
  if (Number.isNaN(parsed.getTime())) {
    return undefined
  }
  return Math.floor(utcDateToExcelSerial(parsed, dateSystem))
}

function createDays360Builtin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2 || args.length > 3) {
      return valueError()
    }

    const startSerial = truncArg(args[0]!)
    const endSerial = truncArg(args[1]!)
    const method = args[2] === undefined ? 0 : integerValue(args[2], 0)
    if (method === undefined || (method !== 0 && method !== 1)) {
      return valueError()
    }

    if (typeof startSerial !== 'number') {
      return startSerial
    }
    if (typeof endSerial !== 'number') {
      return endSerial
    }
    const startParts = excelSerialToDateParts(startSerial, dateSystem)
    const endParts = excelSerialToDateParts(endSerial, dateSystem)
    if (!startParts || !endParts) {
      return valueError()
    }

    let startDay = startParts.day
    let endDay = endParts.day

    if (method === 0) {
      if (startDay === 31) {
        startDay = 30
      }
      if (endDay === 31 && startDay >= 30) {
        endDay = 30
      }
    } else {
      if (startDay === 31) {
        startDay = 30
      }
      if (endDay === 31) {
        endDay = 30
      }
    }

    return numberResult((endParts.year - startParts.year) * 360 + (endParts.month - startParts.month) * 30 + (endDay - startDay))
  }
}

function createIsoWeeknumBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 1) {
      return valueError()
    }

    const serial = truncArg(args[0]!)
    if (typeof serial !== 'number') {
      return serial
    }

    const parts = excelSerialToDateParts(serial, dateSystem)
    if (parts === undefined) {
      return valueError()
    }

    const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day))
    const dow = date.getUTCDay()
    const dayShift = dow === 0 ? 7 : dow
    const shifted = new Date(date.getTime())
    shifted.setUTCDate(date.getUTCDate() + 4 - dayShift)
    const yearStart = new Date(Date.UTC(shifted.getUTCFullYear(), 0, 1))
    const dayOfYear = Math.floor((shifted.getTime() - yearStart.getTime()) / MS_PER_DAY) + 1
    return numberResult(Math.floor((dayOfYear - 1) / 7) + 1)
  }
}

function createTimeValueBuiltin(): Builtin {
  return (value) => {
    if (value === undefined) {
      return valueError()
    }
    const error = firstError([value])
    if (error) {
      return error
    }
    const text = coerceText(value)
    if (text === undefined) {
      return valueError()
    }

    const trimmed = text.trim()
    const amPmMatch = trimmed.match(/^(.+?)\s+([aApP][mM])$/)
    const hasMeridiem = amPmMatch !== null
    const timeText = hasMeridiem ? (amPmMatch?.[1] ?? '') : trimmed
    const timeParts = timeText.split(':')
    if (timeParts.length < 2 || timeParts.length > 3) {
      return valueError()
    }

    const [hoursText, minutesText, secondsText = '0'] = timeParts
    const hours = Number(hoursText)
    const minutes = Number(minutesText)
    const seconds = Number(secondsText)
    if (
      Number.isNaN(hours) ||
      Number.isNaN(minutes) ||
      Number.isNaN(seconds) ||
      !Number.isFinite(hours) ||
      !Number.isFinite(minutes) ||
      !Number.isFinite(seconds)
    ) {
      return valueError()
    }

    const truncHours = Math.trunc(hours)
    const truncMinutes = Math.trunc(minutes)
    const truncSeconds = Math.trunc(seconds)
    const hasPm = hasMeridiem && amPmMatch?.[2]?.toLowerCase() === 'pm'

    if (truncMinutes < 0 || truncMinutes > 59 || truncSeconds < 0 || truncSeconds > 59) {
      return valueError()
    }

    let hourValue = truncHours
    if (hasMeridiem) {
      if (truncHours < 1 || truncHours > 12) {
        return valueError()
      }
      if (truncHours === 12) {
        hourValue = hasPm ? 12 : 0
      } else if (hasPm) {
        hourValue = truncHours + 12
      }
    } else if (truncHours === 24 && truncMinutes === 0 && truncSeconds === 0) {
      hourValue = 0
    } else if (truncHours < 0 || truncHours > 23) {
      return valueError()
    }

    return numberResult((hourValue * 3600 + truncMinutes * 60 + truncSeconds) / SECONDS_PER_DAY)
  }
}

function createYearfracBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2 || args.length > 3) {
      return valueError()
    }

    const startSerial = truncArg(args[0]!)
    const endSerial = truncArg(args[1]!)
    const basis = args[2] === undefined ? 0 : integerValue(args[2])
    if (typeof startSerial !== 'number' || typeof endSerial !== 'number' || basis === undefined || !isValidYearfracBasis(basis)) {
      return valueError()
    }

    const fraction = yearFracByBasis(startSerial, endSerial, basis, dateSystem)
    return fraction === undefined ? valueError() : numberResult(fraction)
  }
}

function normalizeSecondOfDay(serial: number): number | undefined {
  if (!Number.isFinite(serial) || serial < 0) {
    return undefined
  }
  const fraction = serial - floorDateSerial(serial)
  const normalizedFraction = fraction < 0 ? fraction + 1 : fraction
  return Math.floor(normalizedFraction * SECONDS_PER_DAY + 1e-9) % SECONDS_PER_DAY
}

function datedifValue(startSerial: number, endSerial: number, unit: string, dateSystem: ExcelDateSystem): number | undefined {
  if (startSerial > endSerial) {
    return undefined
  }
  const start = excelSerialToDateParts(startSerial, dateSystem)
  const end = excelSerialToDateParts(endSerial, dateSystem)
  if (!start || !end) {
    return undefined
  }
  const totalDays = Math.trunc(endSerial) - Math.trunc(startSerial)
  const totalMonths = (end.year - start.year) * 12 + (end.month - start.month) - (end.day < start.day ? 1 : 0)
  const totalYears = end.year - start.year - (end.month < start.month || (end.month === start.month && end.day < start.day) ? 1 : 0)

  switch (unit) {
    case 'D':
      return totalDays
    case 'M':
      return totalMonths
    case 'Y':
      return totalYears
    case 'YM':
      return ((totalMonths % 12) + 12) % 12
    case 'YD': {
      let comparisonYear = end.year
      let comparison = excelDatePartsToSerial(comparisonYear, start.month, start.day, dateSystem)
      if (comparison === undefined || comparison > endSerial) {
        comparisonYear -= 1
        comparison = excelDatePartsToSerial(comparisonYear, start.month, start.day, dateSystem)
      }
      return comparison === undefined ? undefined : Math.trunc(endSerial) - Math.trunc(comparison)
    }
    case 'MD':
      if (end.day >= start.day) {
        return end.day - start.day
      }
      return daysInExcelMonth(end.year, end.month === 1 ? 12 : end.month - 1) - start.day + end.day
    default:
      return undefined
  }
}

function normalizeTimeSerial(hours: number, minutes: number, seconds: number): number | undefined {
  if (!Number.isFinite(hours) || !Number.isFinite(minutes) || !Number.isFinite(seconds)) {
    return undefined
  }
  if (hours < 0 || minutes < 0 || seconds < 0) {
    return undefined
  }
  if (hours > 32_767 || minutes > 32_767 || seconds > 32_767) {
    return undefined
  }
  const totalSeconds = Math.trunc(hours) * 3600 + Math.trunc(minutes) * 60 + Math.trunc(seconds)
  return (((totalSeconds % SECONDS_PER_DAY) + SECONDS_PER_DAY) % SECONDS_PER_DAY) / SECONDS_PER_DAY
}

export function createDateBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 3) {
      return valueError()
    }

    const year = truncArg(args[0]!)
    const month = truncArg(args[1]!)
    const day = truncArg(args[2]!)
    if (typeof year !== 'number') return year
    if (typeof month !== 'number') return month
    if (typeof day !== 'number') return day

    const serial = excelDatePartsToSerial(year, month, day, dateSystem)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

export function createDateValueBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (dateText) => {
    if (dateText === undefined) {
      return valueError()
    }
    const error = firstError([dateText])
    if (error) {
      return error
    }

    const asNumber = toNumberValueDateValue(dateText)
    if (asNumber !== undefined) {
      return numberResult(asNumber)
    }

    const text = coerceText(dateText)
    if (text === undefined) {
      return valueError()
    }

    const serial = parseDateValueFromText(text, dateSystem)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

function toNumberValueDateValue(value: CellValue): number | undefined {
  const numeric = coerceNumber(value)
  if (numeric === undefined) {
    return undefined
  }
  const truncated = Math.trunc(numeric)
  return Number.isFinite(truncated) ? truncated : undefined
}

function createDatePartBuiltin(part: keyof ExcelDateParts, dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (value) => {
    const error = firstError([value])
    if (error) {
      return error
    }

    const serial = coerceNumber(value)
    if (serial === undefined) {
      return valueError()
    }

    const parts = excelSerialToDateParts(serial, dateSystem)
    return parts ? numberResult(parts[part]) : valueError()
  }
}

function createTimeBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 3) {
      return valueError()
    }

    const hour = truncArg(args[0]!)
    const minute = truncArg(args[1]!)
    const second = truncArg(args[2]!)
    if (typeof hour !== 'number') return hour
    if (typeof minute !== 'number') return minute
    if (typeof second !== 'number') return second

    const serial = normalizeTimeSerial(hour, minute, second)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

function createTimePartBuiltin(part: 'hour' | 'minute' | 'second'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    const [value] = args
    if (value === undefined) {
      return valueError()
    }

    const serial = coerceNumber(value)
    if (serial === undefined) {
      return valueError()
    }
    const seconds = normalizeSecondOfDay(serial)
    if (seconds === undefined) {
      return valueError()
    }

    switch (part) {
      case 'hour':
        return numberResult(Math.floor(seconds / 3600))
      case 'minute':
        return numberResult(Math.floor((seconds % 3600) / 60))
      case 'second':
        return numberResult(seconds % 60)
    }
  }
}

function createWeekdayBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 1 || args.length > 2) {
      return valueError()
    }
    const serial = coerceNumber(args[0]!)
    if (serial === undefined || serial < 0) {
      return valueError()
    }

    const weekdayIndex = excelSerialWeekdayIndex(serial, dateSystem)
    if (weekdayIndex === undefined) {
      return valueError()
    }
    const sundayOne = weekdayIndex + 1
    if (args.length === 1) {
      return numberResult(sundayOne)
    }

    const returnType = truncArg(args[1]!)
    if (typeof returnType !== 'number') {
      return returnType
    }
    if (returnType === 3) {
      return numberResult(sundayOne === 1 ? 6 : sundayOne - 2)
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
      17: 1,
    }
    const startDay = startDayMap[returnType]
    if (startDay === undefined) {
      return valueError()
    }
    return numberResult(((sundayOne - startDay + 7) % 7) + 1)
  }
}

function createDaysBuiltin(): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 2) {
      return valueError()
    }
    const endSerial = truncArg(args[0]!)
    const startSerial = truncArg(args[1]!)
    if (typeof endSerial !== 'number') {
      return endSerial
    }
    if (typeof startSerial !== 'number') {
      return startSerial
    }
    return numberResult(endSerial - startSerial)
  }
}

function createWeeknumBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 1 || args.length > 2) {
      return valueError()
    }

    const serial = truncArg(args[0]!)
    if (typeof serial !== 'number') {
      return serial
    }

    const returnType = args[1] === undefined ? 1 : truncArg(args[1])
    if (typeof returnType !== 'number') {
      return returnType
    }

    const dateParts = excelSerialToDateParts(serial, dateSystem)
    if (!dateParts) {
      return valueError()
    }

    let weekStartDay: number
    if (returnType === 1 || returnType === 17) {
      weekStartDay = 0
    } else if (returnType === 2 || returnType === 11) {
      weekStartDay = 1
    } else if (returnType === 12) {
      weekStartDay = 2
    } else if (returnType === 13) {
      weekStartDay = 3
    } else if (returnType === 14) {
      weekStartDay = 4
    } else if (returnType === 15) {
      weekStartDay = 5
    } else if (returnType === 16) {
      weekStartDay = 6
    } else {
      return valueError()
    }

    const serialJan1 = excelDatePartsToSerial(dateParts.year, 1, 1, dateSystem)
    if (serialJan1 === undefined) {
      return valueError()
    }

    const jan1Weekday = excelSerialWeekdayIndex(serialJan1, dateSystem)
    if (jan1Weekday === undefined) {
      return valueError()
    }
    const shift = (jan1Weekday - weekStartDay + 7) % 7

    let dayOfYear = dateParts.day
    for (let month = 1; month < dateParts.month; month += 1) {
      dayOfYear += daysInExcelMonth(dateParts.year, month)
    }

    return numberResult(Math.floor((dayOfYear - 1 + shift) / 7) + 1)
  }
}

function isWeekendSerial(serial: number, dateSystem: ExcelDateSystem): boolean {
  const dow = excelSerialWeekdayIndex(serial, dateSystem)
  return dow === 0 || dow === 6
}

function weekendSerialDay(serial: number, dateSystem: ExcelDateSystem): number | undefined {
  return excelSerialWeekdayIndex(serial, dateSystem)
}

function weekendMaskFromCode(code: number): Set<number> | undefined {
  const twoDayWeekendMap: Record<number, readonly number[]> = {
    1: [6, 0],
    2: [0, 1],
    3: [1, 2],
    4: [2, 3],
    5: [3, 4],
    6: [4, 5],
    7: [5, 6],
  }
  if (code >= 1 && code <= 7) {
    return new Set(twoDayWeekendMap[code])
  }
  if (code >= 11 && code <= 17) {
    return new Set([(code - 10) % 7])
  }
  return undefined
}

function weekendMaskFromString(maskText: string): Set<number> | undefined {
  const trimmed = maskText.trim()
  if (!/^[01]{7}$/.test(trimmed) || trimmed === '1111111') {
    return undefined
  }
  const days = new Set<number>()
  for (let index = 0; index < trimmed.length; index += 1) {
    if (trimmed[index] !== '1') {
      continue
    }
    const dow = index === 6 ? 0 : index + 1
    days.add(dow)
  }
  return days
}

function normalizeWeekendMask(weekendArg: CellValue | undefined): Set<number> | CellValue {
  if (weekendArg === undefined) {
    return new Set([0, 6])
  }
  if (weekendArg.tag === ValueTag.Error) {
    return weekendArg
  }
  if (weekendArg.tag === ValueTag.String) {
    const mask = weekendMaskFromString(weekendArg.value)
    return mask ?? valueError()
  }
  const code = integerValue(weekendArg)
  if (code === undefined) {
    return valueError()
  }
  const mask = weekendMaskFromCode(code)
  return mask ?? valueError()
}

function normalizeHolidayDateSet(holidays: readonly CellValue[] | undefined): Set<number> | CellValue {
  if (!holidays || holidays.length === 0) {
    return new Set<number>()
  }

  const set = new Set<number>()
  for (const holiday of holidays) {
    const raw = coerceNumber(holiday)
    if (raw === undefined) {
      return valueError()
    }
    set.add(Math.trunc(raw))
  }
  return set
}

function isWeekendWithMask(serial: number, weekendDays: ReadonlySet<number>, dateSystem: ExcelDateSystem): boolean {
  const day = weekendSerialDay(serial, dateSystem)
  return day === undefined || weekendDays.has(day)
}

function offsetWorkday(start: number, offset: number, isWorkday: (serial: number) => boolean): number {
  let cursor = Math.trunc(start)
  const direction = offset >= 0 ? 1 : -1
  let remaining = Math.abs(offset)
  while (remaining > 0) {
    cursor += direction
    if (isWorkday(cursor)) {
      remaining -= 1
    }
  }
  return cursor
}

function createWorkdayBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2) {
      return valueError()
    }

    const start = truncArg(args[0]!)
    const offset = truncArg(args[1]!)
    if (typeof start !== 'number') {
      return start
    }
    if (typeof offset !== 'number') {
      return offset
    }

    const holidays = normalizeHolidayDateSet(args.slice(2))
    if (isErrorValue(holidays)) {
      return holidays
    }

    const isWorkday = (value: number): boolean => !isWeekendSerial(value, dateSystem) && !holidays.has(Math.trunc(value))
    return numberResult(offsetWorkday(start, offset, isWorkday))
  }
}

function createNetworkdaysBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2) {
      return valueError()
    }

    const start = truncArg(args[0]!)
    const end = truncArg(args[1]!)
    if (typeof start !== 'number') {
      return start
    }
    if (typeof end !== 'number') {
      return end
    }

    const holidays = normalizeHolidayDateSet(args.slice(2))
    if (isErrorValue(holidays)) {
      return holidays
    }

    const isWorkday = (value: number): boolean => !isWeekendSerial(value, dateSystem) && !holidays.has(Math.trunc(value))
    const step = start <= end ? 1 : -1
    let count = 0
    for (let cursor = Math.trunc(start); ; cursor += step) {
      if (isWorkday(cursor)) {
        count += step
      }
      if (cursor === Math.trunc(end)) {
        break
      }
    }
    return numberResult(count)
  }
}

function createWorkdayIntlBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2) {
      return valueError()
    }

    const start = truncArg(args[0]!)
    const offset = truncArg(args[1]!)
    if (typeof start !== 'number') {
      return start
    }
    if (typeof offset !== 'number') {
      return offset
    }

    const weekendDays = normalizeWeekendMask(args[2])
    if (isErrorValue(weekendDays)) {
      return weekendDays
    }

    const holidays = normalizeHolidayDateSet(args.length <= 3 ? undefined : args.slice(3))
    if (isErrorValue(holidays)) {
      return holidays
    }

    const isWorkday = (value: number): boolean => !isWeekendWithMask(value, weekendDays, dateSystem) && !holidays.has(Math.trunc(value))
    return numberResult(offsetWorkday(start, offset, isWorkday))
  }
}

function createNetworkdaysIntlBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length < 2) {
      return valueError()
    }

    const start = truncArg(args[0]!)
    const end = truncArg(args[1]!)
    if (typeof start !== 'number') {
      return start
    }
    if (typeof end !== 'number') {
      return end
    }

    const weekendDays = normalizeWeekendMask(args[2])
    if (isErrorValue(weekendDays)) {
      return weekendDays
    }

    const holidays = normalizeHolidayDateSet(args.length <= 3 ? undefined : args.slice(3))
    if (isErrorValue(holidays)) {
      return holidays
    }

    const isWorkday = (value: number): boolean => !isWeekendWithMask(value, weekendDays, dateSystem) && !holidays.has(Math.trunc(value))
    const step = start <= end ? 1 : -1
    let count = 0
    for (let cursor = Math.trunc(start); ; cursor += step) {
      if (isWorkday(cursor)) {
        count += step
      }
      if (cursor === Math.trunc(end)) {
        break
      }
    }
    return numberResult(count)
  }
}

export function createTodayBuiltin(now: DateTimeProvider = () => new Date(), dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length > 0) {
      return valueError()
    }
    return numberResult(Math.floor(utcDateToExcelSerial(now(), dateSystem)))
  }
}

export function createNowBuiltin(now: DateTimeProvider = () => new Date(), dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length > 0) {
      return valueError()
    }
    return numberResult(utcDateToExcelSerial(now(), dateSystem))
  }
}

export function createRandBuiltin(random: RandomProvider = () => Math.random()): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length > 0) {
      return valueError()
    }

    const next = random()
    if (!Number.isFinite(next)) {
      return valueError()
    }

    const bounded = Math.min(Math.max(next, 0), 1 - Number.EPSILON)
    return numberResult(bounded)
  }
}

export function createEdateBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (startDate, months) => {
    const error = firstError([startDate, months])
    if (error) {
      return error
    }

    const startSerial = coerceNumber(startDate)
    const monthOffset = truncArg(months)
    if (startSerial === undefined) {
      return valueError()
    }
    if (typeof monthOffset !== 'number') {
      return monthOffset
    }

    const serial = addMonthsToExcelDate(startSerial, monthOffset, dateSystem)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

export function createEomonthBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (startDate, months) => {
    const error = firstError([startDate, months])
    if (error) {
      return error
    }

    const startSerial = coerceNumber(startDate)
    const monthOffset = truncArg(months)
    if (startSerial === undefined) {
      return valueError()
    }
    if (typeof monthOffset !== 'number') {
      return monthOffset
    }

    const serial = endOfMonthExcelDate(startSerial, monthOffset, dateSystem)
    return serial === undefined ? valueError() : numberResult(serial)
  }
}

export function createDatedifBuiltin(dateSystem: ExcelDateSystem = '1900'): Builtin {
  return (...args) => {
    const error = firstError(args)
    if (error) {
      return error
    }
    if (args.length !== 3) {
      return valueError()
    }
    const startSerial = truncArg(args[0]!)
    const endSerial = truncArg(args[1]!)
    const unit = coerceText(args[2]!)?.trim().toUpperCase()
    if (typeof startSerial !== 'number' || typeof endSerial !== 'number' || !unit) {
      return valueError()
    }
    const value = datedifValue(startSerial, endSerial, unit, dateSystem)
    return value === undefined ? valueError() : numberResult(value)
  }
}

const datetimePlaceholderBuiltins = createBlockedBuiltinMap(datetimePlaceholderBuiltinNames)

export function createDateTimeBuiltins(dateSystem: ExcelDateSystem = '1900'): Record<string, Builtin> {
  return {
    DATE: createDateBuiltin(dateSystem),
    DATEVALUE: createDateValueBuiltin(dateSystem),
    YEAR: createDatePartBuiltin('year', dateSystem),
    MONTH: createDatePartBuiltin('month', dateSystem),
    DAY: createDatePartBuiltin('day', dateSystem),
    TIME: createTimeBuiltin(),
    HOUR: createTimePartBuiltin('hour'),
    MINUTE: createTimePartBuiltin('minute'),
    SECOND: createTimePartBuiltin('second'),
    WEEKDAY: createWeekdayBuiltin(dateSystem),
    DAYS: createDaysBuiltin(),
    WEEKNUM: createWeeknumBuiltin(dateSystem),
    DAYS360: createDays360Builtin(dateSystem),
    ISOWEEKNUM: createIsoWeeknumBuiltin(dateSystem),
    TIMEVALUE: createTimeValueBuiltin(),
    YEARFRAC: createYearfracBuiltin(dateSystem),
    WORKDAY: createWorkdayBuiltin(dateSystem),
    'WORKDAY.INTL': createWorkdayIntlBuiltin(dateSystem),
    NETWORKDAYS: createNetworkdaysBuiltin(dateSystem),
    'NETWORKDAYS.INTL': createNetworkdaysIntlBuiltin(dateSystem),
    TODAY: createTodayBuiltin(() => new Date(), dateSystem),
    NOW: createNowBuiltin(() => new Date(), dateSystem),
    RAND: createRandBuiltin(),
    EDATE: createEdateBuiltin(dateSystem),
    EOMONTH: createEomonthBuiltin(dateSystem),
    DATEDIF: createDatedifBuiltin(dateSystem),
    ...datetimePlaceholderBuiltins,
  }
}

export const datetimeBuiltins: Record<string, Builtin> = createDateTimeBuiltins()
