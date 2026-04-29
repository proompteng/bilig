import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import {
  addMonthsToExcelDate,
  createNowBuiltin,
  createRandBuiltin,
  createTodayBuiltin,
  datetimeBuiltins,
  endOfMonthExcelDate,
  excelDatePartsToSerial,
  excelSerialToDateParts,
  utcDateToExcelSerial,
} from '../builtins/datetime.js'
import { excelDateTimeFixtureSuite } from '../../../excel-fixtures/src/datetime-fixtures.js'

describe('datetime builtins', () => {
  it('converts between Excel serials and date parts in the 1900 system', () => {
    expect(excelDatePartsToSerial(1900, 1, 1)).toBe(1)
    expect(excelDatePartsToSerial(1900, 2, 29)).toBe(60)
    expect(excelDatePartsToSerial(1900, 3, 1)).toBe(61)
    expect(excelDatePartsToSerial(2024, 2, 29)).toBe(45351)

    expect(excelSerialToDateParts(60)).toEqual({ year: 1900, month: 2, day: 29 })
    expect(excelSerialToDateParts(61)).toEqual({ year: 1900, month: 3, day: 1 })
    expect(excelSerialToDateParts(45351)).toEqual({ year: 2024, month: 2, day: 29 })
  })

  it('supports DATE with Excel-style year and month/day normalization', () => {
    expect(
      datetimeBuiltins.DATE({ tag: ValueTag.Number, value: 2024 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 29 }),
    ).toEqual({ tag: ValueTag.Number, value: 45351 })

    expect(
      datetimeBuiltins.DATE({ tag: ValueTag.Number, value: 24 }, { tag: ValueTag.Number, value: 14 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({ tag: ValueTag.Number, value: 9164 })

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.String, value: '2024', stringId: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 29 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 29 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })

  it('extracts YEAR, MONTH, and DAY from serial inputs including the leap-year bug date', () => {
    expect(datetimeBuiltins.YEAR({ tag: ValueTag.Number, value: 45351 })).toEqual({
      tag: ValueTag.Number,
      value: 2024,
    })
    expect(datetimeBuiltins.MONTH({ tag: ValueTag.Number, value: 45351.75 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(datetimeBuiltins.DAY({ tag: ValueTag.Number, value: 60 })).toEqual({
      tag: ValueTag.Number,
      value: 29,
    })

    expect(datetimeBuiltins.YEAR({ tag: ValueTag.String, value: '45351', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.YEAR({ tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Number,
      value: 1899,
    })
    expect(datetimeBuiltins.MONTH({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
  })

  it('supports TIME plus HOUR, MINUTE, SECOND, and WEEKDAY extraction', () => {
    const sundaySerial = excelDatePartsToSerial(2026, 3, 15)!

    expect(
      datetimeBuiltins.TIME({ tag: ValueTag.Number, value: 12 }, { tag: ValueTag.Number, value: 30 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({ tag: ValueTag.Number, value: 0.5208333333333334 })

    expect(datetimeBuiltins.HOUR({ tag: ValueTag.Number, value: 0.5208333333333334 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(datetimeBuiltins.MINUTE({ tag: ValueTag.Number, value: 0.5208333333333334 })).toEqual({
      tag: ValueTag.Number,
      value: 30,
    })
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.Number, value: 0.5208449074074074 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: sundaySerial })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: sundaySerial }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: sundaySerial }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
  })

  it('returns #VALUE for unsupported time-part coercions and weekday return types', () => {
    expect(
      datetimeBuiltins.TIME({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 30 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.HOUR({ tag: ValueTag.String, value: '12:30', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 99 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    const sundaySerial = excelDatePartsToSerial(2026, 3, 15)!
    const weekdayTypes = [1, 2, 11, 12, 13, 14, 15, 16, 17]
    const expectedForSunday = [1, 7, 7, 6, 5, 4, 3, 2, 1]
    weekdayTypes.forEach((type, i) => {
      expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: sundaySerial }, { tag: ValueTag.Number, value: type })).toEqual({
        tag: ValueTag.Number,
        value: expectedForSunday[i],
      })
    })
  })

  it('supports DAYS, WEEKNUM, WORKDAY, and NETWORKDAYS', () => {
    const fridaySerial = excelDatePartsToSerial(2026, 3, 13)!
    const mondayHoliday = excelDatePartsToSerial(2026, 3, 16)!
    const fridayNextWeek = excelDatePartsToSerial(2026, 3, 20)!

    expect(datetimeBuiltins.DAYS({ tag: ValueTag.Number, value: fridayNextWeek }, { tag: ValueTag.Number, value: fridaySerial })).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })

    expect(
      datetimeBuiltins.WEEKNUM({
        tag: ValueTag.Number,
        value: excelDatePartsToSerial(2026, 3, 15)!,
      }),
    ).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({ tag: ValueTag.Number, value: 11 })

    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: mondayHoliday,
    })
    expect(
      datetimeBuiltins.WORKDAY(
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: mondayHoliday },
      ),
    ).toEqual({ tag: ValueTag.Number, value: mondayHoliday + 1 })

    expect(
      datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: fridayNextWeek }),
    ).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(
      datetimeBuiltins.NETWORKDAYS(
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: fridayNextWeek },
        { tag: ValueTag.Number, value: mondayHoliday },
        { tag: ValueTag.Number, value: fridayNextWeek },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
  })

  it('supports WORKDAY.INTL and NETWORKDAYS.INTL weekend masks', () => {
    const fridaySerial = excelDatePartsToSerial(2026, 3, 13)!
    const sundaySerial = excelDatePartsToSerial(2026, 3, 15)!
    const mondaySerial = excelDatePartsToSerial(2026, 3, 16)!
    const tuesdaySerial = excelDatePartsToSerial(2026, 3, 17)!

    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: mondaySerial,
    })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 7 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: mondaySerial })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: '0000011', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: mondaySerial })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: sundaySerial },
      ),
    ).toEqual({ tag: ValueTag.Number, value: tuesdaySerial + 1 })

    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: tuesdaySerial }),
    ).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.Number, value: 7 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.String, value: '1000001', stringId: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: sundaySerial },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })

    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 99 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.String, value: '1111111', stringId: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('supports DATEDIF units', () => {
    const start = excelDatePartsToSerial(2020, 1, 15)!
    const end = excelDatePartsToSerial(2021, 3, 20)!

    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'D', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: end - start })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'M', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'Y', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'YM', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'YD', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 64 })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'MD', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 5 })

    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.String, value: 'D', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'BAD', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('creates deterministic TODAY and NOW builtins from injected UTC dates', () => {
    const fixedNow = new Date('2026-03-19T15:45:30.000Z')
    const TODAY = createTodayBuiltin(() => fixedNow)
    const NOW = createNowBuiltin(() => fixedNow)

    expect(TODAY()).toEqual({ tag: ValueTag.Number, value: 46100 })
    expect(NOW()).toEqual({ tag: ValueTag.Number, value: 46100.65659722222 })
    expect(utcDateToExcelSerial(fixedNow)).toBe(46100.65659722222)

    expect(TODAY({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NOW({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
  })

  it('supports RAND with Excel-style numeric bounds and injectable randomness', () => {
    const RAND = createRandBuiltin(() => 0.625)
    const highRAND = createRandBuiltin(() => 2)
    const lowRAND = createRandBuiltin(() => -0.5)
    const invalidRAND = createRandBuiltin(() => Number.NaN)

    expect(RAND()).toEqual({ tag: ValueTag.Number, value: 0.625 })
    expect(highRAND()).toEqual({ tag: ValueTag.Number, value: 1 - Number.EPSILON })
    expect(lowRAND()).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(invalidRAND()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.RAND({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('returns explicit errors for WORKDAY, NETWORKDAYS, TODAY, NOW, RAND, EDATE, and EOMONTH edge inputs', () => {
    expect(datetimeBuiltins.WORKDAY()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: 46094 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.WORKDAY(
        { tag: ValueTag.Number, value: 46094 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins.NETWORKDAYS()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 46095 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: 46094 }, { tag: ValueTag.String, value: 'bad', stringId: 1 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins.NETWORKDAYS(
        { tag: ValueTag.Number, value: 46094 },
        { tag: ValueTag.Number, value: 46095 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(
      createTodayBuiltin(() => new Date('2026-03-19T00:00:00.000Z'))({
        tag: ValueTag.Error,
        code: ErrorCode.Name,
      }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name })
    expect(
      createNowBuiltin(() => new Date('2026-03-19T00:00:00.000Z'))({
        tag: ValueTag.Number,
        value: 1,
      }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(createRandBuiltin(() => 0.5)({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    expect(datetimeBuiltins.EDATE({ tag: ValueTag.Number, value: 45322 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.Number, value: 45322 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers DAYS and WEEKNUM validation and alternate return-type branches', () => {
    const sampleDate = excelDatePartsToSerial(2026, 3, 15)!

    expect(datetimeBuiltins.DAYS({ tag: ValueTag.Number, value: 10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DAYS({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DAYS({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.String, value: 'bad', stringId: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 12 })).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 13 })).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 14 })).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 15 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.String, value: 'bad', stringId: 3 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.String, value: 'bad', stringId: 4 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: Number.NaN }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports EDATE month shifting with end-of-month clamping', () => {
    expect(addMonthsToExcelDate(45322, 1)).toBe(45351)
    expect(addMonthsToExcelDate(45351, -1)).toBe(45320)

    expect(datetimeBuiltins.EDATE({ tag: ValueTag.Number, value: 45322 }, { tag: ValueTag.Number, value: 1.9 })).toEqual({
      tag: ValueTag.Number,
      value: 45351,
    })

    expect(datetimeBuiltins.EDATE({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports DATEVALUE for serial numbers and ISO date strings', () => {
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Number, value: 1.2 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.String, value: '2024-02-29', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: excelDatePartsToSerial(2024, 2, 29)!,
    })
    expect(datetimeBuiltins.DATEVALUE()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.String, value: 'not-a-date', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports DAYS360 and YEARFRAC across basis modes', () => {
    const jan1 = excelDatePartsToSerial(2024, 1, 1)!
    const jul1 = excelDatePartsToSerial(2024, 7, 1)!
    const feb28 = excelDatePartsToSerial(2023, 2, 28)!
    const mar31 = excelDatePartsToSerial(2023, 3, 31)!

    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: feb28 }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Number,
      value: 33,
    })
    expect(
      datetimeBuiltins.DAYS360(
        { tag: ValueTag.Number, value: feb28 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 32 })
    expect(
      datetimeBuiltins.DAYS360(
        { tag: ValueTag.Number, value: feb28 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins.YEARFRAC({ tag: ValueTag.Number, value: jan1 }, { tag: ValueTag.Number, value: jul1 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5,
    })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 182 / 366 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 182 / 360 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 182 / 365 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0.5 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 182 / 366 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 9 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('supports EOMONTH end-of-month lookups', () => {
    expect(endOfMonthExcelDate(45337, 0)).toBe(45351)
    expect(endOfMonthExcelDate(45337, 1)).toBe(45382)

    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.Number, value: 45337 }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 45382,
    })

    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.Number, value: 45337 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
  })

  it('supports TIMEVALUE parsing and extended week-number variants', () => {
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12:30 PM', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5208333333333334,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12:00 AM', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '1:02:03 pm', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: (13 * 3600 + 2 * 60 + 3) / 86_400,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '24:00', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '13:00 PM', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12:60', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIMEVALUE()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(
      datetimeBuiltins.ISOWEEKNUM({
        tag: ValueTag.Number,
        value: excelDatePartsToSerial(2026, 1, 1)!,
      }),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 12 }),
    ).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 16 }),
    ).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 21 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('handles reverse ranges, weekend starts, and invalid helper inputs', () => {
    const fridaySerial = excelDatePartsToSerial(2026, 3, 13)!
    const saturdaySerial = excelDatePartsToSerial(2026, 3, 14)!
    const fridayNextWeek = excelDatePartsToSerial(2026, 3, 20)!

    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: saturdaySerial }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: excelDatePartsToSerial(2026, 3, 16)!,
    })
    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Number,
      value: excelDatePartsToSerial(2026, 3, 12)!,
    })
    expect(
      datetimeBuiltins.WORKDAY(
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(
      datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: fridayNextWeek }, { tag: ValueTag.Number, value: fridaySerial }),
    ).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(
      datetimeBuiltins.NETWORKDAYS(
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: fridayNextWeek },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(excelSerialToDateParts(Number.NaN)).toBeUndefined()
    expect(excelDatePartsToSerial(10_000, 1, 1)).toBeUndefined()
    expect(addMonthsToExcelDate(Number.NaN, 1)).toBeUndefined()
    expect(endOfMonthExcelDate(Number.NaN, 1)).toBeUndefined()
  })

  it('covers remaining datetime coercion and calendar validation branches', () => {
    const jan31 = excelDatePartsToSerial(2024, 1, 31)!
    const feb28 = excelDatePartsToSerial(2023, 2, 28)!
    const feb29 = excelDatePartsToSerial(2024, 2, 29)!
    const mar31 = excelDatePartsToSerial(2024, 3, 31)!

    expect(excelDatePartsToSerial(Number.NaN, 1, 1)).toBeUndefined()
    expect(excelDatePartsToSerial(-1, 1, 1)).toBeUndefined()
    expect(excelSerialToDateParts(Number.POSITIVE_INFINITY)).toBeUndefined()
    expect(addMonthsToExcelDate(jan31, Number.NaN)).toBeUndefined()
    expect(endOfMonthExcelDate(jan31, Number.NaN)).toBeUndefined()

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: Number.NaN },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.YEAR({ tag: ValueTag.Number, value: Number.NaN })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.MONTH({ tag: ValueTag.Number, value: Number.NaN })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.DAY({ tag: ValueTag.Number, value: Number.NaN })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.Number, value: 33_000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.TIME({ tag: ValueTag.Boolean, value: true }, { tag: ValueTag.Boolean, value: false }, { tag: ValueTag.Empty }),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1 / 24,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.Number, value: 0.5 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.Boolean, value: true })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: 'bad:00', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '-1:00', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '23:59:60', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: Number.NaN }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Number,
      value: 60,
    })

    const leapSpan = datetimeBuiltins.YEARFRAC(
      { tag: ValueTag.Number, value: feb28 },
      { tag: ValueTag.Number, value: feb29 },
      { tag: ValueTag.Number, value: 1 },
    )
    expect(leapSpan.tag).toBe(ValueTag.Number)
    if (leapSpan.tag === ValueTag.Number) {
      expect(leapSpan.value).toBeCloseTo(366 / 365.5, 12)
    }
    const multiYearSpan = datetimeBuiltins.YEARFRAC(
      { tag: ValueTag.Number, value: excelDatePartsToSerial(2020, 1, 1)! },
      { tag: ValueTag.Number, value: excelDatePartsToSerial(2023, 7, 1)! },
      { tag: ValueTag.Number, value: 1 },
    )
    expect(multiYearSpan.tag).toBe(ValueTag.Number)
    if (multiYearSpan.tag === ValueTag.Number) {
      expect(multiYearSpan.value).toBeCloseTo(1277 / 365.25, 12)
    }

    expect(datetimeBuiltins.ISOWEEKNUM()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.ISOWEEKNUM({ tag: ValueTag.Number, value: Number.NaN })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.String, value: '000000', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers additional datetime edge validation paths', () => {
    const jan31 = excelDatePartsToSerial(2024, 1, 31)!
    const mar31 = excelDatePartsToSerial(2024, 3, 31)!

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: 2024 },
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.HOUR()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.MINUTE({ tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.WEEKDAY()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: jan31 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.ISOWEEKNUM({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.YEARFRAC({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.YEARFRAC({ tag: ValueTag.Number, value: jan31 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.YEARFRAC({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Number, value: Number.NaN })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: jan31 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: mar31 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: jan31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DATEDIF({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.String, value: 'D', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
  })

  it('ships a focused datetime fixture suite for later aggregation', () => {
    expect(excelDateTimeFixtureSuite.id).toBe('datetime-serial-1900')
    expect(excelDateTimeFixtureSuite.sheets).toEqual([{ name: 'Sheet1' }])
    expect(excelDateTimeFixtureSuite.cases).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: 'date-time:serial-addition',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 45299 } }],
        }),
        expect.objectContaining({
          id: 'date-time:date-empty-year-coercion',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 1 } }],
        }),
        expect.objectContaining({
          id: 'date-time:date-text-error',
          outputs: [
            {
              address: 'A2',
              expected: { kind: 'error', code: ErrorCode.Value, display: '#VALUE!' },
            },
          ],
        }),
        expect.objectContaining({
          id: 'date-time:date-constructor-leap-day',
          formula: '=DATE(2024,2,29)',
        }),
        expect.objectContaining({
          id: 'date-time:year-boolean-coercion',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 1900 } }],
        }),
        expect.objectContaining({
          id: 'date-time:month-text-error',
          outputs: [
            {
              address: 'A2',
              expected: { kind: 'error', code: ErrorCode.Value, display: '#VALUE!' },
            },
          ],
        }),
        expect.objectContaining({
          id: 'date-time:day-leap-bug-serial',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 29 } }],
        }),
        expect.objectContaining({
          id: 'date-time:time-basic',
          outputs: [{ address: 'A1', expected: { kind: 'number', value: 0.5208333333333334 } }],
        }),
        expect.objectContaining({
          id: 'date-time:hour-basic',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 12 } }],
        }),
        expect.objectContaining({
          id: 'date-time:minute-basic',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 30 } }],
        }),
        expect.objectContaining({
          id: 'date-time:second-basic',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 1 } }],
        }),
        expect.objectContaining({
          id: 'date-time:weekday-basic',
          outputs: [{ address: 'A1', expected: { kind: 'number', value: 1 } }],
        }),
        expect.objectContaining({
          id: 'date-time:edate-month-shift',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 45351 } }],
        }),
        expect.objectContaining({
          id: 'date-time:edate-boolean-coercion',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 32 } }],
        }),
        expect.objectContaining({
          id: 'date-time:edate-text-error',
          outputs: [
            {
              address: 'A2',
              expected: { kind: 'error', code: ErrorCode.Value, display: '#VALUE!' },
            },
          ],
        }),
        expect.objectContaining({
          id: 'date-time:eomonth-boolean-coercion',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 60 } }],
        }),
        expect.objectContaining({
          id: 'date-time:eomonth-text-error',
          outputs: [
            {
              address: 'A3',
              expected: { kind: 'error', code: ErrorCode.Value, display: '#VALUE!' },
            },
          ],
        }),
        expect.objectContaining({
          id: 'volatile:today-captured-utc',
          outputs: [{ address: 'A1', expected: { kind: 'number', value: 46100 } }],
        }),
        expect.objectContaining({
          id: 'volatile:rand-captured',
          outputs: [{ address: 'A1', expected: { kind: 'number', value: 0.625 } }],
        }),
      ]),
    )
  })

  it('covers yearFracByBasis complex branches and edge cases', () => {
    const d1 = excelDatePartsToSerial(2023, 1, 1)!
    const d2 = excelDatePartsToSerial(2024, 1, 1)!
    const d3 = excelDatePartsToSerial(2025, 1, 1)!

    // Basis 1: Actual/actual
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: d1 },
        { tag: ValueTag.Number, value: d2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 365 / 365 })

    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: d2 },
        { tag: ValueTag.Number, value: d3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 366 / 366 })

    // Multi-year crossing
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: d1 },
        { tag: ValueTag.Number, value: d3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2.0009124087591244 })
  })

  it('covers more TIMEVALUE parsing scenarios', () => {
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '  12:30 PM  ', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5208333333333334,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12:00:00', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '11:59:59 PM', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: (23 * 3600 + 59 * 60 + 59) / 86400,
    })
  })

  it('covers date boundary limits and normalization', () => {
    expect(excelDatePartsToSerial(-1, 1, 1)).toBeUndefined()
    expect(excelDatePartsToSerial(10000, 1, 1)).toBeUndefined()
    expect(addMonthsToExcelDate(1, 120000)).toBeUndefined()
    expect(endOfMonthExcelDate(1, 120000)).toBeUndefined()
  })

  it('covers remaining datetime public validation and edge branches', () => {
    const jan31 = excelDatePartsToSerial(2024, 1, 31)!
    const feb29 = excelDatePartsToSerial(2024, 2, 29)!
    const mar31 = excelDatePartsToSerial(2024, 3, 31)!

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: 2024 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: 2024 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATE({ tag: ValueTag.Number, value: 10000 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DATE()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Boolean, value: true })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Empty })).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Number, value: Number.POSITIVE_INFINITY })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.String, value: '   ', stringId: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.YEAR({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.MONTH({ tag: ValueTag.String, value: 'bad', stringId: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DAY({ tag: ValueTag.Number, value: Number.NEGATIVE_INFINITY })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.Error, code: ErrorCode.NA },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.String, value: 'bad', stringId: 5 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.String, value: 'bad', stringId: 6 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.TIME({ tag: ValueTag.Number, value: 32768 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIME()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins.HOUR()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.MINUTE({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.Number, value: -0.25 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(datetimeBuiltins.DAYS({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: jan31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Number, value: 17 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })

    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Error, code: ErrorCode.Ref },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: mar31 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Number, value: 11 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 53,
    })

    expect(datetimeBuiltins.EDATE({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.String, value: 'bad', stringId: 7 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.String, value: 'D', stringId: 8 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: feb29 },
        { tag: ValueTag.String, value: 'YD', stringId: 9 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 29,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: feb29 },
        { tag: ValueTag.String, value: 'MD', stringId: 10 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 29,
    })
  })
})
