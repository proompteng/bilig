import * as fc from 'fast-check'
import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import { datetimeBuiltins, excelDatePartsToSerial, excelSerialToDateParts } from '../builtins/datetime.js'

const validDatePartsArbitrary = fc
  .record({
    year: fc.integer({ min: 1900, max: 2099 }),
    month: fc.integer({ min: 1, max: 12 }),
  })
  .chain(({ year, month }) =>
    fc.record({
      year: fc.constant(year),
      month: fc.constant(month),
      day: fc.integer({ min: 1, max: daysInMonth(year, month) }),
    }),
  )

const timeTextCaseArbitrary = fc.oneof(
  fc
    .record({
      hour: fc.integer({ min: 0, max: 23 }),
      minute: fc.integer({ min: 0, max: 59 }),
      second: fc.integer({ min: 0, max: 59 }),
    })
    .map(({ hour, minute, second }) => ({
      text: `${hour}:${pad2(minute)}:${pad2(second)}`,
      secondsOfDay: hour * 3600 + minute * 60 + second,
    })),
  fc
    .record({
      hour: fc.integer({ min: 1, max: 12 }),
      minute: fc.integer({ min: 0, max: 59 }),
      second: fc.integer({ min: 0, max: 59 }),
      meridiem: fc.constantFrom('AM', 'PM'),
    })
    .map(({ hour, minute, second, meridiem }) => {
      const normalizedHour = hour === 12 ? (meridiem === 'AM' ? 0 : 12) : meridiem === 'PM' ? hour + 12 : hour
      return {
        text: `${hour}:${pad2(minute)}:${pad2(second)} ${meridiem}`,
        secondsOfDay: normalizedHour * 3600 + minute * 60 + second,
      }
    }),
)

describe('formula datetime fuzz', () => {
  it('round-trips valid Gregorian dates through Excel serials and extractors', async () => {
    await runProperty({
      suite: 'formula/datetime/date-serial-roundtrip',
      arbitrary: validDatePartsArbitrary,
      predicate: ({ year, month, day }) => {
        const serial = excelDatePartsToSerial(year, month, day)

        expect(serial).toBeDefined()
        expect(excelSerialToDateParts(serial!)).toEqual({ year, month, day })
        expect(datetimeBuiltins.DATE(num(year), num(month), num(day))).toEqual(num(serial!))
        expect(datetimeBuiltins.YEAR(num(serial!))).toEqual(num(year))
        expect(datetimeBuiltins.MONTH(num(serial!))).toEqual(num(month))
        expect(datetimeBuiltins.DAY(num(serial!))).toEqual(num(day))
      },
    })
  })

  it('keeps TIMEVALUE parsing aligned with TIME fractions for generated clock text', async () => {
    await runProperty({
      suite: 'formula/datetime/timevalue-clock-fraction',
      arbitrary: timeTextCaseArbitrary,
      predicate: ({ text: timeText, secondsOfDay }) => {
        const actual = datetimeBuiltins.TIMEVALUE(string(timeText))
        const expected = secondsOfDay / 86_400

        expectNumberClose(actual, expected)
        expect(
          datetimeBuiltins.TIME(num(Math.floor(secondsOfDay / 3600)), num(Math.floor((secondsOfDay % 3600) / 60)), num(secondsOfDay % 60)),
        ).toEqual(num(expected))
      },
    })
  })
})

function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate()
}

function pad2(value: number): string {
  return `${value}`.padStart(2, '0')
}

function expectNumberClose(value: CellValue, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    return
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

function num(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function string(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}
