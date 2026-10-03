import { describe, expect, it } from 'vitest'
import { numberToWorkPaperDate, numberToWorkPaperDateTime, numberToWorkPaperTime } from '../work-paper-date-time.js'

describe('work paper date/time helpers', () => {
  it('converts Excel serials to date-time parts', () => {
    expect(numberToWorkPaperDateTime(2.5)).toEqual({
      year: 1900,
      month: 1,
      day: 2,
      hours: 12,
      minutes: 0,
      seconds: 0,
    })
  })

  it('splits date and time projections', () => {
    expect(numberToWorkPaperDate(2.5)).toEqual({ year: 1900, month: 1, day: 2 })
    expect(numberToWorkPaperTime(2.5)).toEqual({ hours: 12, minutes: 0, seconds: 0 })
  })

  it('returns undefined for invalid serials', () => {
    expect(numberToWorkPaperDateTime(Number.NaN)).toBeUndefined()
    expect(numberToWorkPaperDate(Number.POSITIVE_INFINITY)).toBeUndefined()
    expect(numberToWorkPaperTime(Number.NEGATIVE_INFINITY)).toBeUndefined()
  })

  it('rounds fractional days to seconds', () => {
    expect(numberToWorkPaperTime(1 + 1 / 86_400)).toEqual({ hours: 0, minutes: 0, seconds: 1 })
  })
})
