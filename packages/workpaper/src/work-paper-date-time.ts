import { excelSerialToDateParts } from '@bilig/formula'
import type { WorkPaperDateTime } from './work-paper-types.js'

export function numberToWorkPaperDateTime(value: number): WorkPaperDateTime | undefined {
  const dateParts = excelSerialToDateParts(value)
  if (!dateParts) {
    return undefined
  }
  const whole = Math.floor(value)
  const fraction = value - whole
  const totalSeconds = Math.round(Math.max(0, fraction) * 86_400)
  const hours = Math.floor(totalSeconds / 3_600) % 24
  const minutes = Math.floor((totalSeconds % 3_600) / 60)
  const seconds = totalSeconds % 60
  return {
    year: dateParts.year,
    month: dateParts.month,
    day: dateParts.day,
    hours,
    minutes,
    seconds,
  }
}

export function numberToWorkPaperDate(value: number): Omit<WorkPaperDateTime, 'hours' | 'minutes' | 'seconds'> | undefined {
  const dateTime = numberToWorkPaperDateTime(value)
  if (!dateTime) {
    return undefined
  }
  const { year, month, day } = dateTime
  return { year, month, day }
}

export function numberToWorkPaperTime(value: number): Pick<WorkPaperDateTime, 'hours' | 'minutes' | 'seconds'> | undefined {
  const dateTime = numberToWorkPaperDateTime(value)
  if (!dateTime) {
    return undefined
  }
  const { hours, minutes, seconds } = dateTime
  return { hours, minutes, seconds }
}
