import { ErrorCode } from './protocol'
import {
  excelDayPartFromSerial,
  excelMonthPartFromSerial,
  excelSecondOfDay,
  excelSerialWhole,
  excelYearPartFromSerial,
} from './date-finance'
import { parseNumericText } from './text-special'

function roundToDigits(value: f64, digits: i32): f64 {
  if (digits >= 0) {
    const factor = Math.pow(10.0, <f64>digits)
    return Math.round(value * factor) / factor
  }
  const factor = Math.pow(10.0, <f64>-digits)
  return Math.round(value / factor) * factor
}

function formatThousandsText(integerPart: string): string {
  let output = ''
  let groupCount = 0
  for (let index = integerPart.length - 1; index >= 0; index--) {
    output = integerPart.charAt(index) + output
    groupCount += 1
    if (index > 0 && groupCount % 3 == 0) {
      output = ',' + output
    }
  }
  return output
}

function zeroPadIntegerText(value: i64, width: i32): string {
  let output = value.toString()
  while (output.length < width) {
    output = '0' + output
  }
  return output
}

export function formatFixedText(value: f64, decimals: i32, includeThousands: bool): string | null {
  if (!isFinite(value)) {
    return null
  }
  const rounded = roundToDigits(value, decimals)
  const sign = rounded < 0.0 ? '-' : ''
  const unsigned = Math.abs(rounded)
  const fixedDecimals = decimals >= 0 ? decimals : 0
  const scale = Math.pow(10.0, <f64>fixedDecimals)
  let integerPartValue = <i64>Math.floor(unsigned)
  let scaledFraction = fixedDecimals > 0 ? <i64>Math.round((unsigned - <f64>integerPartValue) * scale) : 0
  if (fixedDecimals > 0 && scaledFraction >= <i64>scale) {
    integerPartValue += 1
    scaledFraction -= <i64>scale
  }
  const integerPart = integerPartValue.toString()
  const fractionPart = fixedDecimals > 0 ? zeroPadIntegerText(scaledFraction, fixedDecimals) : ''
  const normalizedInteger = includeThousands ? formatThousandsText(integerPart) : integerPart
  return fractionPart.length > 0 ? `${sign}${normalizedInteger}.${fractionPart}` : `${sign}${normalizedInteger}`
}

export function splitFormatSectionsText(format: string): Array<string> {
  const sections = new Array<string>()
  let current = ''
  let inQuotes = false
  let bracketDepth = 0
  let escaped = false
  for (let index = 0; index < format.length; index += 1) {
    const char = format.charAt(index)
    if (escaped) {
      current += char
      escaped = false
      continue
    }
    if (char == '\\') {
      current += char
      escaped = true
      continue
    }
    if (char == '"') {
      current += char
      inQuotes = !inQuotes
      continue
    }
    if (!inQuotes && char == '[') {
      bracketDepth += 1
      current += char
      continue
    }
    if (!inQuotes && char == ']' && bracketDepth > 0) {
      bracketDepth -= 1
      current += char
      continue
    }
    if (!inQuotes && bracketDepth == 0 && char == ';') {
      sections.push(current)
      current = ''
      continue
    }
    current += char
  }
  sections.push(current)
  return sections
}

export function stripFormatDecorationsText(section: string): string {
  let output = ''
  let inQuotes = false
  for (let index = 0; index < section.length; index += 1) {
    const char = section.charAt(index)
    if (inQuotes) {
      if (char == '"') {
        inQuotes = false
      } else {
        output += char
      }
      continue
    }
    if (char == '"') {
      inQuotes = true
      continue
    }
    if (char == '\\') {
      if (index + 1 < section.length) {
        output += section.charAt(index + 1)
        index += 1
      }
      continue
    }
    if (char == '_') {
      output += ' '
      index += 1
      continue
    }
    if (char == '*') {
      index += 1
      continue
    }
    if (char == '[') {
      const end = section.indexOf(']', index + 1)
      if (end >= 0) {
        index = end
      }
      continue
    }
    output += char
  }
  return output
}

function isFormatPlaceholderCode(code: i32): bool {
  return code == 48 || code == 35 || code == 63
}

function countOccurrences(text: string, needle: string): i32 {
  let count = 0
  for (let index = 0; index < text.length; index += 1) {
    if (text.charAt(index) == needle) {
      count += 1
    }
  }
  return count
}

function countZeroPlaceholders(text: string): i32 {
  let count = 0
  for (let index = 0; index < text.length; index += 1) {
    if (text.charCodeAt(index) == 48) {
      count += 1
    }
  }
  return count
}

function countDigitPlaceholders(text: string): i32 {
  let count = 0
  for (let index = 0; index < text.length; index += 1) {
    if (isFormatPlaceholderCode(text.charCodeAt(index))) {
      count += 1
    }
  }
  return count
}

function zeroPadPositiveText(value: i32, width: i32): string {
  let output = value.toString()
  while (output.length < width) {
    output = '0' + output
  }
  return output
}

function trimOptionalFractionText(fraction: string, minDigits: i32): string {
  let trimmed = fraction
  while (trimmed.length > minDigits && trimmed.endsWith('0')) {
    trimmed = trimmed.slice(0, trimmed.length - 1)
  }
  return trimmed
}

function monthNameText(month: i32, longName: bool): string {
  switch (month) {
    case 1:
      return longName ? 'January' : 'Jan'
    case 2:
      return longName ? 'February' : 'Feb'
    case 3:
      return longName ? 'March' : 'Mar'
    case 4:
      return longName ? 'April' : 'Apr'
    case 5:
      return 'May'
    case 6:
      return longName ? 'June' : 'Jun'
    case 7:
      return longName ? 'July' : 'Jul'
    case 8:
      return longName ? 'August' : 'Aug'
    case 9:
      return longName ? 'September' : 'Sep'
    case 10:
      return longName ? 'October' : 'Oct'
    case 11:
      return longName ? 'November' : 'Nov'
    case 12:
      return longName ? 'December' : 'Dec'
    default:
      return ''
  }
}

function weekdayNameText(index: i32, longName: bool): string {
  switch (index) {
    case 0:
      return longName ? 'Sunday' : 'Sun'
    case 1:
      return longName ? 'Monday' : 'Mon'
    case 2:
      return longName ? 'Tuesday' : 'Tue'
    case 3:
      return longName ? 'Wednesday' : 'Wed'
    case 4:
      return longName ? 'Thursday' : 'Thu'
    case 5:
      return longName ? 'Friday' : 'Fri'
    case 6:
      return longName ? 'Saturday' : 'Sat'
    default:
      return ''
  }
}

function excelWeekdayIndexFromSerial(tag: u8, value: f64): i32 {
  const whole = excelSerialWhole(tag, value)
  if (whole == i32.MIN_VALUE || whole < 0) {
    return i32.MIN_VALUE
  }
  const adjustedWhole = whole < 60 ? whole : whole - 1
  return ((adjustedWhole % 7) + 7) % 7
}

export function containsDateTimeTokens(text: string): bool {
  const upper = text.toUpperCase()
  if (upper.indexOf('AM/PM') >= 0 || upper.indexOf('A/P') >= 0) {
    return true
  }
  for (let index = 0; index < upper.length; index += 1) {
    const code = upper.charCodeAt(index)
    if (code == 89 || code == 68 || code == 72 || code == 83 || code == 77) {
      return true
    }
  }
  return false
}

export function formatTextSectionText(value: string, section: string): string {
  const cleaned = stripFormatDecorationsText(section)
  let output = ''
  for (let index = 0; index < cleaned.length; index += 1) {
    if (cleaned.charAt(index) == '@') {
      output += value
    } else {
      output += cleaned.charAt(index)
    }
  }
  return output
}

const DATE_TOKEN_NONE: i32 = 0
const DATE_TOKEN_YEAR: i32 = 1
const DATE_TOKEN_MONTH: i32 = 2
const DATE_TOKEN_MINUTE: i32 = 3
const DATE_TOKEN_DAY: i32 = 4
const DATE_TOKEN_HOUR: i32 = 5
const DATE_TOKEN_SECOND: i32 = 6

function nextDateTokenKind(text: string, start: i32): i32 {
  for (let index = start; index < text.length; index += 1) {
    const remainder = text.substring(index).toUpperCase()
    if (remainder.startsWith('AM/PM') || remainder.startsWith('A/P')) {
      return DATE_TOKEN_NONE
    }
    const lower = text.charAt(index).toLowerCase()
    if (lower == 'y') return DATE_TOKEN_YEAR
    if (lower == 'm') return DATE_TOKEN_MONTH
    if (lower == 'd') return DATE_TOKEN_DAY
    if (lower == 'h') return DATE_TOKEN_HOUR
    if (lower == 's') return DATE_TOKEN_SECOND
  }
  return DATE_TOKEN_NONE
}

function formatAmPmText(token: string, hour: i32): string {
  const isPm = hour >= 12
  const upper = token.toUpperCase()
  if (upper == 'A/P') {
    const letter = isPm ? 'P' : 'A'
    return token == token.toLowerCase() ? letter.toLowerCase() : letter
  }
  return token == token.toLowerCase() ? (isPm ? 'pm' : 'am') : isPm ? 'PM' : 'AM'
}

export function formatDateTimePatternText(tag: u8, value: f64, section: string): string | null {
  const cleaned = stripFormatDecorationsText(section)
  const year = excelYearPartFromSerial(tag, value)
  const month = excelMonthPartFromSerial(tag, value)
  const day = excelDayPartFromSerial(tag, value)
  const secondOfDay = excelSecondOfDay(tag, value)
  const weekdayIndex = excelWeekdayIndexFromSerial(tag, value)
  if (
    year == i32.MIN_VALUE ||
    month == i32.MIN_VALUE ||
    day == i32.MIN_VALUE ||
    secondOfDay == i32.MIN_VALUE ||
    weekdayIndex == i32.MIN_VALUE
  ) {
    return null
  }
  const hour24 = secondOfDay / 3600
  const minute = (secondOfDay % 3600) / 60
  const second = secondOfDay % 60
  const hasAmPm = cleaned.toUpperCase().indexOf('AM/PM') >= 0 || cleaned.toUpperCase().indexOf('A/P') >= 0
  let output = ''
  let previousKind = DATE_TOKEN_NONE
  for (let index = 0; index < cleaned.length; ) {
    const remainderUpper = cleaned.substring(index).toUpperCase()
    if (remainderUpper.startsWith('AM/PM')) {
      output += formatAmPmText(cleaned.substring(index, index + 5), hour24)
      index += 5
      previousKind = DATE_TOKEN_NONE
      continue
    }
    if (remainderUpper.startsWith('A/P')) {
      output += formatAmPmText(cleaned.substring(index, index + 3), hour24)
      index += 3
      previousKind = DATE_TOKEN_NONE
      continue
    }
    const lower = cleaned.charAt(index).toLowerCase()
    if (lower == 'y' || lower == 'm' || lower == 'd' || lower == 'h' || lower == 's') {
      let end = index + 1
      while (end < cleaned.length && cleaned.charAt(end).toLowerCase() == lower) {
        end += 1
      }
      const token = cleaned.substring(index, end)
      if (lower == 'y') {
        output += token.length == 2 ? zeroPadPositiveText(year % 100, 2) : year.toString()
        previousKind = DATE_TOKEN_YEAR
      } else if (lower == 'd') {
        output +=
          token.length == 1
            ? day.toString()
            : token.length == 2
              ? zeroPadPositiveText(day, 2)
              : token.length == 3
                ? weekdayNameText(weekdayIndex, false)
                : weekdayNameText(weekdayIndex, true)
        previousKind = DATE_TOKEN_DAY
      } else if (lower == 'h') {
        const hourValue = hasAmPm ? ((hour24 + 11) % 12) + 1 : hour24
        output += token.length >= 2 ? zeroPadPositiveText(hourValue, 2) : hourValue.toString()
        previousKind = DATE_TOKEN_HOUR
      } else if (lower == 's') {
        output += token.length >= 2 ? zeroPadPositiveText(second, 2) : second.toString()
        previousKind = DATE_TOKEN_SECOND
      } else {
        const nextKind = nextDateTokenKind(cleaned, end)
        const minuteToken = previousKind == DATE_TOKEN_HOUR || previousKind == DATE_TOKEN_MINUTE || nextKind == DATE_TOKEN_SECOND
        if (minuteToken) {
          output += token.length >= 2 ? zeroPadPositiveText(minute, 2) : minute.toString()
          previousKind = DATE_TOKEN_MINUTE
        } else {
          output +=
            token.length == 1
              ? month.toString()
              : token.length == 2
                ? zeroPadPositiveText(month, 2)
                : token.length == 3
                  ? monthNameText(month, false)
                  : monthNameText(month, true)
          previousKind = DATE_TOKEN_MONTH
        }
      }
      index = end
      continue
    }
    output += cleaned.charAt(index)
    index += 1
  }
  return output
}

function formatScientificPatternText(value: f64, core: string): string {
  let exponentIndex = -1
  for (let index = 0; index + 1 < core.length; index += 1) {
    const upper = core.charAt(index).toUpperCase()
    const sign = core.charAt(index + 1)
    if (upper == 'E' && (sign == '+' || sign == '-')) {
      exponentIndex = index
      break
    }
  }
  const mantissaPattern = core.slice(0, exponentIndex)
  const exponentPattern = core.slice(exponentIndex + 2)
  const dotIndex = mantissaPattern.indexOf('.')
  const fractionPattern = dotIndex >= 0 ? mantissaPattern.slice(dotIndex + 1) : ''
  const maxFractionDigits = countDigitPlaceholders(fractionPattern)
  const minFractionDigits = countZeroPlaceholders(fractionPattern)
  let exponentValue = value == 0.0 ? 0 : <i32>Math.floor(Math.log(value) / Math.log(10.0))
  let mantissaValue = value == 0.0 ? 0.0 : value / Math.pow(10.0, <f64>exponentValue)
  mantissaValue = roundToDigits(mantissaValue, maxFractionDigits)
  if (mantissaValue >= 10.0) {
    mantissaValue /= 10.0
    exponentValue += 1
  }
  let mantissa = formatFixedText(mantissaValue, maxFractionDigits, false)
  if (mantissa == null) {
    mantissa = '0'
  }
  const mantissaDot = mantissa.indexOf('.')
  if (mantissaDot >= 0) {
    const integerPart = mantissa.slice(0, mantissaDot)
    const trimmedFraction = trimOptionalFractionText(mantissa.slice(mantissaDot + 1), minFractionDigits)
    mantissa = trimmedFraction.length > 0 ? `${integerPart}.${trimmedFraction}` : integerPart
  }
  return `${mantissa}E${exponentValue < 0 ? '-' : '+'}${zeroPadPositiveText(<i32>Math.abs(<f64>exponentValue), exponentPattern.length)}`
}

export function formatNumericPatternText(value: f64, section: string, autoNegative: bool): string {
  const cleaned = stripFormatDecorationsText(section)
  let firstPlaceholder = -1
  let lastPlaceholder = -1
  for (let index = 0; index < cleaned.length; index += 1) {
    if (isFormatPlaceholderCode(cleaned.charCodeAt(index))) {
      if (firstPlaceholder == -1) {
        firstPlaceholder = index
      }
      lastPlaceholder = index
    }
  }
  if (firstPlaceholder == -1) {
    return autoNegative && !cleaned.startsWith('-') ? `-${cleaned}` : cleaned
  }
  const prefix = cleaned.slice(0, firstPlaceholder)
  const core = cleaned.slice(firstPlaceholder, lastPlaceholder + 1)
  const suffix = cleaned.slice(lastPlaceholder + 1)
  const scaledValue = Math.abs(value) * Math.pow(100.0, <f64>countOccurrences(cleaned, '%'))
  let numericText = ''
  let exponentIndex = -1
  for (let index = 0; index + 1 < core.length; index += 1) {
    const upper = core.charAt(index).toUpperCase()
    const sign = core.charAt(index + 1)
    if (upper == 'E' && (sign == '+' || sign == '-')) {
      exponentIndex = index
      break
    }
  }
  if (exponentIndex >= 0) {
    numericText = formatScientificPatternText(scaledValue, core)
  } else {
    const decimalIndex = core.indexOf('.')
    const integerPatternRaw = decimalIndex >= 0 ? core.slice(0, decimalIndex) : core
    const fractionPattern = decimalIndex >= 0 ? core.slice(decimalIndex + 1) : ''
    const integerPattern = integerPatternRaw.replace(',', '')
    const maxFractionDigits = countDigitPlaceholders(fractionPattern)
    const minFractionDigits = countZeroPlaceholders(fractionPattern)
    const minIntegerDigits = countZeroPlaceholders(integerPattern)
    const rounded = roundToDigits(scaledValue, maxFractionDigits)
    let fixed = formatFixedText(rounded, maxFractionDigits, false)
    if (fixed == null) {
      fixed = '0'
    }
    const fixedDot = fixed.indexOf('.')
    let integerPart: string = fixedDot >= 0 ? fixed.slice(0, fixedDot) : fixed
    let fractionPart: string = fixedDot >= 0 ? fixed.slice(fixedDot + 1) : ''
    while (integerPart.length < minIntegerDigits) {
      integerPart = '0' + integerPart
    }
    if (integerPatternRaw.indexOf(',') >= 0) {
      integerPart = formatThousandsText(integerPart)
    }
    fractionPart = trimOptionalFractionText(fractionPart, minFractionDigits)
    numericText = fractionPart.length > 0 ? `${integerPart}.${fractionPart}` : integerPart
  }
  const combined = `${prefix}${numericText}${suffix}`
  return autoNegative && !combined.startsWith('-') ? `-${combined}` : combined
}

const BAHTTEXT_MAX_SATANG: f64 = 9007199254740991.0

function trimLeadingZerosText(text: string): string {
  let index = 0
  while (index + 1 < text.length && text.charCodeAt(index) == 48) {
    index += 1
  }
  return text.slice(index)
}

function isAllZeroText(text: string): bool {
  for (let index = 0; index < text.length; index += 1) {
    if (text.charCodeAt(index) != 48) {
      return false
    }
  }
  return true
}

function bahtDigitWord(digit: i32): string {
  switch (digit) {
    case 0:
      return 'ศูนย์'
    case 1:
      return 'หนึ่ง'
    case 2:
      return 'สอง'
    case 3:
      return 'สาม'
    case 4:
      return 'สี่'
    case 5:
      return 'ห้า'
    case 6:
      return 'หก'
    case 7:
      return 'เจ็ด'
    case 8:
      return 'แปด'
    case 9:
      return 'เก้า'
    default:
      return ''
  }
}

function bahtScaleWord(position: i32): string {
  switch (position) {
    case 1:
      return 'สิบ'
    case 2:
      return 'ร้อย'
    case 3:
      return 'พัน'
    case 4:
      return 'หมื่น'
    case 5:
      return 'แสน'
    default:
      return ''
  }
}

function bahtSegmentText(text: string): string {
  const normalized = trimLeadingZerosText(text)
  if (normalized.length == 0 || isAllZeroText(normalized)) {
    return ''
  }

  let output = ''
  let hasHigherNonZero = false
  const length = normalized.length
  for (let index = 0; index < length; index += 1) {
    const digit = <i32>(normalized.charCodeAt(index) - 48)
    if (digit == 0) {
      continue
    }
    const position = length - index - 1
    if (position == 0) {
      output += digit == 1 && hasHigherNonZero ? 'เอ็ด' : bahtDigitWord(digit)
    } else if (position == 1) {
      if (digit == 1) {
        output += 'สิบ'
      } else if (digit == 2) {
        output += 'ยี่สิบ'
      } else {
        output += bahtDigitWord(digit) + 'สิบ'
      }
    } else {
      output += bahtDigitWord(digit) + bahtScaleWord(position)
    }
    hasHigherNonZero = true
  }
  return output
}

function bahtIntegerText(text: string): string {
  const normalized = trimLeadingZerosText(text)
  if (normalized.length == 0 || isAllZeroText(normalized)) {
    return 'ศูนย์'
  }
  if (normalized.length > 6) {
    const split = normalized.length - 6
    return bahtIntegerText(normalized.slice(0, split)) + 'ล้าน' + bahtSegmentText(normalized.slice(split))
  }
  const segment = bahtSegmentText(normalized)
  return segment.length == 0 ? 'ศูนย์' : segment
}

export function bahtTextFromNumber(value: f64): string | null {
  if (!isFinite(value)) {
    return null
  }
  const absolute = Math.abs(value)
  const scaled = Math.round(absolute * 100.0)
  if (!isFinite(scaled) || scaled > BAHTTEXT_MAX_SATANG) {
    return null
  }
  const scaledSatang = <i64>scaled
  const baht = scaledSatang / 100
  const satang = <i32>(scaledSatang % 100)
  const prefix = value < 0.0 ? 'ลบ' : ''
  const bahtText = bahtIntegerText(baht.toString())
  if (satang == 0) {
    return prefix + bahtText + 'บาทถ้วน'
  }
  return prefix + bahtText + 'บาท' + bahtSegmentText(satang.toString()) + 'สตางค์'
}

function removeAsciiWhitespace(input: string): string {
  let result = ''
  for (let index = 0; index < input.length; index += 1) {
    const char = input.charCodeAt(index)
    if (char <= 32) {
      continue
    }
    result += String.fromCharCode(char)
  }
  return result
}

export function numberValueParseText(input: string, decimalSeparator: string, groupSeparator: string): f64 {
  const compact = removeAsciiWhitespace(input)
  if (compact.length == 0) {
    return 0.0
  }

  let percentCount = 0
  let coreEnd = compact.length
  while (coreEnd > 0 && compact.charCodeAt(coreEnd - 1) == 37) {
    coreEnd -= 1
    percentCount += 1
  }
  const core = coreEnd == compact.length ? compact : compact.slice(0, coreEnd)
  if (core.indexOf('%') >= 0) {
    return NaN
  }

  const decimal = decimalSeparator.length == 0 ? '.' : String.fromCharCode(decimalSeparator.charCodeAt(0))
  const group = groupSeparator.length == 0 ? '' : String.fromCharCode(groupSeparator.charCodeAt(0))
  if (decimal.length > 0 && group.length > 0 && decimal == group) {
    return NaN
  }

  const decimalIndex = decimal.length == 0 ? -1 : core.indexOf(decimal)
  if (decimalIndex >= 0 && core.indexOf(decimal, decimalIndex + decimal.length) >= 0) {
    return NaN
  }

  let normalized = ''
  for (let index = 0; index < core.length; index += 1) {
    const char = String.fromCharCode(core.charCodeAt(index))
    if (group.length > 0 && char == group) {
      if (decimalIndex >= 0 && index > decimalIndex) {
        return NaN
      }
      continue
    }
    if (decimal.length > 0 && char == decimal) {
      normalized += '.'
      continue
    }
    normalized += char
  }
  if (normalized == '' || normalized == '.' || normalized == '+' || normalized == '-') {
    return NaN
  }

  const parsed = parseNumericText(normalized)
  if (!isFinite(parsed)) {
    return NaN
  }
  return parsed / Math.pow(100.0, <f64>percentCount)
}

export function errorLabel(code: i32): string {
  if (code == ErrorCode.Div0) return '#DIV/0!'
  if (code == ErrorCode.Ref) return '#REF!'
  if (code == ErrorCode.Value) return '#VALUE!'
  if (code == ErrorCode.Name) return '#NAME?'
  if (code == ErrorCode.NA) return '#N/A'
  if (code == ErrorCode.Cycle) return '#CYCLE!'
  if (code == ErrorCode.Spill) return '#SPILL!'
  if (code == ErrorCode.Blocked) return '#BLOCKED!'
  return '#ERROR!'
}

function hexDigit(value: i32): string {
  return String.fromCharCode(value < 10 ? 48 + value : 55 + value)
}

export function jsonQuoteText(input: string): string {
  let result = '"'
  for (let index = 0; index < input.length; index += 1) {
    const code = input.charCodeAt(index)
    if (code == 34) {
      result += '\\"'
    } else if (code == 92) {
      result += '\\\\'
    } else if (code == 8) {
      result += '\\b'
    } else if (code == 12) {
      result += '\\f'
    } else if (code == 10) {
      result += '\\n'
    } else if (code == 13) {
      result += '\\r'
    } else if (code == 9) {
      result += '\\t'
    } else if (code < 32) {
      result += '\\u00' + hexDigit((code >> 4) & 0xf) + hexDigit(code & 0xf)
    } else {
      result += String.fromCharCode(code)
    }
  }
  return result + '"'
}
