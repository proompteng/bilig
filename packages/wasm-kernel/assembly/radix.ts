const BASE_DIGITS: string = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'

export function toBaseText(value: i64, radix: i32, minLength: i32): string {
  if (value == 0) {
    let zeroText = '0'
    while (zeroText.length < minLength) {
      zeroText = '0' + zeroText
    }
    return zeroText
  }
  let current = value
  let output = ''
  while (current > 0) {
    const digit = <i32>(current % <i64>radix)
    output = BASE_DIGITS.charAt(digit) + output
    current /= <i64>radix
  }
  while (output.length < minLength) {
    output = '0' + output
  }
  return output
}

function baseDigitValue(code: i32): i32 {
  if (code >= 48 && code <= 57) {
    return code - 48
  }
  if (code >= 65 && code <= 90) {
    return code - 55
  }
  if (code >= 97 && code <= 122) {
    return code - 87
  }
  return -1
}

export function isValidBaseText(text: string, radix: i32): bool {
  if (text.length == 0) {
    return false
  }
  for (let index = 0; index < text.length; index += 1) {
    const digit = baseDigitValue(text.charCodeAt(index))
    if (digit < 0 || digit >= radix) {
      return false
    }
  }
  return true
}

export function parseBaseText(text: string, radix: i32): f64 {
  let output = 0.0
  for (let index = 0; index < text.length; index += 1) {
    const digit = baseDigitValue(text.charCodeAt(index))
    if (digit < 0 || digit >= radix) {
      return NaN
    }
    output = output * <f64>radix + <f64>digit
  }
  return output
}

function powRadixI64(radix: i32, width: i32): i64 {
  let output: i64 = 1
  for (let index = 0; index < width; index += 1) {
    output *= <i64>radix
  }
  return output
}

export function parseSignedRadixText(text: string, radix: i32, width: i32): i64 {
  if (text.length == 0 || text.length > width || !isValidBaseText(text, radix)) {
    return i64.MIN_VALUE
  }
  let parsed: i64 = 0
  for (let index = 0; index < text.length; index += 1) {
    const digit = baseDigitValue(text.charCodeAt(index))
    if (digit < 0 || digit >= radix) {
      return i64.MIN_VALUE
    }
    parsed = parsed * <i64>radix + <i64>digit
  }
  const fullRange = powRadixI64(radix, width)
  const negativeThreshold = fullRange / 2
  return text.length == width && parsed >= negativeThreshold ? parsed - fullRange : parsed
}

export function formatSignedRadixText(
  numeric: i64,
  radix: i32,
  minLength: i32,
  negativeWidth: i32,
  minValue: i64,
  maxValue: i64,
): string | null {
  if (numeric < minValue || numeric > maxValue) {
    return null
  }
  if (numeric < 0) {
    const encoded = numeric + powRadixI64(radix, negativeWidth)
    return toBaseText(encoded, radix, negativeWidth)
  }
  const raw = toBaseText(numeric, radix, 0)
  if (minLength < raw.length) {
    return null
  }
  return toBaseText(numeric, radix, minLength)
}
