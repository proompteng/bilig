import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { scalarText, trimAsciiWhitespace } from './text-codec'
import { formatFixedText } from './text-format'
import { coerceNonNegativeShift } from './numeric-core'
import { formatSignedRadixText, isValidBaseText, parseBaseText, parseSignedRadixText, toBaseText } from './radix'
import {
  CONVERT_GROUP_TEMPERATURE,
  convertKelvinToTemperature,
  convertTemperatureToKelvin,
  resolveConvertUnit,
  resolvedConvertFactor,
  resolvedConvertGroup,
  resolvedConvertTemperature,
} from './unit-convert'
import { toNumberExact } from './operands'
import { coercePositiveIntegerArg } from './builtin-args'
import { STACK_KIND_SCALAR, writeResult, writeStringResult } from './result-io'

const MAX_SAFE_INTEGER_F64: f64 = 9007199254740991.0

function columnLabelText(column: i32): string | null {
  if (column < 1) {
    return null
  }
  let current = column
  let label = ''
  while (current > 0) {
    const offset = (current - 1) % 26
    label = String.fromCharCode(65 + offset) + label
    current = (current - 1) / 26
  }
  return label
}

function escapeSheetNameText(value: string): string {
  let output = ''
  for (let index = 0; index < value.length; index++) {
    const char = value.charAt(index)
    output += char
    if (char == "'") {
      output += "'"
    }
  }
  return output
}

function digitCount(value: i32): i32 {
  if (value <= 0) {
    return 1
  }
  let current = value
  let count = 0
  while (current > 0) {
    count += 1
    current /= 10
  }
  return count
}

function isValidDollarFractionNative(fraction: i32): bool {
  if (fraction <= 0) {
    return false
  }
  if (fraction == 1) {
    return true
  }
  let current = fraction
  while ((current & 1) == 0) {
    current >>= 1
  }
  return current == 1
}

function parsePositiveDigits(value: string): i32 {
  if (value.length == 0) {
    return 0
  }
  let output = 0
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index)
    if (code < 48 || code > 57) {
      return i32.MIN_VALUE
    }
    output = output * 10 + (code - 48)
  }
  return output
}

function dollarFractionalNumerator(value: f64): i32 {
  const absoluteText = Math.abs(value).toString()
  const dot = absoluteText.indexOf('.')
  if (dot < 0) {
    return 0
  }
  return parsePositiveDigits(absoluteText.substring(dot + 1))
}

function signedRadixInputText(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  if (tag == ValueTag.Error) {
    return null
  }
  if (tag == ValueTag.String) {
    const text = scalarText(
      tag,
      value,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (text == null) {
      return null
    }
    return trimAsciiWhitespace(text).toUpperCase()
  }
  const numeric = toNumberExact(tag, value)
  if (!isFinite(numeric)) {
    return null
  }
  return (<i64>numeric).toString().toUpperCase()
}

function roundToPlacesNative(value: f64, places: i32): f64 {
  const scale = Math.pow(10.0, <f64>places)
  return Math.round(value * scale) / scale
}

function roundToSignificantDigitsNative(value: f64, digits: i32): f64 {
  if (value == 0.0 || !isFinite(value)) {
    return value
  }
  const exponent = <i32>Math.floor(Math.log(Math.abs(value)) / Math.log(10.0))
  const scale = Math.pow(10.0, <f64>(digits - exponent - 1))
  return Math.round(value * scale) / scale
}

function euroRateNative(code: string): f64 {
  if (code == 'BEF' || code == 'LUF') return 40.3399
  if (code == 'DEM') return 1.95583
  if (code == 'ESP') return 166.386
  if (code == 'FRF') return 6.55957
  if (code == 'IEP') return 0.787564
  if (code == 'ITL') return 1936.27
  if (code == 'NLG') return 2.20371
  if (code == 'ATS') return 13.7603
  if (code == 'PTE') return 200.482
  if (code == 'FIM') return 5.94573
  if (code == 'GRD') return 340.75
  if (code == 'SIT') return 239.64
  if (code == 'EUR') return 1.0
  return NaN
}

function euroCalculationPrecisionNative(code: string): i32 {
  if (code == 'BEF' || code == 'LUF' || code == 'ESP' || code == 'ITL' || code == 'PTE' || code == 'GRD') {
    return 0
  }
  if (
    code == 'DEM' ||
    code == 'FRF' ||
    code == 'IEP' ||
    code == 'NLG' ||
    code == 'ATS' ||
    code == 'FIM' ||
    code == 'SIT' ||
    code == 'EUR'
  ) {
    return 2
  }
  return i32.MIN_VALUE
}

export function tryApplyFormatConvertBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if (builtinId == BuiltinId.Address && argc >= 2 && argc <= 5) {
    const row = coercePositiveIntegerArg(tagStack[base], valueStack[base], true, 1)
    const column = coercePositiveIntegerArg(tagStack[base + 1], valueStack[base + 1], true, 1)
    if (row == i32.MIN_VALUE || column == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const absNumeric = argc >= 3 ? toNumberExact(tagStack[base + 2], valueStack[base + 2]) : 1.0
    const refStyleNumeric = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 1.0
    if (!isFinite(absNumeric) || !isFinite(refStyleNumeric)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const absNum = <i32>absNumeric
    const refStyle = <i32>refStyleNumeric
    if (absNum < 1 || absNum > 4 || (refStyle != 1 && refStyle != 2)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let sheetPrefix = ''
    if (argc == 5) {
      if (tagStack[base + 4] == ValueTag.Empty || tagStack[base + 4] != ValueTag.String) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const sheetText = scalarText(
        tagStack[base + 4],
        valueStack[base + 4],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (sheetText == null) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      sheetPrefix = `'${escapeSheetNameText(sheetText)}'!`
    }
    const columnLabel = columnLabelText(column)
    if (columnLabel == null) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (refStyle == 2) {
      const rowLabel = absNum == 1 || absNum == 2 ? row.toString() : `[${row}]`
      const colLabel = absNum == 1 || absNum == 3 ? column.toString() : `[${column}]`
      return writeStringResult(base, `${sheetPrefix}R${rowLabel}C${colLabel}`, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const rowLabel = absNum == 1 || absNum == 2 ? `$${row.toString()}` : row.toString()
    const colLabel = absNum == 1 || absNum == 3 ? `$${columnLabel}` : columnLabel
    return writeStringResult(base, `${sheetPrefix}${colLabel}${rowLabel}`, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Dollar && argc >= 1 && argc <= 3) {
    const value = toNumberExact(tagStack[base], valueStack[base])
    const decimalsNumeric = argc >= 2 ? toNumberExact(tagStack[base + 1], valueStack[base + 1]) : 2.0
    let noCommasValue = 0.0
    if (argc >= 3) {
      const numeric = toNumberExact(tagStack[base + 2], valueStack[base + 2])
      noCommasValue = isNaN(numeric) ? 0.0 : numeric
    }
    if (!isFinite(value) || !isFinite(decimalsNumeric)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const text = formatFixedText(value, <i32>decimalsNumeric, noCommasValue == 0.0)
    if (text == null || text.length == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const normalizedText = text.startsWith('-') ? text.slice(1) : text
    return writeStringResult(base, value < 0.0 ? `-$${normalizedText}` : `$${text}`, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Dollarde || builtinId == BuiltinId.Dollarfr) && argc == 2) {
    const value = toNumberExact(tagStack[base], valueStack[base])
    const fractionNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    if (!isFinite(value) || !isFinite(fractionNumeric)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const fraction = <i32>fractionNumeric
    if (!isValidDollarFractionNative(fraction)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (builtinId == BuiltinId.Dollarde) {
      const integerPart = <i32>Math.floor(Math.abs(value))
      const fractionalNumerator = dollarFractionalNumerator(value)
      if (fractionalNumerator == i32.MIN_VALUE || fractionalNumerator >= fraction) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const sign = value < 0.0 ? -1.0 : 1.0
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        sign * (<f64>integerPart + <f64>fractionalNumerator / <f64>fraction),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    const sign = value < 0.0 ? -1.0 : 1.0
    const absolute = Math.abs(value)
    const integerPart = <i32>Math.floor(absolute)
    const fractional = absolute - <f64>integerPart
    const width = digitCount(fraction)
    const scaledNumerator = <i32>Math.round(fractional * <f64>fraction)
    const carry = scaledNumerator / fraction
    const numerator = scaledNumerator - carry * fraction
    const outputValue = <f64>(integerPart + carry) + <f64>numerator / Math.pow(10.0, <f64>width)
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sign * outputValue, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Base && (argc == 2 || argc == 3)) {
    const numberNumeric = toNumberExact(tagStack[base], valueStack[base])
    const radixNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const minLengthNumeric = argc == 3 ? toNumberExact(tagStack[base + 2], valueStack[base + 2]) : 0.0
    if (!isFinite(numberNumeric) || !isFinite(radixNumeric) || !isFinite(minLengthNumeric)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numberValue = <i64>numberNumeric
    const radixValue = <i32>radixNumeric
    const minLength = <i32>minLengthNumeric
    if (numberValue < 0 || radixValue < 2 || radixValue > 36 || minLength < 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeStringResult(base, toBaseText(numberValue, radixValue, minLength), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Decimal && argc == 2) {
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const radixNumeric = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    if (!isFinite(radixNumeric)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const radixValue = <i32>radixNumeric
    if (radixValue < 2 || radixValue > 36) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let raw = ''
    if (tagStack[base] == ValueTag.String) {
      const text = scalarText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (text == null) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      raw = trimAsciiWhitespace(text)
    } else {
      const numeric = toNumberExact(tagStack[base], valueStack[base])
      if (!isFinite(numeric)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      raw = (<i64>numeric).toString()
    }
    if (raw.length == 0 || !isValidBaseText(raw, radixValue)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      parseBaseText(raw, radixValue),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if ((builtinId == BuiltinId.Bin2dec || builtinId == BuiltinId.Hex2dec || builtinId == BuiltinId.Oct2dec) && argc == 1) {
    const radix = builtinId == BuiltinId.Bin2dec ? 2 : builtinId == BuiltinId.Hex2dec ? 16 : 8
    const raw = signedRadixInputText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (raw == null) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const numeric = parseSignedRadixText(raw, radix, 10)
    if (numeric == i64.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>numeric, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (
    (builtinId == BuiltinId.Bin2hex ||
      builtinId == BuiltinId.Bin2oct ||
      builtinId == BuiltinId.Dec2bin ||
      builtinId == BuiltinId.Dec2hex ||
      builtinId == BuiltinId.Dec2oct ||
      builtinId == BuiltinId.Hex2bin ||
      builtinId == BuiltinId.Hex2oct ||
      builtinId == BuiltinId.Oct2bin ||
      builtinId == BuiltinId.Oct2hex) &&
    (argc == 1 || argc == 2)
  ) {
    let numeric: i64 = i64.MIN_VALUE
    let radix = 10
    let negativeWidth = 10
    let minValue: i64 = -549755813888
    let maxValue: i64 = 549755813887

    if (builtinId == BuiltinId.Bin2hex || builtinId == BuiltinId.Bin2oct) {
      const raw = signedRadixInputText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (raw != null) {
        numeric = parseSignedRadixText(raw, 2, 10)
      }
      if (builtinId == BuiltinId.Bin2oct) {
        radix = 8
        minValue = -536870912
        maxValue = 536870911
      } else {
        radix = 16
      }
    } else if (builtinId == BuiltinId.Hex2bin || builtinId == BuiltinId.Hex2oct) {
      const raw = signedRadixInputText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (raw != null) {
        numeric = parseSignedRadixText(raw, 16, 10)
      }
      if (builtinId == BuiltinId.Hex2bin) {
        radix = 2
        minValue = -512
        maxValue = 511
      } else {
        radix = 8
        minValue = -536870912
        maxValue = 536870911
      }
    } else if (builtinId == BuiltinId.Oct2bin || builtinId == BuiltinId.Oct2hex) {
      const raw = signedRadixInputText(
        tagStack[base],
        valueStack[base],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (raw != null) {
        numeric = parseSignedRadixText(raw, 8, 10)
      }
      if (builtinId == BuiltinId.Oct2bin) {
        radix = 2
        minValue = -512
        maxValue = 511
      } else {
        radix = 16
      }
    } else {
      const inputNumeric = toNumberExact(tagStack[base], valueStack[base])
      if (isFinite(inputNumeric) && Math.abs(inputNumeric) <= MAX_SAFE_INTEGER_F64) {
        numeric = <i64>inputNumeric
      }
      if (builtinId == BuiltinId.Dec2bin) {
        radix = 2
        minValue = -512
        maxValue = 511
      } else if (builtinId == BuiltinId.Dec2oct) {
        radix = 8
        minValue = -536870912
        maxValue = 536870911
      } else {
        radix = 16
      }
    }

    const places = argc == 2 ? coerceNonNegativeShift(tagStack[base + 1], valueStack[base + 1]) : 0
    if (numeric == i64.MIN_VALUE || places == i64.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const text = formatSignedRadixText(numeric, radix, <i32>places, negativeWidth, minValue, maxValue)
    if (text == null) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Convert && argc == 3) {
    const numeric = toNumberExact(tagStack[base], valueStack[base])
    if (!isFinite(numeric) || tagStack[base + 1] != ValueTag.String || tagStack[base + 2] != ValueTag.String) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const fromText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const toText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (fromText == null || toText == null || !resolveConvertUnit(fromText)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const fromGroup = resolvedConvertGroup
    const fromFactor = resolvedConvertFactor
    const fromTemperature = resolvedConvertTemperature
    if (!resolveConvertUnit(toText)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const toGroup = resolvedConvertGroup
    const toFactor = resolvedConvertFactor
    const toTemperature = resolvedConvertTemperature
    let result = NaN
    if (fromGroup == CONVERT_GROUP_TEMPERATURE || toGroup == CONVERT_GROUP_TEMPERATURE) {
      if (fromGroup != CONVERT_GROUP_TEMPERATURE || toGroup != CONVERT_GROUP_TEMPERATURE) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      result = convertKelvinToTemperature(toTemperature, convertTemperatureToKelvin(fromTemperature, numeric))
    } else {
      if (fromGroup != toGroup) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      result = (numeric * fromFactor) / toFactor
    }
    if (!isFinite(result)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Euroconvert && argc >= 3 && argc <= 5) {
    const numeric = toNumberExact(tagStack[base], valueStack[base])
    if (!isFinite(numeric) || tagStack[base + 1] != ValueTag.String || tagStack[base + 2] != ValueTag.String) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const targetText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const fullPrecisionNumeric = argc >= 4 ? toNumberExact(tagStack[base + 3], valueStack[base + 3]) : 0.0
    const triangulationNumeric = argc == 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : NaN
    if (
      sourceText == null ||
      targetText == null ||
      !isFinite(fullPrecisionNumeric) ||
      (argc == 5 && (!isFinite(triangulationNumeric) || triangulationNumeric < 3.0))
    ) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const sourceRate = euroRateNative(sourceText)
    const targetRate = euroRateNative(targetText)
    const targetPrecision = euroCalculationPrecisionNative(targetText)
    if (!isFinite(sourceRate) || !isFinite(targetRate) || targetPrecision == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (sourceText == targetText) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, numeric, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let euroValue = sourceText == 'EUR' ? numeric : numeric / sourceRate
    if (sourceText != 'EUR' && argc == 5) {
      euroValue = roundToSignificantDigitsNative(euroValue, <i32>triangulationNumeric)
    }
    let result = targetText == 'EUR' ? euroValue : euroValue * targetRate
    if (fullPrecisionNumeric == 0.0) {
      result = roundToPlacesNative(result, targetPrecision)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
