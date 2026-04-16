import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { inputCellNumeric, inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot, toNumberOrNaN } from './operands'
import { collectStatValuesFromArgs, isNumericResult, lastStatCollectionErrorCode } from './builtin-args'
import { coerceInteger } from './numeric-core'
import {
  interpolateSortedPercentRank,
  interpolateSortedPercentile,
  kurtosisOf,
  modeSingleOf,
  populationVarianceOf,
  sampleVarianceOf,
  skewPopulationOf,
  skewSampleOf,
  sortNumericValues,
  truncateToSignificance,
} from './statistics-core'
import {
  collectNumericCellRangeSeriesFromSlot,
  collectNumericValuesFromArgs,
  collectNumericValuesFromSlot,
  collectSampleNumbersFromSlot,
  orderStatisticErrorCode,
  sampleCollectionErrorCode,
} from './statistics-tests'
import { STACK_KIND_SCALAR, writeArrayResult, writeResult } from './result-io'
import { allocateSpillArrayResult, readSpillArrayNumber, writeSpillArrayNumber } from './vm'

export function tryApplyStatisticsSummaryBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
): i32 {
  if ((builtinId == BuiltinId.Rank || builtinId == BuiltinId.RankEq || builtinId == BuiltinId.RankAvg) && (argc == 2 || argc == 3)) {
    if (kindStack[base] != STACK_KIND_SCALAR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const target = toNumberOrNaN(tagStack[base], valueStack[base])
    if (!isFinite(target)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let order = 0
    if (argc == 3) {
      if (kindStack[base + 2] != STACK_KIND_SCALAR) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[base + 2] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[base + 2],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }
      const requestedOrder = coerceInteger(tagStack[base + 2], valueStack[base + 2])
      if (requestedOrder == i32.MIN_VALUE || (requestedOrder != 0 && requestedOrder != 1)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      order = requestedOrder
    }

    const arraySlot = base + 1
    const rows = inputRowsFromSlot(arraySlot, kindStack, rangeIndexStack, rangeRowCounts)
    const cols = inputColsFromSlot(arraySlot, kindStack, rangeIndexStack, rangeColCounts)
    if (rows < 1 || cols < 1) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let preceding = 0
    let ties = 0
    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        const valueTag = inputCellTag(
          arraySlot,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        )
        if (valueTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            inputCellScalarValue(
              arraySlot,
              row,
              col,
              kindStack,
              valueStack,
              tagStack,
              rangeIndexStack,
              rangeOffsets,
              rangeLengths,
              rangeRowCounts,
              rangeColCounts,
              rangeMembers,
              cellTags,
              cellNumbers,
              cellStringIds,
              cellErrors,
            ),
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          )
        }

        const numeric = inputCellNumeric(
          arraySlot,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        )
        if (!isFinite(numeric)) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }

        if (numeric == target) {
          ties += 1
          continue
        }
        if ((order == 0 && numeric > target) || (order == 1 && numeric < target)) {
          preceding += 1
        }
      }
    }

    if (ties == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const rank = builtinId == BuiltinId.RankAvg ? <f64>preceding + (<f64>ties + 1.0) / 2.0 : <f64>(preceding + 1)
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, rank, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (
    (builtinId == BuiltinId.Stdev ||
      builtinId == BuiltinId.StdevP ||
      builtinId == BuiltinId.StdevS ||
      builtinId == BuiltinId.Stdeva ||
      builtinId == BuiltinId.Stdevp ||
      builtinId == BuiltinId.Stdevpa ||
      builtinId == BuiltinId.Var ||
      builtinId == BuiltinId.VarP ||
      builtinId == BuiltinId.VarS ||
      builtinId == BuiltinId.Vara ||
      builtinId == BuiltinId.Varp ||
      builtinId == BuiltinId.Varpa ||
      builtinId == BuiltinId.Skew ||
      builtinId == BuiltinId.SkewP ||
      builtinId == BuiltinId.Kurt) &&
    argc >= 1
  ) {
    const values = collectStatValuesFromArgs(
      base,
      argc,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
      builtinId == BuiltinId.Stdeva || builtinId == BuiltinId.Stdevpa || builtinId == BuiltinId.Vara || builtinId == BuiltinId.Varpa,
    )
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        lastStatCollectionErrorCode(),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    let result = NaN
    if (builtinId == BuiltinId.Stdev || builtinId == BuiltinId.StdevS || builtinId == BuiltinId.Stdeva) {
      result = Math.sqrt(sampleVarianceOf(values))
    } else if (builtinId == BuiltinId.StdevP || builtinId == BuiltinId.Stdevp || builtinId == BuiltinId.Stdevpa) {
      result = Math.sqrt(populationVarianceOf(values))
    } else if (builtinId == BuiltinId.Var || builtinId == BuiltinId.VarS || builtinId == BuiltinId.Vara) {
      result = sampleVarianceOf(values)
    } else if (builtinId == BuiltinId.VarP || builtinId == BuiltinId.Varp || builtinId == BuiltinId.Varpa) {
      result = populationVarianceOf(values)
    } else if (builtinId == BuiltinId.Skew) {
      result = skewSampleOf(values)
    } else if (builtinId == BuiltinId.SkewP) {
      result = skewPopulationOf(values)
    } else {
      result = kurtosisOf(values)
    }

    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNumericResult(result) ? <u8>ValueTag.Number : <u8>ValueTag.Error,
      isNumericResult(result) ? result : ErrorCode.Value,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Median && argc >= 1) {
    const values = collectNumericValuesFromArgs(
      base,
      argc,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    )
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (values.length == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    sortNumericValues(values)
    const middle = values.length >>> 1
    const median = values.length % 2 == 0 ? (unchecked(values[middle - 1]) + unchecked(values[middle])) / 2.0 : unchecked(values[middle])
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, median, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Small || builtinId == BuiltinId.Large) && argc == 2) {
    if (kindStack[base + 1] != STACK_KIND_SCALAR || tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        tagStack[base + 1] == ValueTag.Error ? valueStack[base + 1] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    const position = coerceInteger(tagStack[base + 1], valueStack[base + 1])
    const values = collectNumericValuesFromSlot(
      base,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    )
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (position < 1 || position > values.length) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    sortNumericValues(values)
    const index = builtinId == BuiltinId.Small ? position - 1 : values.length - position
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      unchecked(values[index]),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (
    (builtinId == BuiltinId.Percentile ||
      builtinId == BuiltinId.PercentileInc ||
      builtinId == BuiltinId.PercentileExc ||
      builtinId == BuiltinId.Quartile ||
      builtinId == BuiltinId.QuartileInc ||
      builtinId == BuiltinId.QuartileExc) &&
    argc == 2
  ) {
    if (kindStack[base + 1] != STACK_KIND_SCALAR || tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        tagStack[base + 1] == ValueTag.Error ? valueStack[base + 1] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    const values = collectNumericValuesFromSlot(
      base,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    )
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (values.length == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    sortNumericValues(values)

    if (builtinId == BuiltinId.Quartile || builtinId == BuiltinId.QuartileInc || builtinId == BuiltinId.QuartileExc) {
      const quartile = coerceInteger(tagStack[base + 1], valueStack[base + 1])
      const exclusive = builtinId == BuiltinId.QuartileExc
      if (quartile < 0 || quartile > 4 || (exclusive && (quartile <= 0 || quartile >= 4))) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (!exclusive && quartile == 0) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          unchecked(values[0]),
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }
      if (!exclusive && quartile == 4) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Number,
          unchecked(values[values.length - 1]),
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }
      const quartileValue = interpolateSortedPercentile(values, <f64>quartile / 4.0, exclusive)
      if (isNaN(quartileValue)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, quartileValue, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const percentile = toNumberOrNaN(tagStack[base + 1], valueStack[base + 1])
    const percentileValue = interpolateSortedPercentile(values, percentile, builtinId == BuiltinId.PercentileExc)
    if (isNaN(percentileValue)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, percentileValue, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (
    (builtinId == BuiltinId.Percentrank || builtinId == BuiltinId.PercentrankInc || builtinId == BuiltinId.PercentrankExc) &&
    (argc == 2 || argc == 3)
  ) {
    if (kindStack[base + 1] != STACK_KIND_SCALAR || tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        tagStack[base + 1] == ValueTag.Error ? valueStack[base + 1] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (argc == 3 && (kindStack[base + 2] != STACK_KIND_SCALAR || tagStack[base + 2] == ValueTag.Error)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        argc == 3 && tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    const values = collectNumericValuesFromSlot(
      base,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    )
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        orderStatisticErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (values.length < 2) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const target = toNumberOrNaN(tagStack[base + 1], valueStack[base + 1])
    const significance = argc == 3 ? coerceInteger(tagStack[base + 2], valueStack[base + 2]) : 3
    if (!isFinite(target) || significance < 1) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    sortNumericValues(values)
    const percentRank = interpolateSortedPercentRank(values, target, builtinId == BuiltinId.PercentrankExc)
    if (isNaN(percentRank)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      truncateToSignificance(percentRank, significance),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if ((builtinId == BuiltinId.Mode || builtinId == BuiltinId.ModeSngl) && argc >= 1) {
    const values = new Array<f64>()
    for (let index = 0; index < argc; index += 1) {
      const collected = collectSampleNumbersFromSlot(
        base + index,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
      if (collected === null) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          sampleCollectionErrorCode,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }
      for (let cursor = 0; cursor < collected.length; cursor += 1) {
        values.push(unchecked(collected[cursor]))
      }
    }
    const mode = modeSingleOf(values)
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      isNaN(mode) ? <u8>ValueTag.Error : <u8>ValueTag.Number,
      isNaN(mode) ? ErrorCode.NA : mode,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.ModeMult && argc >= 1) {
    const values = new Array<f64>()
    for (let index = 0; index < argc; index += 1) {
      const collected = collectSampleNumbersFromSlot(
        base + index,
        kindStack,
        valueStack,
        tagStack,
        rangeIndexStack,
        rangeOffsets,
        rangeLengths,
        rangeRowCounts,
        rangeColCounts,
        rangeMembers,
        cellTags,
        cellNumbers,
        cellStringIds,
        cellErrors,
      )
      if (collected === null) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          sampleCollectionErrorCode,
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }
      for (let cursor = 0; cursor < collected.length; cursor += 1) {
        values.push(unchecked(collected[cursor]))
      }
    }
    if (values.length == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    sortNumericValues(values)
    let maxCount = 0
    let modeCount = 0
    for (let start = 0; start < values.length; ) {
      let end = start + 1
      while (end < values.length && unchecked(values[end]) == unchecked(values[start])) {
        end += 1
      }
      const count = end - start
      if (count > maxCount) {
        maxCount = count
        modeCount = 1
      } else if (count == maxCount) {
        modeCount += 1
      }
      start = end
    }
    if (maxCount < 2) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const arrayIndex = allocateSpillArrayResult(modeCount, 1)
    let output = 0
    for (let start = 0; start < values.length; ) {
      let end = start + 1
      while (end < values.length && unchecked(values[end]) == unchecked(values[start])) {
        end += 1
      }
      if (end - start == maxCount) {
        writeSpillArrayNumber(arrayIndex, output, unchecked(values[start]))
        output += 1
      }
      start = end
    }
    return writeArrayResult(base, arrayIndex, modeCount, 1, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Frequency && argc == 2) {
    const dataValues = collectSampleNumbersFromSlot(
      base,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    )
    if (dataValues === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    const binValues = collectSampleNumbersFromSlot(
      base + 1,
      kindStack,
      valueStack,
      tagStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeRowCounts,
      rangeColCounts,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellStringIds,
      cellErrors,
    )
    if (binValues === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    sortNumericValues(binValues)
    const bucketCount = binValues.length + 1
    const arrayIndex = allocateSpillArrayResult(bucketCount, 1)
    for (let index = 0; index < bucketCount; index += 1) {
      writeSpillArrayNumber(arrayIndex, index, 0.0)
    }
    for (let dataIndex = 0; dataIndex < dataValues.length; dataIndex += 1) {
      const value = unchecked(dataValues[dataIndex])
      let bucket = binValues.length
      for (let binIndex = 0; binIndex < binValues.length; binIndex += 1) {
        if (value <= unchecked(binValues[binIndex])) {
          bucket = binIndex
          break
        }
      }
      writeSpillArrayNumber(arrayIndex, bucket, readSpillArrayNumber(arrayIndex, bucket) + 1.0)
    }
    return writeArrayResult(base, arrayIndex, bucketCount, 1, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Prob && (argc == 3 || argc == 4)) {
    if (kindStack[base + 2] != STACK_KIND_SCALAR || (argc == 4 && kindStack[base + 3] != STACK_KIND_SCALAR)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base + 2] == ValueTag.Error || (argc == 4 && tagStack[base + 3] == ValueTag.Error)) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        tagStack[base + 2] == ValueTag.Error ? valueStack[base + 2] : valueStack[base + 3],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const lower = toNumberOrNaN(tagStack[base + 2], valueStack[base + 2])
    const upper = argc == 4 ? toNumberOrNaN(tagStack[base + 3], valueStack[base + 3]) : lower
    if (!isFinite(lower) || !isFinite(upper) || upper < lower) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const xRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const xCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    const probabilityRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts)
    const probabilityCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts)
    if (xRows < 1 || xCols < 1 || probabilityRows < 1 || probabilityCols < 1 || xRows != probabilityRows || xCols != probabilityCols) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let probabilitySum = 0.0
    let total = 0.0
    for (let row = 0; row < xRows; row += 1) {
      for (let col = 0; col < xCols; col += 1) {
        const xTag = inputCellTag(
          base,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        )
        const xValue = inputCellScalarValue(
          base,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        )
        if (xTag == ValueTag.Error) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, xValue, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        const probabilityTag = inputCellTag(
          base + 1,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        )
        const probabilityValue = inputCellScalarValue(
          base + 1,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        )
        if (probabilityTag == ValueTag.Error) {
          return writeResult(
            base,
            STACK_KIND_SCALAR,
            <u8>ValueTag.Error,
            probabilityValue,
            rangeIndexStack,
            valueStack,
            tagStack,
            kindStack,
          )
        }
        if (xTag != ValueTag.Number || probabilityTag != ValueTag.Number) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        if (!isFinite(probabilityValue) || probabilityValue < 0 || probabilityValue > 1) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        probabilitySum += probabilityValue
        if (xValue >= lower && xValue <= upper) {
          total += probabilityValue
        }
      }
    }

    if (Math.abs(probabilitySum - 1.0) > 1e-9) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, total, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Trimmean && argc == 2) {
    if (kindStack[base + 1] != STACK_KIND_SCALAR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base + 1] == ValueTag.Error) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        valueStack[base + 1],
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    const percent = toNumberOrNaN(tagStack[base + 1], valueStack[base + 1])
    if (!isFinite(percent) || percent < 0.0 || percent >= 1.0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const values = collectNumericCellRangeSeriesFromSlot(
      base,
      kindStack,
      tagStack,
      valueStack,
      rangeIndexStack,
      rangeOffsets,
      rangeLengths,
      rangeMembers,
      cellTags,
      cellNumbers,
      cellErrors,
      false,
    )
    if (values === null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        <f64>sampleCollectionErrorCode,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }
    if (values.length == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    let excluded = <i32>Math.floor(<f64>values.length * percent)
    if ((excluded & 1) == 1) {
      excluded -= 1
    }
    const retainedCount = values.length - excluded
    if (retainedCount <= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    sortNumericValues(values)
    const trimEachSide = excluded >>> 1
    let sum = 0.0
    for (let index = trimEachSide; index < values.length - trimEachSide; index += 1) {
      sum += unchecked(values[index])
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      sum / <f64>retainedCount,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  return -1
}
