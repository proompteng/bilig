import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { inputCellNumeric, inputCellScalarValue, inputCellTag, inputColsFromSlot, inputRowsFromSlot, toNumberOrNaN } from './operands'
import { rangeSupportedScalarOnly } from './builtin-args'
import { STACK_KIND_SCALAR, writeArrayResult, writeResult } from './result-io'
import { allocateSpillArrayResult, writeSpillArrayNumber, writeSpillArrayValue } from './vm'
import {
  collectPairedNumericStats,
  pairedCenteredCrossProducts,
  pairedCenteredSumSquaresX,
  pairedCenteredSumSquaresY,
  pairedSampleCount,
  pairedSumX,
  pairedSumY,
} from './statistics-tests'

export function tryApplyRegressionBuiltin(
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
  if (
    (builtinId == BuiltinId.Correl ||
      builtinId == BuiltinId.Covar ||
      builtinId == BuiltinId.Pearson ||
      builtinId == BuiltinId.CovarianceP ||
      builtinId == BuiltinId.CovarianceS ||
      builtinId == BuiltinId.Intercept ||
      builtinId == BuiltinId.Rsq ||
      builtinId == BuiltinId.Slope ||
      builtinId == BuiltinId.Steyx) &&
    argc == 2
  ) {
    const statsError = collectPairedNumericStats(
      base,
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
    if (statsError != 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, <f64>statsError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const centeredSumSquaresX = pairedCenteredSumSquaresX()
    const centeredSumSquaresY = pairedCenteredSumSquaresY()
    const centeredCrossProducts = pairedCenteredCrossProducts()

    if (builtinId == BuiltinId.Covar || builtinId == BuiltinId.CovarianceP) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        centeredCrossProducts / <f64>pairedSampleCount,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    if (builtinId == BuiltinId.CovarianceS) {
      if (pairedSampleCount < 2) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        centeredCrossProducts / <f64>(pairedSampleCount - 1),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    if (builtinId == BuiltinId.Correl || builtinId == BuiltinId.Pearson || builtinId == BuiltinId.Rsq) {
      const denominator = Math.sqrt(centeredSumSquaresX * centeredSumSquaresY)
      if (pairedSampleCount < 2 || denominator <= 0 || !isFinite(denominator)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const correlation = centeredCrossProducts / denominator
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        builtinId == BuiltinId.Rsq ? correlation * correlation : correlation,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    if (centeredSumSquaresX == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const slope = centeredCrossProducts / centeredSumSquaresX
    if (builtinId == BuiltinId.Slope) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, slope, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const intercept = pairedSampleCount == 0 ? NaN : (pairedSumY - slope * pairedSumX) / <f64>pairedSampleCount
    if (builtinId == BuiltinId.Intercept) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, intercept, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    if (builtinId == BuiltinId.Steyx) {
      if (pairedSampleCount <= 2) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const residualSumSquares = max<f64>(0, centeredSumSquaresY - slope * centeredCrossProducts)
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Number,
        Math.sqrt(residualSumSquares / <f64>(pairedSampleCount - 2)),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      )
    }

    return -1
  }

  if (builtinId == BuiltinId.Forecast && argc == 3) {
    if (kindStack[base] != STACK_KIND_SCALAR) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const targetX = toNumberOrNaN(tagStack[base], valueStack[base])
    if (!isFinite(targetX)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const statsError = collectPairedNumericStats(
      base + 1,
      base + 2,
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
    if (statsError != 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, <f64>statsError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const centeredSumSquaresX = pairedCenteredSumSquaresX()
    if (centeredSumSquaresX == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const slope = pairedCenteredCrossProducts() / centeredSumSquaresX
    const intercept = pairedSampleCount == 0 ? NaN : (pairedSumY - slope * pairedSumX) / <f64>pairedSampleCount
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      intercept + slope * targetX,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if ((builtinId == BuiltinId.Trend || builtinId == BuiltinId.Growth) && argc >= 1 && argc <= 4) {
    let includeIntercept = true
    if (argc == 4) {
      if (kindStack[base + 3] != STACK_KIND_SCALAR) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[base + 3] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[base + 3],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }
      if (tagStack[base + 3] == ValueTag.Boolean || tagStack[base + 3] == ValueTag.Number) {
        includeIntercept = valueStack[base + 3] != 0
      } else if (tagStack[base + 3] == ValueTag.Empty) {
        includeIntercept = false
      } else {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    const knownYRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const knownYCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    const sampleCount = knownYRows * knownYCols
    if (knownYRows < 1 || knownYCols < 1 || sampleCount < 1) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const knownYValues = new Array<f64>()
    for (let row = 0; row < knownYRows; row += 1) {
      for (let col = 0; col < knownYCols; col += 1) {
        const yTag = inputCellTag(
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
        const yRaw = inputCellScalarValue(
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
        if (yTag == ValueTag.Error) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, yRaw, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        const numeric = toNumberOrNaN(yTag, yRaw)
        if (!isFinite(numeric) || (builtinId == BuiltinId.Growth && numeric <= 0.0)) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        knownYValues.push(builtinId == BuiltinId.Growth ? Math.log(numeric) : numeric)
      }
    }

    let knownXRows = knownYRows
    let knownXCols = knownYCols
    const knownXValues = new Array<f64>()
    if (argc >= 2) {
      knownXRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts)
      knownXCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts)
      if (knownXRows < 1 || knownXCols < 1 || knownXRows * knownXCols != sampleCount) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      for (let row = 0; row < knownXRows; row += 1) {
        for (let col = 0; col < knownXCols; col += 1) {
          const numeric = inputCellNumeric(
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
          if (!isFinite(numeric)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            )
          }
          knownXValues.push(numeric)
        }
      }
    } else {
      for (let index = 0; index < sampleCount; index += 1) {
        knownXValues.push(<f64>(index + 1))
      }
    }

    let slope = 0.0
    let intercept = 0.0
    if (includeIntercept) {
      let sumX = 0.0
      let sumY = 0.0
      for (let index = 0; index < sampleCount; index += 1) {
        sumX += unchecked(knownXValues[index])
        sumY += unchecked(knownYValues[index])
      }
      const meanX = sumX / <f64>sampleCount
      const meanY = sumY / <f64>sampleCount
      let sumSquaresX = 0.0
      let sumCrossProducts = 0.0
      for (let index = 0; index < sampleCount; index += 1) {
        const xDeviation = unchecked(knownXValues[index]) - meanX
        const yDeviation = unchecked(knownYValues[index]) - meanY
        sumSquaresX += xDeviation * xDeviation
        sumCrossProducts += xDeviation * yDeviation
      }
      if (sumSquaresX == 0.0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      slope = sumCrossProducts / sumSquaresX
      intercept = meanY - slope * meanX
    } else {
      let sumSquaresX = 0.0
      let sumCrossProducts = 0.0
      for (let index = 0; index < sampleCount; index += 1) {
        const xValue = unchecked(knownXValues[index])
        const yValue = unchecked(knownYValues[index])
        sumSquaresX += xValue * xValue
        sumCrossProducts += xValue * yValue
      }
      if (sumSquaresX == 0.0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      slope = sumCrossProducts / sumSquaresX
    }

    let predictionRows = knownYRows
    let predictionCols = knownYCols
    if (argc >= 3) {
      predictionRows = inputRowsFromSlot(base + 2, kindStack, rangeIndexStack, rangeRowCounts)
      predictionCols = inputColsFromSlot(base + 2, kindStack, rangeIndexStack, rangeColCounts)
      if (predictionRows < 1 || predictionCols < 1) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    } else if (argc >= 2) {
      predictionRows = knownXRows
      predictionCols = knownXCols
    }

    const predictionCount = predictionRows * predictionCols
    if (predictionCount == 1) {
      let predictionX = argc >= 2 ? unchecked(knownXValues[0]) : 1.0
      if (argc >= 3) {
        predictionX = inputCellNumeric(
          base + 2,
          0,
          0,
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
        if (!isFinite(predictionX)) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
      }
      let result = intercept + slope * predictionX
      if (builtinId == BuiltinId.Growth) {
        result = Math.exp(result)
      }
      if (!isFinite(result)) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, result, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const arrayIndex = allocateSpillArrayResult(predictionRows, predictionCols)
    if (argc >= 3) {
      let cursor = 0
      for (let row = 0; row < predictionRows; row += 1) {
        for (let col = 0; col < predictionCols; col += 1) {
          const predictionX = inputCellNumeric(
            base + 2,
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
          if (!isFinite(predictionX)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            )
          }
          let result = intercept + slope * predictionX
          if (builtinId == BuiltinId.Growth) {
            result = Math.exp(result)
          }
          if (!isFinite(result)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            )
          }
          writeSpillArrayNumber(arrayIndex, cursor, result)
          cursor += 1
        }
      }
    } else {
      for (let index = 0; index < predictionCount; index += 1) {
        let result = intercept + slope * unchecked(knownXValues[index])
        if (builtinId == BuiltinId.Growth) {
          result = Math.exp(result)
        }
        if (!isFinite(result)) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        writeSpillArrayNumber(arrayIndex, index, result)
      }
    }

    return writeArrayResult(base, arrayIndex, predictionRows, predictionCols, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Linest || builtinId == BuiltinId.Logest) && argc >= 1 && argc <= 4) {
    let includeIntercept = true
    let includeStats = false
    if (argc >= 3) {
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
      if (tagStack[base + 2] == ValueTag.Boolean || tagStack[base + 2] == ValueTag.Number) {
        includeIntercept = valueStack[base + 2] != 0
      } else if (tagStack[base + 2] == ValueTag.Empty) {
        includeIntercept = false
      } else {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }
    if (argc == 4) {
      if (kindStack[base + 3] != STACK_KIND_SCALAR) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (tagStack[base + 3] == ValueTag.Error) {
        return writeResult(
          base,
          STACK_KIND_SCALAR,
          <u8>ValueTag.Error,
          valueStack[base + 3],
          rangeIndexStack,
          valueStack,
          tagStack,
          kindStack,
        )
      }
      if (tagStack[base + 3] == ValueTag.Boolean || tagStack[base + 3] == ValueTag.Number) {
        includeStats = valueStack[base + 3] != 0
      } else if (tagStack[base + 3] == ValueTag.Empty) {
        includeStats = false
      } else {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    const knownYRows = inputRowsFromSlot(base, kindStack, rangeIndexStack, rangeRowCounts)
    const knownYCols = inputColsFromSlot(base, kindStack, rangeIndexStack, rangeColCounts)
    const sampleCount = knownYRows * knownYCols
    if (sampleCount < 1) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const knownYValues = new Array<f64>()
    for (let row = 0; row < knownYRows; row += 1) {
      for (let col = 0; col < knownYCols; col += 1) {
        const yTag = inputCellTag(
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
        const yRaw = inputCellScalarValue(
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
        if (yTag == ValueTag.Error) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, yRaw, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        const numeric = toNumberOrNaN(yTag, yRaw)
        if (!isFinite(numeric) || (builtinId == BuiltinId.Logest && numeric <= 0.0)) {
          return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
        }
        knownYValues.push(builtinId == BuiltinId.Logest ? Math.log(numeric) : numeric)
      }
    }

    const knownXValues = new Array<f64>()
    if (argc >= 2) {
      const knownXRows = inputRowsFromSlot(base + 1, kindStack, rangeIndexStack, rangeRowCounts)
      const knownXCols = inputColsFromSlot(base + 1, kindStack, rangeIndexStack, rangeColCounts)
      if (knownXRows * knownXCols != sampleCount) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      for (let row = 0; row < knownXRows; row += 1) {
        for (let col = 0; col < knownXCols; col += 1) {
          const numeric = inputCellNumeric(
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
          if (!isFinite(numeric)) {
            return writeResult(
              base,
              STACK_KIND_SCALAR,
              <u8>ValueTag.Error,
              ErrorCode.Value,
              rangeIndexStack,
              valueStack,
              tagStack,
              kindStack,
            )
          }
          knownXValues.push(numeric)
        }
      }
    } else {
      for (let index = 0; index < sampleCount; index += 1) {
        knownXValues.push(<f64>(index + 1))
      }
    }

    let sumX = 0.0
    let sumY = 0.0
    for (let index = 0; index < sampleCount; index += 1) {
      sumX += unchecked(knownXValues[index])
      sumY += unchecked(knownYValues[index])
    }

    let slope = 0.0
    let intercept = 0.0
    let totalSumSquares = 0.0
    let sumSquaresX = 0.0
    let sumCrossProducts = 0.0
    if (includeIntercept) {
      const meanX = sumX / <f64>sampleCount
      const meanY = sumY / <f64>sampleCount
      for (let index = 0; index < sampleCount; index += 1) {
        const xDeviation = unchecked(knownXValues[index]) - meanX
        const yDeviation = unchecked(knownYValues[index]) - meanY
        sumSquaresX += xDeviation * xDeviation
        sumCrossProducts += xDeviation * yDeviation
        totalSumSquares += yDeviation * yDeviation
      }
      if (sumSquaresX == 0.0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      slope = sumCrossProducts / sumSquaresX
      intercept = meanY - slope * meanX
    } else {
      for (let index = 0; index < sampleCount; index += 1) {
        const xValue = unchecked(knownXValues[index])
        const yValue = unchecked(knownYValues[index])
        sumSquaresX += xValue * xValue
        sumCrossProducts += xValue * yValue
        totalSumSquares += yValue * yValue
      }
      if (sumSquaresX == 0.0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      slope = sumCrossProducts / sumSquaresX
    }

    let residualSumSquares = 0.0
    for (let index = 0; index < sampleCount; index += 1) {
      const residual = unchecked(knownYValues[index]) - (intercept + slope * unchecked(knownXValues[index]))
      residualSumSquares += residual * residual
    }
    residualSumSquares = max<f64>(0.0, residualSumSquares)
    const regressionSumSquares = max<f64>(0.0, totalSumSquares - residualSumSquares)

    const leading = builtinId == BuiltinId.Logest ? Math.exp(slope) : slope
    const trailing = builtinId == BuiltinId.Logest ? (includeIntercept ? Math.exp(intercept) : 1.0) : includeIntercept ? intercept : 0.0
    if (!isFinite(leading) || !isFinite(trailing)) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    const resultRows = includeStats ? 5 : 1
    const arrayIndex = allocateSpillArrayResult(resultRows, 2)
    writeSpillArrayNumber(arrayIndex, 0, leading)
    writeSpillArrayNumber(arrayIndex, 1, trailing)

    if (includeStats) {
      const degreesFreedom = sampleCount - (includeIntercept ? 2 : 1)
      let slopeStandardError = NaN
      let interceptStandardError = NaN
      let rSquared = NaN
      let standardErrorY = NaN
      let fStatistic = NaN

      if (degreesFreedom > 0) {
        const meanSquaredError = residualSumSquares / <f64>degreesFreedom
        standardErrorY = Math.sqrt(meanSquaredError)
        if (includeIntercept) {
          const meanX = sumX / <f64>sampleCount
          if (sumSquaresX > 0.0) {
            slopeStandardError = Math.sqrt(meanSquaredError / sumSquaresX)
            interceptStandardError = Math.sqrt(meanSquaredError * (1.0 / <f64>sampleCount + (meanX * meanX) / sumSquaresX))
          }
        } else if (sumSquaresX > 0.0) {
          slopeStandardError = Math.sqrt(meanSquaredError / sumSquaresX)
          interceptStandardError = 0.0
        }
        if (residualSumSquares == 0.0) {
          fStatistic = Infinity
        } else {
          fStatistic = regressionSumSquares / (residualSumSquares / <f64>degreesFreedom)
        }
      }

      if (totalSumSquares == 0.0) {
        rSquared = residualSumSquares == 0.0 ? 1.0 : NaN
      } else {
        rSquared = 1.0 - residualSumSquares / totalSumSquares
      }

      if (isFinite(slopeStandardError)) {
        writeSpillArrayNumber(arrayIndex, 2, slopeStandardError)
      } else {
        writeSpillArrayValue(arrayIndex, 2, <u8>ValueTag.Error, ErrorCode.Div0)
      }
      if (isFinite(interceptStandardError)) {
        writeSpillArrayNumber(arrayIndex, 3, interceptStandardError)
      } else {
        writeSpillArrayValue(arrayIndex, 3, <u8>ValueTag.Error, ErrorCode.Div0)
      }
      if (isFinite(rSquared)) {
        writeSpillArrayNumber(arrayIndex, 4, rSquared)
      } else {
        writeSpillArrayValue(arrayIndex, 4, <u8>ValueTag.Error, ErrorCode.Div0)
      }
      if (isFinite(standardErrorY)) {
        writeSpillArrayNumber(arrayIndex, 5, standardErrorY)
      } else {
        writeSpillArrayValue(arrayIndex, 5, <u8>ValueTag.Error, ErrorCode.Div0)
      }
      if (isFinite(fStatistic)) {
        writeSpillArrayNumber(arrayIndex, 6, fStatistic)
      } else {
        writeSpillArrayValue(arrayIndex, 6, <u8>ValueTag.Error, ErrorCode.Div0)
      }
      if (degreesFreedom > 0) {
        writeSpillArrayNumber(arrayIndex, 7, <f64>degreesFreedom)
      } else {
        writeSpillArrayValue(arrayIndex, 7, <u8>ValueTag.Error, ErrorCode.Div0)
      }
      writeSpillArrayNumber(arrayIndex, 8, regressionSumSquares)
      writeSpillArrayNumber(arrayIndex, 9, residualSumSquares)
    }

    return writeArrayResult(base, arrayIndex, resultRows, 2, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
