const DIRECT_AGGREGATE_OP_SUM: u8 = 1
const DIRECT_AGGREGATE_OP_AVERAGE: u8 = 2
const DIRECT_AGGREGATE_OP_COUNT: u8 = 3
const DIRECT_AGGREGATE_OP_MIN: u8 = 4
const DIRECT_AGGREGATE_OP_MAX: u8 = 5

export function evalDenseNumericRowAggregateBatch(
  aggregateKind: u8,
  values: Float64Array,
  rowCount: i32,
  prefixColCount: i32,
  startColOffset: i32,
  aggregateColCount: i32,
  resultOffset: f64,
  outNumbers: Float64Array,
): void {
  if (rowCount <= 0 || prefixColCount <= 0 || aggregateColCount <= 0) {
    return
  }
  if (startColOffset < 0 || startColOffset + aggregateColCount > prefixColCount) {
    return
  }

  for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
    const baseOffset = rowOffset * prefixColCount + startColOffset
    let sum: f64 = 0
    let minimum: f64 = Infinity
    let maximum: f64 = -Infinity

    for (let colOffset = 0; colOffset < aggregateColCount; colOffset++) {
      const value = values[baseOffset + colOffset]
      sum += value
      if (value < minimum) {
        minimum = value
      }
      if (value > maximum) {
        maximum = value
      }
    }

    if (aggregateKind == DIRECT_AGGREGATE_OP_SUM) {
      outNumbers[rowOffset] = sum + resultOffset
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_AVERAGE) {
      outNumbers[rowOffset] = sum / aggregateColCount + resultOffset
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_COUNT) {
      outNumbers[rowOffset] = aggregateColCount + resultOffset
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_MIN) {
      outNumbers[rowOffset] = minimum + resultOffset
    } else if (aggregateKind == DIRECT_AGGREGATE_OP_MAX) {
      outNumbers[rowOffset] = maximum + resultOffset
    } else {
      outNumbers[rowOffset] = NaN
    }
  }
}
