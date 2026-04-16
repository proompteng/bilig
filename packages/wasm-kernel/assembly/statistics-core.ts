export function meanOf(values: Array<f64>): f64 {
  let sum = 0.0
  for (let index = 0; index < values.length; index++) {
    sum += unchecked(values[index])
  }
  return values.length == 0 ? NaN : sum / <f64>values.length
}

export function sampleVarianceOf(values: Array<f64>): f64 {
  if (values.length < 2) {
    return NaN
  }
  const mean = meanOf(values)
  let squared = 0.0
  for (let index = 0; index < values.length; index++) {
    const deviation = unchecked(values[index]) - mean
    squared += deviation * deviation
  }
  return squared / <f64>(values.length - 1)
}

export function populationVarianceOf(values: Array<f64>): f64 {
  if (values.length == 0) {
    return NaN
  }
  const mean = meanOf(values)
  let squared = 0.0
  for (let index = 0; index < values.length; index++) {
    const deviation = unchecked(values[index]) - mean
    squared += deviation * deviation
  }
  return squared / <f64>values.length
}

export function modeSingleOf(values: Array<f64>): f64 {
  let bestValue = 0.0
  let bestCount = 1
  let found = false
  const unique = new Array<f64>()
  const counts = new Array<i32>()
  for (let index = 0; index < values.length; index++) {
    const value = unchecked(values[index])
    let match = -1
    for (let cursor = 0; cursor < unique.length; cursor++) {
      if (unchecked(unique[cursor]) == value) {
        match = cursor
        break
      }
    }
    if (match >= 0) {
      counts[match] = unchecked(counts[match]) + 1
    } else {
      unique.push(value)
      counts.push(1)
      match = unique.length - 1
    }
    const count = unchecked(counts[match])
    if (count > bestCount || (count == bestCount && (!found || value < bestValue))) {
      bestCount = count
      bestValue = value
      found = true
    }
  }
  return bestCount >= 2 && found ? bestValue : NaN
}

export function skewSampleOf(values: Array<f64>): f64 {
  if (values.length < 3) {
    return NaN
  }
  const mean = meanOf(values)
  const stddev = Math.sqrt(sampleVarianceOf(values))
  if (!(stddev > 0.0)) {
    return NaN
  }
  let moment3 = 0.0
  for (let index = 0; index < values.length; index++) {
    const deviation = unchecked(values[index]) - mean
    moment3 += deviation * deviation * deviation
  }
  const n = <f64>values.length
  return (n * moment3) / ((n - 1.0) * (n - 2.0) * stddev * stddev * stddev)
}

export function skewPopulationOf(values: Array<f64>): f64 {
  if (values.length == 0) {
    return NaN
  }
  const mean = meanOf(values)
  const stddev = Math.sqrt(populationVarianceOf(values))
  if (!(stddev > 0.0)) {
    return NaN
  }
  let moment3 = 0.0
  for (let index = 0; index < values.length; index++) {
    const deviation = unchecked(values[index]) - mean
    moment3 += deviation * deviation * deviation
  }
  return moment3 / <f64>values.length / (stddev * stddev * stddev)
}

export function kurtosisOf(values: Array<f64>): f64 {
  if (values.length < 4) {
    return NaN
  }
  const mean = meanOf(values)
  const stddev = Math.sqrt(sampleVarianceOf(values))
  if (!(stddev > 0.0)) {
    return NaN
  }
  let sum4 = 0.0
  for (let index = 0; index < values.length; index++) {
    const standardized = (unchecked(values[index]) - mean) / stddev
    sum4 += standardized * standardized * standardized * standardized
  }
  const n = <f64>values.length
  return (n * (n + 1.0) * sum4) / ((n - 1.0) * (n - 2.0) * (n - 3.0)) - (3.0 * (n - 1.0) * (n - 1.0)) / ((n - 2.0) * (n - 3.0))
}

export function sortNumericValues(values: Array<f64>, left: i32 = 0, right: i32 = -1): void {
  if (right < 0) {
    right = values.length - 1
  }
  if (left >= right) {
    return
  }
  let low = left
  let high = right
  const pivot = unchecked(values[(left + right) >>> 1])
  while (low <= high) {
    while (unchecked(values[low]) < pivot) {
      low += 1
    }
    while (unchecked(values[high]) > pivot) {
      high -= 1
    }
    if (low <= high) {
      const swap = unchecked(values[low])
      values[low] = unchecked(values[high])
      values[high] = swap
      low += 1
      high -= 1
    }
  }
  if (left < high) {
    sortNumericValues(values, left, high)
  }
  if (low < right) {
    sortNumericValues(values, low, right)
  }
}

export function interpolateSortedPercentile(values: Array<f64>, percentile: f64, exclusive: bool): f64 {
  const count = values.length
  if (count == 0 || !isFinite(percentile)) {
    return NaN
  }
  if (exclusive) {
    if (!(percentile > 0.0 && percentile < 1.0)) {
      return NaN
    }
    const rank = percentile * (<f64>count + 1.0)
    if (rank < 1.0 || rank > <f64>count) {
      return NaN
    }
    const lowerRank = <i32>Math.floor(rank)
    const upperRank = <i32>Math.ceil(rank)
    if (lowerRank == upperRank) {
      return unchecked(values[lowerRank - 1])
    }
    const lower = unchecked(values[lowerRank - 1])
    const upper = unchecked(values[upperRank - 1])
    return lower + (rank - <f64>lowerRank) * (upper - lower)
  }

  if (percentile < 0.0 || percentile > 1.0) {
    return NaN
  }
  if (count == 1) {
    return unchecked(values[0])
  }
  const rank = percentile * <f64>(count - 1) + 1.0
  const lowerRank = <i32>Math.floor(rank)
  const upperRank = <i32>Math.ceil(rank)
  if (lowerRank == upperRank) {
    return unchecked(values[lowerRank - 1])
  }
  const lower = unchecked(values[lowerRank - 1])
  const upper = unchecked(values[upperRank - 1])
  return lower + (rank - <f64>lowerRank) * (upper - lower)
}

export function truncateToSignificance(value: f64, significance: i32): f64 {
  let scale = 1.0
  for (let index = 0; index < significance; index += 1) {
    scale *= 10.0
  }
  return Math.trunc(value * scale) / scale
}

export function interpolateSortedPercentRank(values: Array<f64>, target: f64, exclusive: bool): f64 {
  const count = values.length
  if (count < 2 || !isFinite(target)) {
    return NaN
  }

  let exactFirst = -1
  let exactLast = -1
  for (let index = 0; index < count; index += 1) {
    if (unchecked(values[index]) != target) {
      continue
    }
    if (exactFirst < 0) {
      exactFirst = index
    }
    exactLast = index
  }

  if (exactFirst >= 0) {
    const averageIndex = (<f64>exactFirst + <f64>exactLast) / 2.0
    return exclusive ? (averageIndex + 1.0) / (<f64>count + 1.0) : averageIndex / <f64>(count - 1)
  }

  if (target < unchecked(values[0]) || target > unchecked(values[count - 1])) {
    return NaN
  }

  let lowerIndex = -1
  for (let index = 0; index < count; index += 1) {
    if (unchecked(values[index]) < target) {
      lowerIndex = index
      continue
    }
    break
  }
  const upperIndex = lowerIndex + 1
  if (lowerIndex < 0 || upperIndex >= count) {
    return NaN
  }

  const lower = unchecked(values[lowerIndex])
  const upper = unchecked(values[upperIndex])
  if (upper == lower) {
    return NaN
  }

  const lowerRank = exclusive ? <f64>(lowerIndex + 1) / (<f64>count + 1.0) : <f64>lowerIndex / <f64>(count - 1)
  const upperRank = exclusive ? <f64>(upperIndex + 1) / (<f64>count + 1.0) : <f64>upperIndex / <f64>(count - 1)
  const fraction = (target - lower) / (upper - lower)
  return lowerRank + fraction * (upperRank - lowerRank)
}
