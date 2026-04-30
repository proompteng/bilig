export interface NumericSummary {
  samples: number[]
  min: number
  median: number
  p95: number
  max: number
  mean: number
  standardDeviation: number
  relativeStandardDeviation: number
  standardError: number
  confidence95: {
    low: number
    high: number
  }
}

export function summarizeNumbers(values: readonly number[]): NumericSummary {
  if (values.length === 0) {
    throw new Error('Cannot summarize an empty sample set')
  }

  const samples = [...values].toSorted((left, right) => left - right)
  const min = samples[0]!
  const max = samples[samples.length - 1]!
  const mean = samples.reduce((sum, value) => sum + value, 0) / samples.length
  const variance = samples.reduce((sum, value) => sum + (value - mean) ** 2, 0) / samples.length
  const standardDeviation = Math.sqrt(variance)
  const standardError = standardDeviation / Math.sqrt(samples.length)
  const confidenceDelta = 1.96 * standardError

  return {
    samples,
    min,
    median: quantile(samples, 0.5),
    p95: quantile(samples, 0.95),
    max,
    mean,
    standardDeviation,
    relativeStandardDeviation: mean === 0 ? 0 : standardDeviation / Math.abs(mean),
    standardError,
    confidence95: {
      low: mean - confidenceDelta,
      high: mean + confidenceDelta,
    },
  }
}

export function quantile(sortedValues: readonly number[], percentile: number): number {
  if (sortedValues.length === 0) {
    throw new Error('Cannot compute a quantile of an empty sample set')
  }
  if (percentile <= 0) {
    return sortedValues[0]!
  }
  if (percentile >= 1) {
    return sortedValues[sortedValues.length - 1]!
  }
  const index = Math.ceil(percentile * sortedValues.length) - 1
  return sortedValues[Math.max(0, Math.min(sortedValues.length - 1, index))]!
}
