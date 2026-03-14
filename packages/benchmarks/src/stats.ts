export interface NumericSummary {
  samples: number[];
  min: number;
  median: number;
  p95: number;
  max: number;
  mean: number;
}

export function summarizeNumbers(values: readonly number[]): NumericSummary {
  if (values.length === 0) {
    throw new Error("Cannot summarize an empty sample set");
  }

  const samples = [...values].sort((left, right) => left - right);
  const min = samples[0]!;
  const max = samples[samples.length - 1]!;
  const mean = samples.reduce((sum, value) => sum + value, 0) / samples.length;

  return {
    samples,
    min,
    median: quantile(samples, 0.5),
    p95: quantile(samples, 0.95),
    max,
    mean
  };
}

export function quantile(sortedValues: readonly number[], percentile: number): number {
  if (sortedValues.length === 0) {
    throw new Error("Cannot compute a quantile of an empty sample set");
  }
  if (percentile <= 0) {
    return sortedValues[0]!;
  }
  if (percentile >= 1) {
    return sortedValues[sortedValues.length - 1]!;
  }
  const index = Math.ceil(percentile * sortedValues.length) - 1;
  return sortedValues[Math.max(0, Math.min(sortedValues.length - 1, index))]!;
}
