import { ValueTag, type CellValue } from "@bilig/protocol";

export function sampleVariance(numbers: readonly number[]): number {
  if (numbers.length <= 1) {
    return 0;
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const squared = numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0);
  return squared / (numbers.length - 1);
}

export function populationVariance(numbers: readonly number[]): number {
  if (numbers.length === 0) {
    return 0;
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const squared = numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0);
  return squared / numbers.length;
}

export function sampleStandardDeviation(numbers: readonly number[]): number {
  const variance = sampleVariance(numbers);
  return variance < 0 ? Number.NaN : Math.sqrt(variance);
}

export function populationStandardDeviation(numbers: readonly number[]): number {
  const variance = populationVariance(numbers);
  return variance < 0 ? Number.NaN : Math.sqrt(variance);
}

export function collectAStyleNumericArgs(args: readonly CellValue[]): number[] {
  const values: number[] = [];
  for (const arg of args) {
    switch (arg.tag) {
      case ValueTag.Number:
        values.push(arg.value);
        break;
      case ValueTag.Boolean:
        values.push(arg.value ? 1 : 0);
        break;
      case ValueTag.String:
        values.push(0);
        break;
      case ValueTag.Empty:
      case ValueTag.Error:
        break;
    }
  }
  return values;
}

export function modeSingle(numbers: readonly number[]): number | undefined {
  const counts = new Map<number, number>();
  let bestValue: number | undefined;
  let bestCount = 1;
  for (const value of numbers) {
    const count = (counts.get(value) ?? 0) + 1;
    counts.set(value, count);
    if (
      count > bestCount ||
      (count === bestCount && bestValue !== undefined && value < bestValue)
    ) {
      bestCount = count;
      bestValue = value;
    }
    if (count > bestCount && bestValue === undefined) {
      bestValue = value;
    }
  }
  return bestCount >= 2 ? bestValue : undefined;
}

export function erfApprox(value: number): number {
  const sign = value < 0 ? -1 : 1;
  const absolute = Math.abs(value);
  const t = 1 / (1 + 0.3275911 * absolute);
  const y =
    1 -
    ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) *
      t *
      Math.exp(-absolute * absolute);
  return sign * y;
}
