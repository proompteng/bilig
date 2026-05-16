export function parseBenchToleranceMultiplier(value: string | undefined, isCi: boolean): number {
  if (value === undefined || value.length === 0) {
    return isCi ? 1.5 : 1
  }
  if (!/^(?:[1-9]\d*|[1-9]\d*\.\d+|0\.\d+)$/u.test(value)) {
    throw new Error(`BILIG_BENCH_TOLERANCE must be a positive number, got ${value}`)
  }
  const parsed = Number(value)
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`BILIG_BENCH_TOLERANCE must be a positive finite number, got ${value}`)
  }
  return parsed
}

export function readBenchToleranceMultiplier(env: { BILIG_BENCH_TOLERANCE?: string | undefined; CI?: string | undefined }): number {
  return parseBenchToleranceMultiplier(env.BILIG_BENCH_TOLERANCE, env.CI !== undefined && env.CI.length > 0)
}
