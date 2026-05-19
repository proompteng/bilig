export type FuzzMode = 'fuzz' | 'replay'

export function parseFuzzMode(value: string | undefined): FuzzMode {
  if (value === undefined || value === 'fuzz') {
    return 'fuzz'
  }
  if (value === 'replay') {
    return value
  }
  throw new Error(`Fuzz mode must be "fuzz" or "replay", got ${value}`)
}

export function resolveVitestFuzzMaxWorkers(availableWorkers: number): number {
  if (!Number.isFinite(availableWorkers) || availableWorkers <= 0) {
    return 1
  }
  return Math.max(1, Math.min(2, Math.ceil(availableWorkers / 2)))
}

export function buildVitestFuzzCommand(files: readonly string[], availableWorkers: number): string[] {
  return ['pnpm', 'exec', 'vitest', 'run', ...files, '--maxWorkers', String(resolveVitestFuzzMaxWorkers(availableWorkers))]
}
