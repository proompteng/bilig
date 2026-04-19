export function resolveVitestFuzzMaxWorkers(availableWorkers: number): number {
  if (!Number.isFinite(availableWorkers) || availableWorkers <= 0) {
    return 1
  }
  return Math.max(1, Math.min(4, Math.ceil(availableWorkers / 2)))
}

export function buildVitestFuzzCommand(files: readonly string[], availableWorkers: number): string[] {
  return ['pnpm', 'exec', 'vitest', 'run', ...files, '--maxWorkers', String(resolveVitestFuzzMaxWorkers(availableWorkers))]
}
