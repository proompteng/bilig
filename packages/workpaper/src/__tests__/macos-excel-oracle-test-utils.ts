import { mkdirSync, mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const cleanupRetryMs = 100
const retriableCleanupCodes = new Set(['EBUSY', 'EINTR', 'ENOTEMPTY'])

export function createExcelAccessibleTempDir(prefix: string): string {
  const root = join(tmpdir(), 'bilig-headless-oracle')
  mkdirSync(root, { recursive: true })
  return mkdtempSync(join(root, prefix))
}

export function removeMacosExcelTestDir(dirPath: string): void {
  for (let attempt = 0; attempt < 5; attempt += 1) {
    try {
      rmSync(dirPath, { recursive: true, force: true })
      return
    } catch (error) {
      if (isRecord(error) && retriableCleanupCodes.has(String(error['code']))) {
        sleepSync(cleanupRetryMs)
        continue
      }
      throw error
    }
  }
  rmSync(dirPath, { recursive: true, force: true })
}

function sleepSync(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
