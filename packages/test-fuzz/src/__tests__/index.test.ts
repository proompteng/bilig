import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import * as fc from 'fast-check'
import type { RunDetails } from 'fast-check'
import { afterEach, describe, expect, it } from 'vitest'
import { captureCounterexample, extractReplayPathForTest, resolveFuzzCaptureEnabled, runProperty } from '../index.js'

const tempDirs: string[] = []

afterEach(() => {
  while (tempDirs.length > 0) {
    const tempDir = tempDirs.pop()
    if (tempDir) {
      rmSync(tempDir, { force: true, recursive: true })
    }
  }
})

function withTempCwd<T>(run: () => T): T {
  const previousCwd = process.cwd()
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-test-fuzz-'))
  tempDirs.push(tempDir)
  process.chdir(tempDir)
  try {
    return run()
  } finally {
    process.chdir(previousCwd)
  }
}

async function withEnv<T>(updates: Record<string, string | undefined>, run: () => Promise<T>): Promise<T> {
  const previous = new Map<string, string | undefined>()
  for (const key of Object.keys(updates)) {
    previous.set(key, process.env[key])
    const value = updates[key]
    if (value === undefined) {
      delete process.env[key]
    } else {
      process.env[key] = value
    }
  }
  try {
    return await run()
  } finally {
    for (const [key, value] of previous) {
      if (value === undefined) {
        delete process.env[key]
      } else {
        process.env[key] = value
      }
    }
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function createRunDetails(seed: number, counterexamplePath: string, counterexample: unknown[]): RunDetails<unknown[]> {
  return {
    failed: true,
    interrupted: false,
    numRuns: 1,
    numSkips: 0,
    numShrinks: 0,
    seed,
    counterexample,
    errorInstance: null,
    counterexamplePath,
    failures: [],
    executionSummary: [],
    verbose: 0,
    runConfiguration: {},
  }
}

function readReplayPath(artifactPath: string): string | null {
  const raw = JSON.parse(readFileSync(artifactPath, 'utf8')) as unknown
  if (!isRecord(raw)) {
    return null
  }
  return typeof raw['replayPath'] === 'string' ? raw['replayPath'] : null
}

describe('test-fuzz replay-path extraction', () => {
  it('extracts replayPath from string counterexamples', () => {
    expect(extractReplayPathForTest(['cmd replayPath="0:1:2"'])).toBe('0:1:2')
  })

  it('extracts replayPath from nested command objects without string coercion', () => {
    const artifactPath = withTempCwd(() =>
      captureCounterexample({
        suite: 'grid/replay-path',
        kind: 'browser',
        details: createRunDetails(123, '0:0', [[{ kind: 'replay', replayPath: '/tmp/replay.json' }]]),
      }),
    )

    expect(readReplayPath(artifactPath)).toBe('/tmp/replay.json')
  })

  it('extracts replayPath from stringified counterexample fragments', () => {
    const artifactPath = withTempCwd(() =>
      captureCounterexample({
        suite: 'grid/replay-string',
        kind: 'browser',
        details: createRunDetails(456, '0:1', ['Command(replayPath="/tmp/from-string.json")']),
      }),
    )

    expect(readReplayPath(artifactPath)).toBe('/tmp/from-string.json')
  })

  it('resolves fuzz capture flags strictly', () => {
    expect(resolveFuzzCaptureEnabled({})).toBe(false)
    expect(resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: '1' })).toBe(true)
    expect(resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: 'true' })).toBe(true)
    expect(resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: '0' })).toBe(false)
    expect(resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: 'false' })).toBe(false)
    expect(() => resolveFuzzCaptureEnabled({ BILIG_FUZZ_CAPTURE: 'yes' })).toThrow(
      'BILIG_FUZZ_CAPTURE must be "1", "true", "0", or "false" when set, got yes',
    )
  })

  it('fails malformed capture config before reporting a fuzz failure', async () => {
    await expect(
      withEnv({ BILIG_FUZZ_CAPTURE: 'yes' }, () =>
        runProperty({
          suite: 'capture/malformed-env',
          arbitrary: fc.constant('value'),
          predicate: () => {
            throw new Error('predicate failure')
          },
          parameters: { numRuns: 1 },
        }),
      ),
    ).rejects.toThrow('BILIG_FUZZ_CAPTURE must be "1", "true", "0", or "false" when set, got yes')
  })
})
