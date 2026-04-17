#!/usr/bin/env bun

import { spawn } from 'node:child_process'
import { fileURLToPath } from 'node:url'

export interface PerfSmokeBenchmarkResult {
  readonly elapsedMs: number
  readonly downstreamCount: number
  readonly metrics: {
    readonly dirtyFormulaCount: number
    readonly wasmFormulaCount: number
  }
}

export interface PerfSmokeDependencies {
  readonly runBenchmark: (downstreamCount?: number) => Promise<PerfSmokeBenchmarkResult>
  readonly buildWasm: () => Promise<void>
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isPerfSmokeBenchmarkResult(value: unknown): value is PerfSmokeBenchmarkResult {
  if (!isRecord(value)) {
    return false
  }
  const metrics = value['metrics']
  return (
    typeof value['elapsedMs'] === 'number' &&
    typeof value['downstreamCount'] === 'number' &&
    isRecord(metrics) &&
    typeof metrics['dirtyFormulaCount'] === 'number' &&
    typeof metrics['wasmFormulaCount'] === 'number'
  )
}

function benchmarkEditScriptPath(): string {
  return fileURLToPath(new URL('../packages/benchmarks/src/benchmark-edit.ts', import.meta.url))
}

async function spawnCommand(command: string, args: readonly string[]): Promise<string> {
  return await new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: process.cwd(),
      env: process.env,
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    let stdout = ''
    let stderr = ''

    child.stdout.on('data', (chunk) => {
      stdout += chunk.toString()
    })
    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString()
    })
    child.on('error', reject)
    child.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(stderr || stdout || `${command} exited with code ${String(code)}`))
        return
      }
      resolve(stdout)
    })
  })
}

export async function buildWasmKernelForPerfSmoke(): Promise<void> {
  await spawnCommand('pnpm', ['wasm:build'])
}

export async function runPerfSmokeBenchmark(downstreamCount = 1_000): Promise<PerfSmokeBenchmarkResult> {
  const stdout = await spawnCommand('node', ['--import', 'tsx', benchmarkEditScriptPath(), String(downstreamCount)])

  try {
    const parsed = JSON.parse(stdout) as unknown
    if (!isPerfSmokeBenchmarkResult(parsed)) {
      throw new Error('Perf smoke benchmark output did not match the expected shape')
    }
    return parsed
  } catch (error) {
    throw new Error(`Failed to parse perf smoke benchmark output: ${error instanceof Error ? error.message : String(error)}`, {
      cause: error,
    })
  }
}

export async function runPerfSmokeGate(
  downstreamCount = 1_000,
  dependencies: PerfSmokeDependencies = {
    runBenchmark: runPerfSmokeBenchmark,
    buildWasm: buildWasmKernelForPerfSmoke,
  },
): Promise<PerfSmokeBenchmarkResult> {
  const firstPass = await dependencies.runBenchmark(downstreamCount)
  if (firstPass.metrics.wasmFormulaCount > 0) {
    return firstPass
  }
  await dependencies.buildWasm()
  return await dependencies.runBenchmark(downstreamCount)
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const { elapsedMs: elapsed, metrics, downstreamCount } = await runPerfSmokeGate()

  if (elapsed > 250) {
    console.warn(`perf smoke exceeded threshold: ${elapsed.toFixed(2)}ms`)
    process.exit(1)
  }

  if (metrics.dirtyFormulaCount < downstreamCount) {
    console.warn(
      `perf smoke failed to mark the expected downstream formulas dirty: expected at least ${downstreamCount}, got ${metrics.dirtyFormulaCount}`,
    )
    process.exit(1)
  }

  if (metrics.wasmFormulaCount === 0) {
    console.warn('perf smoke did not exercise the wasm fast path')
    process.exit(1)
  }

  console.log(JSON.stringify({ elapsedMs: elapsed, downstreamCount, metrics }, null, 2))
}
