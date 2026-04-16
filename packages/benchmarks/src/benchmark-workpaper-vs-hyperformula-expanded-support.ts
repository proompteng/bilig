import { performance } from 'node:perf_hooks'
import { WorkPaper } from '@bilig/headless'
import { ValueTag } from '@bilig/protocol'
import type { RawCellContent as HyperFormulaRawCellContent, Sheet as HyperFormulaSheet } from 'hyperformula'
import { HYPERFORMULA_LICENSE_KEY } from './benchmark-workpaper-vs-hyperformula.js'
import type { MemoryMeasurement } from './metrics.js'
import { measureMemory, sampleMemory } from './metrics.js'

export const { HyperFormula } = await import('hyperformula')
export type HyperFormulaInstance = ReturnType<typeof HyperFormula.buildFromSheets>

export interface BenchmarkSample {
  elapsedMs: number
  memory: MemoryMeasurement
  verification: Record<string, unknown>
}

export function measureWorkPaperBuildFromSheets(
  sheets: Record<string, readonly (readonly (boolean | number | string | null)[])[]>,
  verification: (workbook: WorkPaper) => Record<string, unknown>,
  config?: Parameters<typeof WorkPaper.buildFromSheets>[1],
): BenchmarkSample {
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const workbook = WorkPaper.buildFromSheets(sheets, config)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const result = verification(workbook)
  workbook.dispose()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: result,
  }
}

export function measureHyperFormulaBuildFromSheets(
  sheets: Record<string, HyperFormulaSheet>,
  verification: (workbook: HyperFormulaInstance) => Record<string, unknown>,
  config?: Record<string, unknown>,
): BenchmarkSample {
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const workbook = HyperFormula.buildFromSheets(
    sheets,
    config
      ? {
          licenseKey: HYPERFORMULA_LICENSE_KEY,
          ...config,
        }
      : {
          licenseKey: HYPERFORMULA_LICENSE_KEY,
        },
  )
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const result = verification(workbook)
  workbook.destroy()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: result,
  }
}

export function measureMutationSample<Result>(
  workbook: WorkPaper,
  execute: () => Result,
  verification: (result: Result) => Record<string, unknown>,
): BenchmarkSample {
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const result = execute()
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const resolvedVerification = verification(result)
  workbook.dispose()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: resolvedVerification,
  }
}

export function measureHyperFormulaMutationSample<Result>(
  workbook: HyperFormulaInstance,
  execute: () => Result,
  verification: (result: Result) => Record<string, unknown>,
): BenchmarkSample {
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const result = execute()
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const resolvedVerification = verification(result)
  workbook.destroy()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: resolvedVerification,
  }
}

export function normalizeWorkPaperValue(value: unknown): boolean | number | string | null | { error: unknown } {
  if (!isProtocolValueLike(value)) {
    return null
  }

  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value ?? null
    case ValueTag.Error:
      return { error: value.code ?? 'ERROR' }
    default:
      return { error: `UNKNOWN_TAG_${String(value.tag)}` }
  }
}

export function normalizeHyperFormulaValue(value: unknown): boolean | number | string | null | { error: unknown } {
  if (value === null || typeof value === 'boolean' || typeof value === 'number' || typeof value === 'string') {
    return value
  }
  if (isHyperFormulaErrorLike(value)) {
    return { error: value.value }
  }
  return { error: 'UNKNOWN_VALUE' }
}

export function toHyperFormulaSheet(sheet: ReadonlyArray<ReadonlyArray<unknown>>): HyperFormulaSheet {
  return sheet.map((row) => row.map((cell) => toHyperFormulaCell(cell)))
}

function toHyperFormulaCell(cell: unknown): HyperFormulaRawCellContent {
  if (cell === null || typeof cell === 'boolean' || typeof cell === 'number' || typeof cell === 'string') {
    return cell
  }
  throw new Error(`Unsupported HyperFormula benchmark cell type: ${typeof cell}`)
}

function isProtocolValueLike(value: unknown): value is { code?: unknown; tag: ValueTag; value?: boolean | number | string } {
  if (!value || typeof value !== 'object') {
    return false
  }
  const tag = Reflect.get(value, 'tag')
  return tag === ValueTag.Empty || tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.String || tag === ValueTag.Error
}

function isHyperFormulaErrorLike(value: unknown): value is { value: unknown } {
  return value !== null && typeof value === 'object' && 'value' in value
}
