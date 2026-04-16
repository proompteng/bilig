import { performance } from 'node:perf_hooks'
import { SpreadsheetEngine } from '@bilig/core'
import { seedLoadWorkbook } from './generate-workbook.js'
import {
  buildWorkbookBenchmarkCorpus,
  isWorkbookBenchmarkCorpusId,
  type WorkbookBenchmarkCorpusFamily,
  type WorkbookBenchmarkCorpusId,
} from './workbook-corpus.js'
import { measureMemory, sampleMemory, type MemoryMeasurement } from './metrics.js'

export interface LoadBenchmarkResult {
  scenario: 'load'
  materializedCells: number
  corpusCaseId: WorkbookBenchmarkCorpusId | null
  corpusFamily: WorkbookBenchmarkCorpusFamily | null
  elapsedMs: number
  memory: MemoryMeasurement
}

function resolveLoadBenchmarkInput(input: number | WorkbookBenchmarkCorpusId): {
  readonly workbookName: string
  readonly materializedCells: number
  readonly corpusCaseId: WorkbookBenchmarkCorpusId | null
  readonly corpusFamily: WorkbookBenchmarkCorpusFamily | null
  readonly importSnapshot: (engine: SpreadsheetEngine) => void
} {
  if (typeof input === 'string') {
    const corpus = buildWorkbookBenchmarkCorpus(input)
    return {
      workbookName: corpus.snapshot.workbook.name,
      materializedCells: corpus.materializedCellCount,
      corpusCaseId: corpus.id,
      corpusFamily: corpus.family,
      importSnapshot: (engine: SpreadsheetEngine) => {
        engine.importSnapshot(corpus.snapshot)
      },
    }
  }

  return {
    workbookName: 'benchmark-load',
    materializedCells: input,
    corpusCaseId: null,
    corpusFamily: null,
    importSnapshot: (engine: SpreadsheetEngine) => {
      seedLoadWorkbook(engine, input)
    },
  }
}

export async function runLoadBenchmark(input: number | WorkbookBenchmarkCorpusId = 10_000): Promise<LoadBenchmarkResult> {
  const resolved = resolveLoadBenchmarkInput(input)
  const engine = new SpreadsheetEngine({ workbookName: resolved.workbookName })
  await engine.ready()

  const memoryBefore = sampleMemory()
  const started = performance.now()
  resolved.importSnapshot(engine)
  const elapsed = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    scenario: 'load',
    materializedCells: resolved.materializedCells,
    corpusCaseId: resolved.corpusCaseId,
    corpusFamily: resolved.corpusFamily,
    elapsedMs: elapsed,
    memory: measureMemory(memoryBefore, memoryAfter),
  }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const rawInput = process.argv[2] ?? '10000'
  const input: number | WorkbookBenchmarkCorpusId = /^\d+$/.test(rawInput)
    ? Number.parseInt(rawInput, 10)
    : isWorkbookBenchmarkCorpusId(rawInput)
      ? rawInput
      : (() => {
          throw new Error(`Unknown workbook benchmark corpus: ${rawInput}`)
        })()
  console.log(JSON.stringify(await runLoadBenchmark(input), null, 2))
}
