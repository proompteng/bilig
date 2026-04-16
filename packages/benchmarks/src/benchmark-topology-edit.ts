import { performance } from 'node:perf_hooks'
import { SpreadsheetEngine } from '@bilig/core'
import type { RecalcMetrics } from '@bilig/protocol'
import { seedTopologyEditWorkbook } from './generate-workbook.js'
import { measureMemory, sampleMemory, type MemoryMeasurement } from './metrics.js'

export interface TopologyEditBenchmarkResult {
  scenario: 'topology-edit'
  chainLength: number
  elapsedMs: number
  metrics: RecalcMetrics
  memory: MemoryMeasurement
}

export async function runTopologyEditBenchmark(chainLength = 10_000): Promise<TopologyEditBenchmarkResult> {
  const engine = new SpreadsheetEngine({ workbookName: 'benchmark-topology-edit' })
  await engine.ready()
  seedTopologyEditWorkbook(engine, chainLength)

  const memoryBefore = sampleMemory()
  const started = performance.now()
  engine.setCellFormula('Sheet1', 'B1', 'A1+A2')
  const elapsed = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    scenario: 'topology-edit',
    chainLength,
    elapsedMs: elapsed,
    metrics: engine.getLastMetrics(),
    memory: measureMemory(memoryBefore, memoryAfter),
  }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const chainLength = Number.parseInt(process.argv[2] ?? '10000', 10)
  console.log(JSON.stringify(await runTopologyEditBenchmark(chainLength), null, 2))
}
