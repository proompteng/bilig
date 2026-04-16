import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { CellFlags } from '../cell-store.js'
import { SpreadsheetEngine } from '../engine.js'
import type { EngineFormulaGraphService } from '../engine/services/formula-graph-service.js'

function isEngineFormulaGraphService(value: unknown): value is EngineFormulaGraphService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'rebuildTopoRanks') === 'function' &&
    typeof Reflect.get(value, 'detectCycles') === 'function' &&
    typeof Reflect.get(value, 'scheduleWasmProgramSync') === 'function'
  )
}

function getGraphService(engine: SpreadsheetEngine): EngineFormulaGraphService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const graph = Reflect.get(runtime, 'graph')
  if (!isEngineFormulaGraphService(graph)) {
    throw new TypeError('Expected engine formula graph service')
  }
  return graph
}

describe('EngineFormulaGraphService', () => {
  it('rebuilds topo ranks for range-node dependents through the service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'graph-topo' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B1)')

    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const d1Index = engine.workbook.getCellIndex('Sheet1', 'D1')

    expect(b1Index).toBeDefined()
    expect(d1Index).toBeDefined()

    engine.workbook.cellStore.topoRanks[b1Index!] = 99
    engine.workbook.cellStore.topoRanks[d1Index!] = 1

    Effect.runSync(getGraphService(engine).rebuildTopoRanks())

    expect(engine.workbook.cellStore.topoRanks[b1Index!]).toBeLessThan(engine.workbook.cellStore.topoRanks[d1Index!])
  })

  it('restores cycle flags and error values when cycle detection reruns through the service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'graph-cycle' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'B1+1')
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')

    expect(a1Index).toBeDefined()
    expect(b1Index).toBeDefined()

    engine.workbook.cellStore.flags[a1Index!] &= ~CellFlags.InCycle
    engine.workbook.cellStore.flags[b1Index!] &= ~CellFlags.InCycle
    engine.workbook.cellStore.cycleGroupIds[a1Index!] = -1
    engine.workbook.cellStore.cycleGroupIds[b1Index!] = -1
    engine.workbook.cellStore.setValue(a1Index!, { tag: ValueTag.Number, value: 1 })
    engine.workbook.cellStore.setValue(b1Index!, { tag: ValueTag.Number, value: 1 })

    Effect.runSync(getGraphService(engine).detectCycles())

    expect((engine.workbook.cellStore.flags[a1Index!] & CellFlags.InCycle) !== 0).toBe(true)
    expect((engine.workbook.cellStore.flags[b1Index!] & CellFlags.InCycle) !== 0).toBe(true)
    expect(engine.workbook.cellStore.cycleGroupIds[a1Index!]).toBeGreaterThanOrEqual(0)
    expect(engine.workbook.cellStore.cycleGroupIds[a1Index!]).toBe(engine.workbook.cellStore.cycleGroupIds[b1Index!])
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })
})
