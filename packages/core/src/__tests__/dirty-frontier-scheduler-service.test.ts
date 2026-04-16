import { describe, expect, it } from 'vitest'
import { FormulaTable } from '../formula-table.js'
import { WorkbookStore } from '../workbook-store.js'
import { RangeRegistry } from '../range-registry.js'
import { RecalcScheduler } from '../scheduler.js'
import { makeCellEntity, makeExactLookupColumnEntity, makeSortedLookupColumnEntity } from '../entity-ids.js'
import { createEngineDirtyFrontierSchedulerService } from '../engine/services/dirty-frontier-scheduler-service.js'

describe('createEngineDirtyFrontierSchedulerService', () => {
  it('collects dirty formulas in topological order from changed roots', () => {
    const workbook = new WorkbookStore('dirty-frontier')
    workbook.createSheet('Sheet1')
    const a1 = workbook.ensureCell('Sheet1', 'A1')
    const b1 = workbook.ensureCell('Sheet1', 'B1')
    const c1 = workbook.ensureCell('Sheet1', 'C1')

    workbook.cellStore.topoRanks[b1] = 0
    workbook.cellStore.topoRanks[c1] = 1

    const formulas = new FormulaTable<{ cellIndex: number }>(workbook.cellStore)
    formulas.set(b1, { cellIndex: b1 })
    formulas.set(c1, { cellIndex: c1 })

    const scheduler = createEngineDirtyFrontierSchedulerService({
      state: {
        workbook,
        formulas,
        ranges: new RangeRegistry(),
        scheduler: new RecalcScheduler(),
      },
      getEntityDependents: (entityId) => {
        if (entityId === makeCellEntity(a1)) {
          return Uint32Array.of(b1)
        }
        if (entityId === makeCellEntity(b1)) {
          return Uint32Array.of(c1)
        }
        return new Uint32Array()
      },
    })

    const result = scheduler.collectDirty(Uint32Array.of(a1))

    expect(Array.from(result.orderedFormulaCellIndices.subarray(0, result.orderedFormulaCount))).toEqual([b1, c1])
  })

  it('seeds dirty formulas from reverse lookup-column subscribers', () => {
    const workbook = new WorkbookStore('dirty-frontier-lookup')
    workbook.createSheet('Sheet1')
    const a1 = workbook.ensureCell('Sheet1', 'A1')
    const b1 = workbook.ensureCell('Sheet1', 'B1')
    workbook.cellStore.topoRanks[b1] = 0

    const formulas = new FormulaTable<{ cellIndex: number }>(workbook.cellStore)
    formulas.set(b1, { cellIndex: b1 })
    const lookupEntity = makeExactLookupColumnEntity(workbook.getSheet('Sheet1')!.id, 0)

    const scheduler = createEngineDirtyFrontierSchedulerService({
      state: {
        workbook,
        formulas,
        ranges: new RangeRegistry(),
        scheduler: new RecalcScheduler(),
      },
      getEntityDependents: (entityId) => {
        if (entityId === makeCellEntity(a1)) {
          return Uint32Array.of(lookupEntity)
        }
        if (entityId === lookupEntity) {
          return Uint32Array.of(b1)
        }
        return new Uint32Array()
      },
    })

    const result = scheduler.collectDirty(Uint32Array.of(a1))

    expect(Array.from(result.orderedFormulaCellIndices.subarray(0, result.orderedFormulaCount))).toEqual([b1])
  })

  it('schedules exact and sorted lookup subscribers on the same column without collisions', () => {
    const workbook = new WorkbookStore('dirty-frontier-lookup-collision')
    workbook.createSheet('Sheet1')
    const a1 = workbook.ensureCell('Sheet1', 'A1')
    const b1 = workbook.ensureCell('Sheet1', 'B1')
    const c1 = workbook.ensureCell('Sheet1', 'C1')
    workbook.cellStore.topoRanks[b1] = 0
    workbook.cellStore.topoRanks[c1] = 1

    const formulas = new FormulaTable<{ cellIndex: number }>(workbook.cellStore)
    formulas.set(b1, { cellIndex: b1 })
    formulas.set(c1, { cellIndex: c1 })

    const sheetId = workbook.getSheet('Sheet1')!.id
    const exactLookupEntity = makeExactLookupColumnEntity(sheetId, 0)
    const sortedLookupEntity = makeSortedLookupColumnEntity(sheetId, 0)

    const scheduler = createEngineDirtyFrontierSchedulerService({
      state: {
        workbook,
        formulas,
        ranges: new RangeRegistry(),
        scheduler: new RecalcScheduler(),
      },
      getEntityDependents: (entityId) => {
        if (entityId === makeCellEntity(a1)) {
          return Uint32Array.of(exactLookupEntity, sortedLookupEntity)
        }
        if (entityId === exactLookupEntity) {
          return Uint32Array.of(b1)
        }
        if (entityId === sortedLookupEntity) {
          return Uint32Array.of(c1)
        }
        return new Uint32Array()
      },
    })

    const result = scheduler.collectDirty(Uint32Array.of(a1))

    expect(Array.from(result.orderedFormulaCellIndices.subarray(0, result.orderedFormulaCount))).toEqual([b1, c1])
  })
})
