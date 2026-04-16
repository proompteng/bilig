import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { EnginePivotError } from '../engine/errors.js'
import type { EnginePivotService } from '../engine/services/pivot-service.js'

function isEnginePivotService(value: unknown): value is EnginePivotService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'materializePivot') === 'function' &&
    typeof Reflect.get(value, 'resolvePivotData') === 'function' &&
    typeof Reflect.get(value, 'clearOwnedPivot') === 'function' &&
    typeof Reflect.get(value, 'clearPivotForCell') === 'function'
  )
}

function getPivotService(engine: SpreadsheetEngine): EnginePivotService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null || !('pivot' in runtime)) {
    throw new TypeError('Expected engine runtime to expose a pivot service')
  }
  const pivot = Reflect.get(runtime, 'pivot')
  if (!isEnginePivotService(pivot)) {
    throw new TypeError('Expected engine runtime pivot service')
  }
  return pivot
}

async function buildPivotEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: 'spec' })
  await engine.ready()
  engine.createSheet('Data')
  engine.createSheet('Pivot')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' }, [
    ['Region', 'Notes', 'Product', 'Sales'],
    ['East', 'priority', 'Widget', 10],
    ['West', 'priority', 'Widget', 7],
    ['East', 'priority', 'Gizmo', 5],
  ])
  engine.setPivotTable('Pivot', 'B2', {
    name: 'SalesByRegion',
    source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
    groupBy: ['Region'],
    values: [
      { sourceColumn: 'Sales', summarizeBy: 'sum' },
      { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
    ],
  })
  return engine
}

describe('EnginePivotService', () => {
  it('clears owned pivot output cells without leaving stale ownership behind', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)
    const pivot = engine.getPivotTables()[0]
    if (!pivot) {
      throw new TypeError('Expected pivot table')
    }

    const changed = Effect.runSync(service.clearOwnedPivot(pivot))

    expect(changed.length).toBeGreaterThan(0)
    expect(engine.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Empty })
  })

  it('resolves pivot aggregates through the extracted service boundary', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)

    const resolved = Effect.runSync(
      service.resolvePivotData('Pivot', 'C3', 'Sales', [{ field: 'Region', item: engine.getCellValue('Pivot', 'B3') }]),
    )

    expect(resolved).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('clears pivot ownership for one cell and rematerializes through the service boundary', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)
    const pivot = engine.getPivotTables()[0]
    if (!pivot) {
      throw new TypeError('Expected pivot table')
    }
    const pivotCellIndex = engine.workbook.getCellIndex('Pivot', 'B2')
    if (pivotCellIndex === undefined) {
      throw new TypeError('Expected pivot output cell')
    }

    const cleared = Effect.runSync(service.clearPivotForCell(pivotCellIndex))
    expect(cleared.length).toBeGreaterThan(0)
    expect(engine.getPivotTables()).toEqual([])
    expect(engine.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Empty })

    const rematerialized = Effect.runSync(service.materializePivot(pivot))
    expect(rematerialized.length).toBeGreaterThan(0)
    expect(engine.getCellValue('Pivot', 'B2')).toMatchObject({
      tag: ValueTag.String,
      value: 'Region',
    })
    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('drops orphaned pivot ownership when the pivot metadata is already gone', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)
    const pivotCellIndex = engine.workbook.getCellIndex('Pivot', 'B2')
    if (pivotCellIndex === undefined) {
      throw new TypeError('Expected pivot output cell')
    }

    expect(engine.deletePivotTable('Pivot', 'B2')).toBe(true)
    const pivotOutputOwners = Reflect.get(engine, 'pivotOutputOwners')
    if (!(pivotOutputOwners instanceof Map)) {
      throw new TypeError('Expected pivot output owners')
    }
    pivotOutputOwners.set(pivotCellIndex, 'Pivot!B2')

    expect(Effect.runSync(service.clearPivotForCell(pivotCellIndex))).toEqual([])
    expect(pivotOutputOwners.has(pivotCellIndex)).toBe(false)
  })

  it('wraps materialize, resolve, clear-owned, and clear-for-cell failures with EnginePivotError', async () => {
    const engine = await buildPivotEngine()
    const service = getPivotService(engine)
    const pivot = engine.getPivotTables()[0]
    if (!pivot) {
      throw new TypeError('Expected pivot table')
    }

    const getValueSpy = vi.spyOn(engine.workbook.cellStore, 'getValue').mockImplementation(() => {
      throw new Error('pivot cell explode')
    })
    const materialize = Effect.runSync(Effect.either(service.materializePivot(pivot)))
    expect(materialize._tag).toBe('Left')
    expect(materialize.left).toBeInstanceOf(EnginePivotError)
    expect(materialize.left.message).toContain(`Failed to materialize pivot ${pivot.name}`)
    const clearedOwned = Effect.runSync(Effect.either(service.clearOwnedPivot(pivot)))
    expect(clearedOwned._tag).toBe('Left')
    expect(clearedOwned.left).toBeInstanceOf(EnginePivotError)
    expect(clearedOwned.left.message).toContain(`Failed to clear pivot output ownership for ${pivot.name}`)
    getValueSpy.mockRestore()

    const listPivotsSpy = vi.spyOn(engine.workbook, 'listPivots').mockImplementation(() => {
      throw new Error('resolve explode')
    })
    const resolved = Effect.runSync(Effect.either(service.resolvePivotData('Pivot', 'B2', 'Sales', [])))
    expect(resolved._tag).toBe('Left')
    expect(resolved.left).toBeInstanceOf(EnginePivotError)
    expect(resolved.left.message).toContain('Failed to resolve pivot data for Pivot!B2')
    listPivotsSpy.mockRestore()

    const pivotCellIndex = engine.workbook.getCellIndex('Pivot', 'B2')
    if (pivotCellIndex === undefined) {
      throw new TypeError('Expected pivot output cell')
    }
    const getPivotByKeySpy = vi.spyOn(engine.workbook, 'getPivotByKey').mockImplementation(() => {
      throw new Error('clear cell explode')
    })
    const clearedForCell = Effect.runSync(Effect.either(service.clearPivotForCell(pivotCellIndex)))
    expect(clearedForCell._tag).toBe('Left')
    expect(clearedForCell.left).toBeInstanceOf(EnginePivotError)
    expect(clearedForCell.left.message).toContain(`Failed to clear pivot ownership for cell ${pivotCellIndex}`)
    getPivotByKeySpy.mockRestore()
  })
})
