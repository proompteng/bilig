import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ValueTag, formatCellDisplayValue } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import type { EngineReadService } from '../engine/services/read-service.js'

function isEngineReadService(value: unknown): value is EngineReadService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'getCell') === 'function' &&
    typeof Reflect.get(value, 'getDependencies') === 'function' &&
    typeof Reflect.get(value, 'explainCell') === 'function'
  )
}

function getReadService(engine: SpreadsheetEngine): EngineReadService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const read = Reflect.get(runtime, 'read')
  if (!isEngineReadService(read)) {
    throw new TypeError('Expected engine read service')
  }
  return read
}

describe('EngineReadService', () => {
  it('returns formatted cell snapshots through the extracted read boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'read-cell' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 12)
    engine.setCellFormat('Sheet1', 'A1', '0.00')

    const cell = Effect.runSync(getReadService(engine).getCell('Sheet1', 'A1'))

    expect(cell).toEqual(
      expect.objectContaining({
        sheetName: 'Sheet1',
        address: 'A1',
        value: { tag: ValueTag.Number, value: 12 },
        format: '0.00',
      }),
    )
  })

  it('explains direct precedents and dependents through the extracted read boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'read-deps' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'C1', 'B1+1')

    const explanation = Effect.runSync(getReadService(engine).explainCell('Sheet1', 'B1'))

    expect(explanation.directPrecedents).toEqual(['Sheet1!A1'])
    expect(explanation.directDependents).toEqual(['Sheet1!C1'])
    expect(explanation.value).toEqual({ tag: ValueTag.Number, value: 20 })
  })

  it('infers short date formatting for serials under date-like headers without explicit formats', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'read-date-headers' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Start Date')
    engine.setCellValue('Sheet1', 'A2', 46023)
    engine.setCellValue('Sheet1', 'B1', 'End Date')
    engine.setCellValue('Sheet1', 'B2', 46357)

    const startCell = Effect.runSync(getReadService(engine).getCell('Sheet1', 'A2'))
    const endCell = Effect.runSync(getReadService(engine).getCell('Sheet1', 'B2'))

    expect(startCell.format).toBe('date:short')
    expect(formatCellDisplayValue(startCell.value, startCell.format)).toBe('01/01/2026')
    expect(endCell.format).toBe('date:short')
    expect(formatCellDisplayValue(endCell.value, endCell.format)).toBe('12/01/2026')
  })

  it('infers short date formatting for date formulas without explicit formats', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'read-date-formula' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Month')
    engine.setCellFormula('Sheet1', 'A2', 'DATE(2026,12,1)')

    const monthCell = Effect.runSync(getReadService(engine).getCell('Sheet1', 'A2'))

    expect(monthCell.formula).toBe('DATE(2026,12,1)')
    expect(monthCell.format).toBe('date:short')
    expect(formatCellDisplayValue(monthCell.value, monthCell.format)).toBe('12/01/2026')
  })
})
