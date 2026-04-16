import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { runProperty } from '@bilig/test-fuzz'
import {
  applyMetadataStructuralAction,
  metadataStructuralActionArbitrary,
  restoreMetadataSnapshot,
  type MetadataStructuralAction,
} from './engine-advanced-metadata-fuzz-helpers.js'

type PivotFuzzAction = MetadataStructuralAction | { kind: 'value'; address: string; value: number }

describe('engine pivot fuzz', () => {
  it('preserves pivot metadata and rendered pivot values across structural edits, value updates, and snapshot restore', async () => {
    await runProperty({
      suite: 'core/pivot/structural-roundtrip',
      arbitrary: fc.array(pivotFuzzActionArbitrary, {
        minLength: 1,
        maxLength: 10,
      }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({ workbookName: 'pivot-fuzz' })
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

        actions.forEach((action) => {
          if (action.kind === 'value') {
            engine.setCellValue('Data', action.address, action.value)
            return
          }
          applyMetadataStructuralAction(engine, action)
        })

        const pivotTables = engine.getPivotTables()
        const restored = await restoreMetadataSnapshot(engine, 'pivot-fuzz-restored')
        expect(restored.getPivotTables()).toEqual(pivotTables)
        pivotTables.forEach((pivot) => {
          for (let rowOffset = 0; rowOffset < pivot.rows; rowOffset += 1) {
            for (let colOffset = 0; colOffset < pivot.cols; colOffset += 1) {
              const anchor = parseCellAddress(pivot.address, pivot.sheetName)
              const address = formatAddress(anchor.row + rowOffset, anchor.col + colOffset)
              expect(normalizeCellValue(restored.getCellValue(pivot.sheetName, address))).toEqual(
                normalizeCellValue(engine.getCellValue(pivot.sheetName, address)),
              )
            }
          }
        })
      },
    })
  })
})

const pivotFuzzActionArbitrary = fc.oneof<PivotFuzzAction>(
  metadataStructuralActionArbitrary(['Data', 'Pivot']),
  fc
    .record({
      address: fc.constantFrom('D2', 'D3', 'D4'),
      value: fc.integer({ min: 0, max: 20 }),
    })
    .map((action) => Object.assign({ kind: 'value' as const }, action)),
)

function normalizeCellValue(value: CellValue): unknown {
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
      return { tag: value.tag, value: value.value }
    case ValueTag.Boolean:
      return { tag: value.tag, value: value.value }
    case ValueTag.String:
      return { tag: value.tag, value: value.value }
    case ValueTag.Error:
      return { tag: value.tag, code: value.code }
  }
}
