import { describe, expect, it } from 'vitest'
import { WorkPaperSheetDimensionCache, type WorkPaperSheetDimensionEngine } from '../work-paper-sheet-dimension-cache.js'
import type { WorkPaperSheetDimensions } from '../work-paper-types.js'

function engineWithSpills(spills: readonly string[] = []): WorkPaperSheetDimensionEngine {
  return {
    workbook: {
      listSpills: () => spills.map((sheetName) => ({ sheetName })),
      getSheet: (sheetName) => {
        return sheetName === 'Sheet1' ? { id: 1 } : sheetName === 'Sheet2' ? { id: 2 } : undefined
      },
    },
  }
}

describe('WorkPaperSheetDimensionCache', () => {
  it('caches defensive dimension copies and scans sparse sheet records', () => {
    const cache = new WorkPaperSheetDimensionCache(engineWithSpills())
    const dimensions: WorkPaperSheetDimensions = { width: 2, height: 3 }

    cache.cache(1, dimensions)
    dimensions.width = 8

    expect(cache.get(1)).toEqual({ width: 2, height: 3 })
    expect(
      cache.scan({
        grid: {
          forEachCellEntry(callback) {
            callback(10, 4, 3)
            callback(11, 1, 8)
          },
        },
      }),
    ).toEqual({ width: 9, height: 5 })
  })

  it('does not cache initialized dimensions for sheets with spills', () => {
    const cache = new WorkPaperSheetDimensionCache(engineWithSpills(['Sheet1']))

    cache.cacheInitialized(1, { width: 2, height: 3 })
    cache.cacheInitialized(2, { width: 4, height: 5 })

    expect(cache.get(1)).toBeUndefined()
    expect(cache.get(2)).toEqual({ width: 4, height: 5 })
  })

  it('does not cache scanned dimensions for sheets with spills', () => {
    const cache = new WorkPaperSheetDimensionCache(engineWithSpills(['Sheet1']))

    cache.cacheScanned(1, { width: 2, height: 3 })
    cache.cacheScanned(2, { width: 4, height: 5 })

    expect(cache.get(1)).toBeUndefined()
    expect(cache.get(2)).toEqual({ width: 4, height: 5 })
  })

  it('expands dimensions for non-spill value writes and invalidates edge clears', () => {
    const cache = new WorkPaperSheetDimensionCache(engineWithSpills())
    cache.cache(1, { width: 2, height: 2 })

    cache.updateAfterCellMutationRefs([{ sheetId: 1, mutation: { kind: 'setCellValue', row: 4, col: 3, value: 1 } }])
    expect(cache.get(1)).toEqual({ width: 4, height: 5 })

    cache.updateAfterCellMutationRefs([{ sheetId: 1, mutation: { kind: 'clearCell', row: 4, col: 3 } }])
    expect(cache.get(1)).toBeUndefined()
  })

  it('keeps interior clears and existing-cell writes on cached no-spill dimensions', () => {
    const cache = new WorkPaperSheetDimensionCache(engineWithSpills())
    cache.cache(1, { width: 3, height: 3 })
    cache.updateAfterCellMutationRefs([{ sheetId: 1, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 1 } }])

    expect(cache.get(1)).toEqual({ width: 3, height: 3 })

    cache.updateAfterCellMutationRefs([{ sheetId: 1, mutation: { kind: 'clearCell', row: 1, col: 1 } }])

    expect(cache.get(1)).toEqual({ width: 3, height: 3 })
  })

  it('expands scalar formulas and invalidates dynamic formulas and known spill sheets', () => {
    const cache = new WorkPaperSheetDimensionCache(engineWithSpills(['Sheet2']))
    cache.cache(1, { width: 3, height: 3 })
    cache.cache(2, { width: 3, height: 3 })

    cache.updateAfterCellMutationRefs([{ sheetId: 1, mutation: { kind: 'setCellFormula', row: 0, col: 0, formula: '=A1' } }])
    expect(cache.get(1)).toEqual({ width: 3, height: 3 })

    cache.updateAfterCellMutationRefs([{ sheetId: 1, mutation: { kind: 'setCellFormula', row: 4, col: 3, formula: '=SUM(A5:C5)' } }])
    expect(cache.get(1)).toEqual({ width: 4, height: 5 })

    cache.updateAfterCellMutationRefs([
      { sheetId: 1, mutation: { kind: 'setCellFormula', row: 0, col: 0, formula: '=FILTER(A1:A3,A1:A3>1)' } },
    ])
    cache.updateAfterCellMutationRefs([{ sheetId: 2, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 1 } }])

    expect(cache.get(1)).toBeUndefined()
    expect(cache.get(2)).toBeUndefined()
  })

  it('updates dimensions directly from matrix mutation impact', () => {
    const cache = new WorkPaperSheetDimensionCache(engineWithSpills())
    cache.cache(1, { width: 2, height: 2 })

    cache.updateAfterMatrixMutationImpact({
      hasDynamicFormula: false,
      maxClearCol: -1,
      maxClearRow: -1,
      maxSetCol: 4,
      maxSetRow: 5,
      sheetId: 1,
    })

    expect(cache.get(1)).toEqual({ width: 5, height: 6 })

    cache.updateAfterMatrixMutationImpact({
      hasDynamicFormula: false,
      maxClearCol: 1,
      maxClearRow: 1,
      maxSetCol: -1,
      maxSetRow: -1,
      sheetId: 1,
    })

    expect(cache.get(1)).toEqual({ width: 5, height: 6 })

    cache.updateAfterMatrixMutationImpact({
      hasDynamicFormula: false,
      maxClearCol: 4,
      maxClearRow: 0,
      maxSetCol: -1,
      maxSetRow: -1,
      sheetId: 1,
    })

    expect(cache.get(1)).toBeUndefined()
  })

  it('invalidates matrix mutations with dynamic formulas or known spill sheets', () => {
    const cache = new WorkPaperSheetDimensionCache(engineWithSpills(['Sheet2']))
    cache.cache(1, { width: 2, height: 2 })
    cache.cache(2, { width: 2, height: 2 })

    cache.updateAfterMatrixMutationImpact({
      hasDynamicFormula: true,
      maxClearCol: -1,
      maxClearRow: -1,
      maxSetCol: 0,
      maxSetRow: 0,
      sheetId: 1,
    })
    cache.updateAfterMatrixMutationImpact({
      hasDynamicFormula: false,
      maxClearCol: -1,
      maxClearRow: -1,
      maxSetCol: 0,
      maxSetRow: 0,
      sheetId: 2,
    })

    expect(cache.get(1)).toBeUndefined()
    expect(cache.get(2)).toBeUndefined()
  })

  it('clears dimensions and refreshes spill knowledge on invalidateAll', () => {
    const spills: string[] = ['Sheet1']
    const cache = new WorkPaperSheetDimensionCache(engineWithSpills(spills))

    cache.cacheInitialized(1, { width: 1, height: 1 })
    expect(cache.get(1)).toBeUndefined()

    spills.length = 0
    cache.invalidateAll()
    cache.cacheInitialized(1, { width: 1, height: 1 })

    expect(cache.get(1)).toEqual({ width: 1, height: 1 })
  })
})
