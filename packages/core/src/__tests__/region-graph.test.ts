import { describe, expect, it } from 'vitest'
import { WorkbookStore } from '../workbook-store.js'
import { createRegionGraph } from '../deps/region-graph.js'
import { createEngineCounters } from '../perf/engine-counters.js'

describe('RegionGraph', () => {
  it('interns single-column regions canonically and collects point dependents through interval ownership', () => {
    const workbook = new WorkbookStore('region-graph')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const same = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const overlapping = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 2,
      rowEnd: 4,
      col: 0,
    })

    expect(same).toBe(first)

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [overlapping])

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 0, 0)]).toEqual([10])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 0, 0)).toBe(10)
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 2, 0)].toSorted((left, right) => left - right)).toEqual([10, 20])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 2, 0)).toBe(-2)
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 4, 0)]).toEqual([20])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 4, 0)).toBe(20)
    expect(regionGraph.getFormulaSubscriptions(10)).toEqual([first])

    regionGraph.clearFormulaSubscriptions(10)
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 2, 0)]).toEqual([20])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 2, 0)).toBe(20)
  })

  it('can prepare query indices eagerly before the first point lookup', () => {
    const workbook = new WorkbookStore('region-graph-prepare')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 16,
      rowEnd: 47,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [second])
    regionGraph.prepareQueryIndices()

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 0, 0)]).toEqual([10])
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 20, 0)].toSorted((left, right) => left - right)).toEqual([10, 20])

    regionGraph.clearFormulaSubscriptions(10)
    regionGraph.prepareQueryIndices()

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 20, 0)]).toEqual([20])
  })

  it('answers the first dirty point query without building an interval tree', () => {
    const counters = createEngineCounters()
    const workbook = new WorkbookStore('region-graph-dirty-point-query')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook, counters })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 16,
      rowEnd: 47,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [second])

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 20, 0)].toSorted((left, right) => left - right)).toEqual([10, 20])
    expect(counters.regionQueryIndexBuilds).toBe(0)

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 0, 0)]).toEqual([10])
    expect(counters.regionQueryIndexBuilds).toBe(0)

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 47, 0)]).toEqual([20])
    expect(counters.regionQueryIndexBuilds).toBe(1)

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 32, 0)]).toEqual([20])
    expect(counters.regionQueryIndexBuilds).toBe(1)
  })

  it('replaces a single formula region subscription without disturbing other subscribers', () => {
    const workbook = new WorkbookStore('region-graph-replace-single')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 3,
      rowEnd: 5,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first])
    regionGraph.replaceFormulaSubscriptions(20, [first])
    regionGraph.replaceSingleFormulaSubscription(10, first, second)

    expect(regionGraph.getFormulaSubscriptions(10)).toEqual([second])
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 1, 0)]).toEqual([20])
    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 4, 0)]).toEqual([10])
  })

  it('deduplicates dependents when one formula subscribes to multiple matching regions', () => {
    const workbook = new WorkbookStore('region-graph-dedupe')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook })

    const first = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 5,
      col: 0,
    })
    const second = regionGraph.internSingleColumnRegion({
      sheetName: 'Sheet1',
      rowStart: 2,
      rowEnd: 7,
      col: 0,
    })

    regionGraph.replaceFormulaSubscriptions(10, [first, second])

    expect([...regionGraph.collectFormulaDependentsForCell(sheet.id, 3, 0)]).toEqual([10])
    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 3, 0)).toBe(10)
  })

  it('keeps large sliding-window first point lookups on the dirty ordered scan path', () => {
    const counters = createEngineCounters()
    const workbook = new WorkbookStore('region-graph-sliding-large')
    const sheet = workbook.createSheet('Sheet1')
    const regionGraph = createRegionGraph({ workbook, counters })

    for (let row = 0; row < 1_500; row += 1) {
      const region = regionGraph.internSingleColumnRegion({
        sheetName: 'Sheet1',
        rowStart: row,
        rowEnd: row + 31,
        col: 0,
      })
      regionGraph.replaceFormulaSubscriptions(10_000 + row, [region])
    }

    expect(regionGraph.collectSingleFormulaDependentForCell(sheet.id, 0, 0)).toBe(10_000)
    expect(counters.regionQueryIndexBuilds).toBe(0)
  })
})
