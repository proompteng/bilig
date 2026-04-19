import { describe, expect, it } from 'vitest'
import {
  buildWorkbookBenchmarkCorpus,
  countWorkbookSnapshotCells,
  getWorkbookBenchmarkCorpusDefinition,
  listWorkbookBenchmarkCorpusDefinitions,
} from '../workbook-corpus.js'

describe('workbook benchmark corpus', () => {
  it('builds the dense 100k corpus as an exact-size mixed-value workbook', () => {
    const corpus = buildWorkbookBenchmarkCorpus('dense-mixed-100k')

    expect(corpus.family).toBe('dense-mixed')
    expect(corpus.sheetCount).toBe(1)
    expect(countWorkbookSnapshotCells(corpus.snapshot)).toBe(100_000)
    expect(corpus.snapshot.sheets[0]?.name).toBe('Grid')
    expect(corpus.snapshot.sheets[0]?.cells.slice(0, 4)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 3 },
      { address: 'C1', formula: 'A1*B1' },
      { address: 'D1', value: 'segment-1' },
    ])
  })

  it('builds the multisheet 250k corpus with cross-sheet formulas and a ledger viewport', () => {
    const corpus = buildWorkbookBenchmarkCorpus('analysis-multisheet-250k')

    expect(corpus.family).toBe('analysis-multisheet')
    expect(corpus.primaryViewport).toEqual({
      sheetName: 'Ledger',
      rowStart: 0,
      rowEnd: 39,
      colStart: 0,
      colEnd: 4,
    })
    expect(countWorkbookSnapshotCells(corpus.snapshot)).toBe(250_000)
    expect(corpus.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Inputs', 'Ledger', 'Summary'])
    expect(corpus.snapshot.sheets[0]?.cells.slice(0, 3)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 2 },
      { address: 'C1', formula: 'A1*B1' },
    ])
    expect(corpus.snapshot.sheets[1]?.cells.slice(0, 5)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 5 },
      { address: 'C1', value: 1 },
      { address: 'D1', formula: 'A1*B1' },
      { address: 'E1', formula: 'D1+C1+Inputs!C1' },
    ])
    expect(corpus.snapshot.sheets[2]?.cells[1]).toEqual({
      address: 'B1',
      formula: 'SUM(Ledger!E1:E256)',
    })
  })

  it('builds the wide 250k corpus as a horizontally dense workbook', () => {
    const corpus = buildWorkbookBenchmarkCorpus('wide-mixed-250k')

    expect(corpus.family).toBe('wide-mixed')
    expect(corpus.primaryViewport).toEqual({
      sheetName: 'WideGrid',
      rowStart: 0,
      rowEnd: 39,
      colStart: 0,
      colEnd: 9,
    })
    expect(countWorkbookSnapshotCells(corpus.snapshot)).toBe(250_000)
    expect(corpus.snapshot.sheets[0]?.name).toBe('WideGrid')
    expect(corpus.snapshot.sheets[0]?.cells.slice(0, 12)).toEqual([
      { address: 'A1', value: 'metric-1' },
      { address: 'B1', value: 'metric-2' },
      { address: 'C1', value: 'metric-3' },
      { address: 'D1', value: 'metric-4' },
      { address: 'E1', value: 'metric-5' },
      { address: 'F1', value: 'metric-6' },
      { address: 'G1', value: 'metric-7' },
      { address: 'H1', value: 'metric-8' },
      { address: 'I1', value: 'metric-9' },
      { address: 'J1', value: 'metric-10' },
      { address: 'K1', value: 'metric-11' },
      { address: 'L1', value: 'metric-12' },
    ])
  })

  it('lists stable corpus definitions and resolves them by id', () => {
    const definitions = listWorkbookBenchmarkCorpusDefinitions()

    expect(definitions.map((definition) => definition.id)).toEqual([
      'dense-mixed-100k',
      'dense-mixed-250k',
      'wide-mixed-250k',
      'wide-mixed-frozen-250k',
      'wide-mixed-variable-250k',
      'analysis-multisheet-100k',
      'analysis-multisheet-250k',
    ])
    expect(getWorkbookBenchmarkCorpusDefinition('dense-mixed-250k')).toEqual(definitions[1])
  })

  it('describes deterministic presentation metadata for frozen and variable-width browse corpora', () => {
    expect(getWorkbookBenchmarkCorpusDefinition('wide-mixed-frozen-250k').presentation).toEqual({
      freezeRows: 2,
      freezeCols: 2,
    })
    expect(getWorkbookBenchmarkCorpusDefinition('wide-mixed-variable-250k').presentation?.columnWidths?.slice(0, 4)).toEqual([
      { index: 0, size: 120 },
      { index: 1, size: 224 },
      { index: 2, size: 96 },
      { index: 3, size: 168 },
    ])
  })
})
