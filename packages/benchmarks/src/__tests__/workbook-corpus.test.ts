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

  it('lists stable corpus definitions and resolves them by id', () => {
    const definitions = listWorkbookBenchmarkCorpusDefinitions()

    expect(definitions.map((definition) => definition.id)).toEqual([
      'dense-mixed-100k',
      'dense-mixed-250k',
      'analysis-multisheet-100k',
      'analysis-multisheet-250k',
    ])
    expect(getWorkbookBenchmarkCorpusDefinition('dense-mixed-250k')).toEqual(definitions[1])
  })
})
