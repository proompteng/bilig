import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { ValueTag } from '@bilig/protocol'
import { decodeCellAddress, readXlsxFormulaCacheCellsFromFile, writeSimpleXlsxWorkbook, type SimpleXlsxCell } from '@bilig/xlsx'
import { describe, expect, it } from 'vitest'

import { inspectXlsxCache, recalculateXlsx, recalculateXlsxFileToFile, recalculateXlsxToFile } from '../xlsx.js'

describe('@bilig/workpaper/xlsx native file-to-file recalc', () => {
  it('keeps the public large-workbook path on streaming-native while preserving legacy bytes API exports', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-native-file-recalc-'))
    try {
      const sourcePath = join(tempDir, 'pricing.xlsx')
      const outputPath = join(tempDir, 'pricing.recalculated.xlsx')
      const fallbackOutputPath = join(tempDir, 'pricing.workpaper.xlsx')
      writeFileSync(sourcePath, buildPricingWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        edits: [
          { target: 'Inputs!B2', value: 48 },
          { target: 'Inputs!B3', value: 1500 },
        ],
        reads: ['Summary!B2'],
        maxRssBytes: 350 * 1024 * 1024,
      })

      expect(result.reads['Summary!B2']).toMatchObject({ tag: ValueTag.Number, value: 72_000 })
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      expect(readFormulaCache(outputPath, 'Summary!B2')).toEqual({
        target: 'Summary!B2',
        formula: '=Inputs!B2*Inputs!B3',
        cachedValue: 72_000,
      })
      expect(typeof recalculateXlsx).toBe('function')

      await expect(
        recalculateXlsxFileToFile(sourcePath, {
          outputPath: fallbackOutputPath,
          // @ts-expect-error Runtime guard for legacy JavaScript callers.
          engine: 'workpaper',
          reads: ['Summary!B2'],
        }),
      ).rejects.toThrow(/legacy bytes APIs/u)
      expect(existsSync(fallbackOutputPath)).toBe(false)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('rejects oversized legacy bytes APIs before WorkPaper materialization', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-large-bytes-recalc-'))
    const oversizedInput = new Uint8Array(1_000_001)

    try {
      expect(() => recalculateXlsx(oversizedInput, { fileName: 'large.xlsx' })).toThrow(/legacy bytes API is small-workbook only/u)
      expect(() =>
        recalculateXlsxToFile(oversizedInput, {
          outputPath: join(tempDir, 'large.recalculated.xlsx'),
          fileName: 'large.xlsx',
        }),
      ).toThrow(/Use recalculateXlsxFileToFile\(\)/u)
      expect(() => inspectXlsxCache(oversizedInput, { fileName: 'large.xlsx' })).toThrow(/legacy bytes API is small-workbook only/u)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})

function buildPricingWorkbook(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Inputs',
        cells: [cell('A1', 'Metric'), cell('B1', 'Value'), cell('A2', 'Units'), cell('B2', 40), cell('A3', 'Price'), cell('B3', 1200)],
      },
      {
        name: 'Summary',
        cells: [cell('A1', 'Metric'), cell('B1', 'Value'), cell('A2', 'Revenue'), cell('B2', 1, 'Inputs!B2*Inputs!B3')],
      },
    ],
  })
}

function cell(address: string, value: string | number | boolean | null, formula?: string): SimpleXlsxCell {
  const decoded = decodeCellAddress(address)
  return {
    address,
    row: decoded.r,
    col: decoded.c,
    value,
    ...(formula === undefined ? {} : { formula }),
  }
}

function readFormulaCache(inputPath: string, target: string) {
  return readXlsxFormulaCacheCellsFromFile(inputPath).cells.find((formulaCell) => formulaCell.target === target)
}
