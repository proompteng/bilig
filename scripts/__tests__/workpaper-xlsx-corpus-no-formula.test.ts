import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { describe, expect, it, vi } from 'vitest'

vi.mock('@bilig/excel-import', () => ({
  importXlsx: () => {
    throw new Error(
      'SheetJS xlsx is not a production dependency of @bilig/excel-import. Install xlsx only for legacy SheetJS fallback import/export paths, or use the @bilig/xlsx source-preserving XLSX path.',
    )
  },
}))

const { runWorkPaperXlsxCorpus } = await import('../check-workpaper-xlsx-corpus.ts')

describe('WorkPaper XLSX corpus verifier formula-free fast path', () => {
  it('does not attach an imported runtime snapshot when a workbook has no formulas', () => {
    withTempCorpus((corpusDir) => {
      writeFileSync(
        join(corpusDir, 'no-formulas.xlsx'),
        writeSimpleXlsxWorkbook({
          sheets: [
            {
              name: 'Trial Balance',
              cells: [
                { address: 'A1', row: 0, col: 0, value: 'Account' },
                { address: 'B1', row: 0, col: 1, value: 'Amount' },
                { address: 'A2', row: 1, col: 0, value: 'Cash' },
                { address: 'B2', row: 1, col: 1, value: 1200 },
                { address: 'A3', row: 2, col: 0, value: 'Revenue' },
                { address: 'B3', row: 2, col: 1, value: 3400 },
              ],
            },
          ],
        }),
      )

      const result = runWorkPaperXlsxCorpus([corpusDir])

      expect(result.summary).toMatchObject({
        totalFiles: 1,
        filesProcessed: 1,
        ok: 1,
        failedErrors: 0,
        formulaCells: 0,
      })
      expect(result.files[0]).toMatchObject({
        fileName: 'no-formulas.xlsx',
        status: 'ok',
        formulaCells: 0,
      })
    })
  })

  it('keeps checked-in native formula fixtures passing when optional SheetJS is absent', () => {
    const result = runWorkPaperXlsxCorpus([checkedInThreadedCommentsFixture()])

    expect(result.summary).toMatchObject({
      totalFiles: 1,
      filesProcessed: 1,
      ok: 1,
      failedErrors: 0,
    })
    expect(result.summary.formulaCells).toBeGreaterThan(0)
    expect(result.files[0]).toMatchObject({
      fileName: 'macos-excel-threaded-comments-source.xlsx',
      status: 'ok',
    })
    expect(result.files[0]?.formulaCells).toBeGreaterThan(0)
  })
})

function checkedInThreadedCommentsFixture(): string {
  return join(
    dirname(fileURLToPath(import.meta.url)),
    '../../packages/headless/fixtures/xlsx-corpus/macos-excel-threaded-comments-source.xlsx',
  )
}

function withTempCorpus(run: (corpusDir: string) => void): void {
  const corpusDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-xlsx-corpus-no-formula-'))
  try {
    run(corpusDir)
  } finally {
    rmSync(corpusDir, { recursive: true, force: true })
  }
}
