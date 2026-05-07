import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import * as XLSX from 'xlsx'

vi.mock('@bilig/excel-import', () => ({
  importXlsx: () => {
    throw new Error('formula-free workbooks should not use importXlsx')
  },
}))

const { runWorkPaperXlsxCorpus } = await import('../check-workpaper-xlsx-corpus.ts')

describe('WorkPaper XLSX corpus verifier formula-free fast path', () => {
  it('does not attach an imported runtime snapshot when a workbook has no formulas', () => {
    withTempCorpus((corpusDir) => {
      const workbook = XLSX.utils.book_new()
      const sheet = XLSX.utils.aoa_to_sheet([
        ['Account', 'Amount'],
        ['Cash', 1200],
        ['Revenue', 3400],
      ])
      XLSX.utils.book_append_sheet(workbook, sheet, 'Trial Balance')
      writeFileSync(join(corpusDir, 'no-formulas.xlsx'), XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }))

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
})

function withTempCorpus(run: (corpusDir: string) => void): void {
  const corpusDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-xlsx-corpus-no-formula-'))
  try {
    run(corpusDir)
  } finally {
    rmSync(corpusDir, { recursive: true, force: true })
  }
}
