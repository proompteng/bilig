import { describe, expect, it } from 'vitest'

import { buildWorkbookBenchmarkCorpus } from '../../packages/benchmarks/src/workbook-corpus.js'
import { exportXlsx } from '../../packages/excel-import/src/index.js'
import {
  buildSameCorpusPublicAccessCheck,
  googleSheetsXlsxExportUrl,
  microsoftExcelWebSourceXlsxUrl,
  validateSameCorpusPublicAccessCheck,
  type SameCorpusPublicAccessFetch,
} from '../ui-responsiveness-same-corpus-public-access-check.ts'

describe('same-corpus public access check', () => {
  it('normalizes Google Sheets edit URLs to XLSX export URLs', () => {
    expect(googleSheetsXlsxExportUrl('https://docs.google.com/spreadsheets/d/sheet-id/edit#gid=0')).toBe(
      'https://docs.google.com/spreadsheets/d/sheet-id/export?format=xlsx',
    )
  })

  it('extracts the source XLSX URL from Microsoft Excel Web viewer URLs', () => {
    const source = 'https://example.com/workbook.xlsx'
    expect(microsoftExcelWebSourceXlsxUrl(`https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(source)}`)).toBe(source)
  })

  it('verifies requested public URLs against deterministic same-corpus XLSX bytes', async () => {
    const corpus = buildWorkbookBenchmarkCorpus('wide-mixed-250k')
    const workbookBytes = new Uint8Array(Buffer.from(exportXlsx(corpus.snapshot)))
    const fetchedUrls: string[] = []
    const fetchXlsxBytes: SameCorpusPublicAccessFetch = async (url) => {
      fetchedUrls.push(url)
      return workbookBytes
    }

    const check = await buildSameCorpusPublicAccessCheck({
      corpusId: 'wide-mixed-250k',
      fetchXlsxBytes,
      generatedAt: '2026-05-08T00:00:00.000Z',
      googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      microsoftExcelWebUrl: 'https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fexample.com%2Fworkbook.xlsx',
    })

    expect(fetchedUrls).toEqual(['https://docs.google.com/spreadsheets/d/sheet-id/export?format=xlsx', 'https://example.com/workbook.xlsx'])
    expect(check).toMatchObject({
      suite: 'ui-responsiveness-same-corpus-public-access-check',
      corpusCaseId: 'wide-mixed-250k',
      materializedCells: 250_000,
      requestedProductCount: 2,
      verifiedProductCount: 2,
      allRequestedProductsVerified: true,
    })
    expect(check.products.map((entry) => entry.product)).toEqual(['google-sheets', 'microsoft-excel-web'])
    expect(check.products.every((entry) => entry.corpusVerification.checkedCells.length >= 3)).toBe(true)
    expect(() => validateSameCorpusPublicAccessCheck(check)).not.toThrow()
  })

  it('rejects workbooks that do not match the expected same-corpus fingerprint', async () => {
    const otherCorpus = buildWorkbookBenchmarkCorpus('dense-mixed-250k')
    const wrongWorkbookBytes = new Uint8Array(Buffer.from(exportXlsx(otherCorpus.snapshot)))

    await expect(
      buildSameCorpusPublicAccessCheck({
        corpusId: 'wide-mixed-250k',
        fetchXlsxBytes: async () => wrongWorkbookBytes,
        generatedAt: '2026-05-08T00:00:00.000Z',
        googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        microsoftExcelWebUrl: null,
      }),
    ).rejects.toThrow('Same-corpus XLSX is missing sheet: WideGrid')
  })
})
