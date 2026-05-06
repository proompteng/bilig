import { describe, expect, it } from 'vitest'

import { exportXlsx } from '../../packages/excel-import/src/index.js'
import { buildWorkbookBenchmarkCorpus } from '../../packages/benchmarks/src/workbook-corpus.js'
import {
  buildSameCorpusFingerprint,
  parseCaptureArgs,
  parseEmitXlsxArgs,
  parseSaveStorageStateArgs,
  verifyXlsxCorpusFingerprint,
} from '../capture-ui-responsiveness-same-corpus.ts'

describe('same-corpus UI responsiveness capture CLI', () => {
  it('builds a default Bilig benchmark URL from the selected corpus', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--microsoft-excel-web-url',
      'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      '--corpus',
      'dense-mixed-250k',
    ])

    expect(args).toMatchObject({
      biligUrl: 'http://127.0.0.1:5173/?benchmarkCorpus=dense-mixed-250k',
      biligStorageStatePath: null,
      corpusId: 'dense-mixed-250k',
      deltaX: 0,
      deltaY: 720,
      googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      googleSheetsStorageStatePath: null,
      headless: true,
      microsoftExcelWebUrl: 'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      microsoftExcelWebStorageStatePath: null,
      readyTimeoutMs: 60000,
      sampleCount: 3,
      storageStatePath: null,
    })
    expect(args.outputPath.endsWith('/tmp/ui-capture.json')).toBe(true)
  })

  it('accepts explicit browser and workload options', () => {
    const args = parseCaptureArgs([
      '--output',
      'tmp/ui-capture.json',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--microsoft-excel-web-url',
      'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
      '--bilig-url',
      'http://127.0.0.1:4173/?benchmarkCorpus=wide-mixed-250k',
      '--samples',
      '5',
      '--delta-x',
      '1024',
      '--delta-y',
      '0',
      '--ready-timeout-ms',
      '120000',
      '--storage-state',
      'tmp/shared-state.json',
      '--google-sheets-storage-state',
      'tmp/google-state.json',
      '--microsoft-excel-web-storage-state',
      'tmp/microsoft-state.json',
      '--bilig-storage-state',
      'tmp/bilig-state.json',
      '--headed',
    ])

    expect(args).toMatchObject({
      biligUrl: 'http://127.0.0.1:4173/?benchmarkCorpus=wide-mixed-250k',
      deltaX: 1024,
      deltaY: 0,
      headless: false,
      readyTimeoutMs: 120000,
      sampleCount: 5,
    })
    expect(args.storageStatePath?.endsWith('/tmp/shared-state.json')).toBe(true)
    expect(args.googleSheetsStorageStatePath?.endsWith('/tmp/google-state.json')).toBe(true)
    expect(args.microsoftExcelWebStorageStatePath?.endsWith('/tmp/microsoft-state.json')).toBe(true)
    expect(args.biligStorageStatePath?.endsWith('/tmp/bilig-state.json')).toBe(true)
  })

  it('rejects missing incumbent URLs because the generated proof must be comparable', () => {
    expect(() => parseCaptureArgs(['--output', 'tmp/ui-capture.json'])).toThrow('Missing required arguments.')
  })

  it('parses XLSX emission mode for same-corpus setup', () => {
    const args = parseEmitXlsxArgs(['--emit-xlsx', 'tmp/ui-corpus', '--corpus', 'wide-mixed-variable-250k'])

    expect(args).toMatchObject({
      check: false,
      corpusId: 'wide-mixed-variable-250k',
    })
    expect(args?.targetDirectory.endsWith('/tmp/ui-corpus')).toBe(true)
  })

  it('parses checked XLSX fixture mode', () => {
    const args = parseEmitXlsxArgs(['--emit-xlsx', 'packages/benchmarks/baselines/ui-same-corpus', '--check'])

    expect(args).toMatchObject({
      check: true,
      corpusId: 'wide-mixed-250k',
    })
    expect(args?.targetDirectory.endsWith('/packages/benchmarks/baselines/ui-same-corpus')).toBe(true)
  })

  it('builds deterministic literal-cell fingerprints for same-corpus verification', () => {
    const corpus = buildWorkbookBenchmarkCorpus('wide-mixed-250k')
    const fingerprint = buildSameCorpusFingerprint(corpus)

    expect(fingerprint).toMatchObject({
      sheetName: 'WideGrid',
      materializedCells: 250000,
    })
    expect(fingerprint.checkedCells.length).toBeGreaterThanOrEqual(3)
    expect(fingerprint.checkedCells[0]).toEqual({ address: 'A1', expected: 'metric-1' })
  })

  it('verifies same-corpus XLSX bytes before accepting external timing evidence', () => {
    const corpus = buildWorkbookBenchmarkCorpus('wide-mixed-250k')
    const verification = verifyXlsxCorpusFingerprint(Buffer.from(exportXlsx(corpus.snapshot)), corpus, 'microsoft-excel-web-source-xlsx')

    expect(verification).toMatchObject({
      verified: true,
      method: 'microsoft-excel-web-source-xlsx',
      sheetName: 'WideGrid',
      materializedCells: 250000,
    })
    expect(verification.checkedCells.length).toBeGreaterThanOrEqual(3)
    expect(verification.checkedCells.every((cell) => cell.expected === cell.actual)).toBe(true)
  })

  it('parses storage-state bootstrap mode for authenticated capture', () => {
    const args = parseSaveStorageStateArgs([
      '--save-storage-state',
      'tmp/google-state.json',
      '--auth-product',
      'google-sheets',
      '--google-sheets-url',
      'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      '--ready-timeout-ms',
      '180000',
    ])

    expect(args).toMatchObject({
      authUrl: 'https://docs.google.com/spreadsheets/d/sheet-id/edit',
      corpusId: 'wide-mixed-250k',
      headless: false,
      product: 'google-sheets',
      readyTimeoutMs: 180000,
    })
    expect(args?.targetPath.endsWith('/tmp/google-state.json')).toBe(true)
  })

  it('requires an auth URL in storage-state bootstrap mode', () => {
    expect(() => parseSaveStorageStateArgs(['--save-storage-state', 'tmp/state.json'])).toThrow('Missing auth URL.')
  })

  it('rejects unknown corpus ids', () => {
    expect(() =>
      parseCaptureArgs([
        '--output',
        'tmp/ui-capture.json',
        '--google-sheets-url',
        'https://docs.google.com/spreadsheets/d/sheet-id/edit',
        '--microsoft-excel-web-url',
        'https://view.officeapps.live.com/op/view.aspx?src=example.xlsx',
        '--corpus',
        'tiny-demo',
      ]),
    ).toThrow('Unexpected workbook benchmark corpus id: tiny-demo')
  })
})
