import { spawnSync } from 'node:child_process'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeSimpleXlsxWorkbook, type SimpleXlsxCell } from '@bilig/xlsx'
import { describe, expect, it } from 'vitest'

import { runWorkPaperXlsxCorpus, runWorkPaperXlsxCorpusInChildProcesses } from '../check-workpaper-xlsx-corpus.ts'
import {
  parseWorkPaperXlsxCorpusCliArgs,
  parseWorkPaperXlsxCorpusInternalCliArgs,
  resolveAllowUnisolatedXlsxCorpus,
} from '../workpaper-xlsx-corpus-cli.ts'

describe('WorkPaper XLSX corpus verifier', () => {
  it('keeps a checked-in XLSX compatibility corpus green', () => {
    const result = runWorkPaperXlsxCorpus([checkedInCorpusDir()])

    expect(result.summary).toMatchObject({
      totalFiles: 2,
      filesProcessed: 2,
      ok: 2,
      failedErrors: 0,
      failedTimeouts: 0,
      formulaCells: 327,
      comparableFormulaCells: 322,
      matchingFormulaCells: 322,
      mismatchedFormulaCells: 0,
      skippedFormulaCells: 5,
      matchRate: 1,
    })
    expect(result.files[0]).toMatchObject({
      fileName: 'issue-8-production-regressions.xlsx',
      status: 'ok',
      formulaCells: 14,
    })
    expect(result.files[1]).toMatchObject({
      fileName: 'macos-excel-threaded-comments-source.xlsx',
      status: 'ok',
      formulaCells: 313,
      skippedFormulaCells: 5,
    })
    expect(result.mismatches).toEqual([])
  })

  it('can isolate each workbook check in a child process with the same parity result', () => {
    const direct = runWorkPaperXlsxCorpus([checkedInCorpusDir()])
    const isolated = runWorkPaperXlsxCorpusInChildProcesses([checkedInCorpusDir()], {
      childProcessTimeoutMs: 10_000,
    })

    expect(isolated.summary).toMatchObject({
      totalFiles: direct.summary.totalFiles,
      filesProcessed: direct.summary.filesProcessed,
      ok: direct.summary.ok,
      failedErrors: 0,
      failedTimeouts: 0,
      formulaCells: direct.summary.formulaCells,
      comparableFormulaCells: direct.summary.comparableFormulaCells,
      matchingFormulaCells: direct.summary.matchingFormulaCells,
      mismatchedFormulaCells: 0,
      skippedFormulaCells: direct.summary.skippedFormulaCells,
      matchRate: 1,
    })
    expect(isolated.mismatches).toEqual([])
  })

  it('matches cached formula results from a production-style XLSX reduction corpus', () => {
    withTempCorpus((corpusDir) => {
      writeWorkbook(join(corpusDir, 'issue-regressions.xlsx'), buildIssueRegressionWorkbookBytes())

      const result = runWorkPaperXlsxCorpus([corpusDir])

      expect(result.summary).toMatchObject({
        totalFiles: 1,
        filesProcessed: 1,
        ok: 1,
        failedTimeouts: 0,
        formulaCells: 9,
        comparableFormulaCells: 9,
        matchingFormulaCells: 9,
        mismatchedFormulaCells: 0,
        skippedFormulaCells: 0,
        matchRate: 1,
      })
      expect(result.files[0]?.status).toBe('ok')
      expect(result.mismatches).toEqual([])
    })
  })

  it('fails oversized workbook files before loading them', () => {
    withTempCorpus((corpusDir) => {
      const workbookPath = join(corpusDir, 'oversized.xlsx')
      writeWorkbook(workbookPath, buildIssueRegressionWorkbookBytes())

      const result = runWorkPaperXlsxCorpus([workbookPath], { maxFileBytes: 1 })

      expect(result.summary).toMatchObject({
        totalFiles: 1,
        filesProcessed: 1,
        ok: 0,
        failedErrors: 1,
        failedTimeouts: 0,
      })
      expect(result.files[0]).toMatchObject({
        fileName: 'oversized.xlsx',
        status: 'error',
        formulaCells: 0,
      })
      expect(result.files[0]?.error).toContain('XLSX file exceeds max file size')
    })
  })

  it('keeps WorkPaper corpus materialization small-workbook only by default', () => {
    withTempCorpus((corpusDir) => {
      const workbookPath = join(corpusDir, 'large-workpaper-default.xlsx')
      writeWorkbook(workbookPath, deterministicBytes(1_100_000))

      const result = runWorkPaperXlsxCorpus([workbookPath], { maxFileBytes: 50 * 1024 * 1024 })

      expect(result.summary.failedErrors).toBe(1)
      expect(result.files[0]).toMatchObject({
        fileName: 'large-workpaper-default.xlsx',
        status: 'error',
        formulaCells: 0,
      })
      expect(result.files[0]?.error).toContain('XLSX file exceeds max file size')
      expect(result.files[0]?.error).toContain('--allow-large-workpaper-materialization')
    })
  })

  it('requires an explicit opt-in before honoring large WorkPaper corpus limits', () => {
    withTempCorpus((corpusDir) => {
      const workbookPath = join(corpusDir, 'large-workpaper-legacy.xlsx')
      writeWorkbook(workbookPath, deterministicBytes(1_100_000))

      const result = runWorkPaperXlsxCorpus([workbookPath], {
        allowLargeWorkPaperMaterialization: true,
        maxFileBytes: 50 * 1024 * 1024,
      })

      expect(result.summary.failedErrors).toBe(1)
      expect(result.files[0]?.error).not.toContain('XLSX file exceeds max file size')
      expect(result.files[0]?.error).not.toContain('--allow-large-workpaper-materialization')
    })
  })

  it('refuses unisolated CLI corpus runs unless explicitly enabled for debugging', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_UNISOLATED_XLSX_CORPUS

    const result = spawnSync('bun', [checkerScriptPath(), '--no-isolate', checkedInCorpusDir()], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).toBe(2)
    expect(result.stderr).toContain('--no-isolate is disabled for corpus CLI runs')
  })

  it('rejects malformed unisolated corpus override values', () => {
    expect(resolveAllowUnisolatedXlsxCorpus({})).toBe(false)
    expect(resolveAllowUnisolatedXlsxCorpus({ BILIG_ALLOW_UNISOLATED_XLSX_CORPUS: '1' })).toBe(true)
    expect(resolveAllowUnisolatedXlsxCorpus({ BILIG_ALLOW_UNISOLATED_XLSX_CORPUS: 'true' })).toBe(true)
    expect(resolveAllowUnisolatedXlsxCorpus({ BILIG_ALLOW_UNISOLATED_XLSX_CORPUS: '0' })).toBe(false)
    expect(resolveAllowUnisolatedXlsxCorpus({ BILIG_ALLOW_UNISOLATED_XLSX_CORPUS: 'false' })).toBe(false)
    expect(() => resolveAllowUnisolatedXlsxCorpus({ BILIG_ALLOW_UNISOLATED_XLSX_CORPUS: 'yes' })).toThrow(
      'BILIG_ALLOW_UNISOLATED_XLSX_CORPUS must be "1", "true", "0", or "false" when set, got yes',
    )
  })

  it('fails malformed unisolated CLI override values before running corpus checks', () => {
    const result = spawnSync('bun', [checkerScriptPath(), '--no-isolate', checkedInCorpusDir()], {
      encoding: 'utf8',
      env: { ...process.env, BILIG_ALLOW_UNISOLATED_XLSX_CORPUS: 'yes' },
    })

    expect(result.status).toBe(2)
    expect(result.stderr).toContain('BILIG_ALLOW_UNISOLATED_XLSX_CORPUS must be "1", "true", "0", or "false" when set, got yes')
  })

  it('refuses unisolated CLI directory sweeps even with the debug escape hatch', () => {
    const result = spawnSync('bun', [checkerScriptPath(), '--no-isolate', checkedInCorpusDir()], {
      encoding: 'utf8',
      env: { ...process.env, BILIG_ALLOW_UNISOLATED_XLSX_CORPUS: '1' },
    })

    expect(result.status).toBe(2)
    expect(result.stderr).toContain('--no-isolate only supports one explicit XLSX file')
  })

  it('refuses broad CLI directory sweeps while the corpus stop marker is active', () => {
    withTempCorpus((tempDir) => {
      const stopMarkerPath = join(tempDir, 'stop.md')
      writeFileSync(stopMarkerPath, 'stop')
      writeWorkbook(join(tempDir, 'issue-regressions.xlsx'), buildIssueRegressionWorkbookBytes())

      const result = spawnSync('bun', [checkerScriptPath(), '--corpus-run-stop-marker', stopMarkerPath, tempDir], {
        encoding: 'utf8',
      })

      expect(result.status).toBe(2)
      expect(result.stderr).toContain(
        'workpaper:xlsx-corpus directory sweep is disabled while the XLSX fixture corpus stop marker is active',
      )
      expect(result.stderr).toContain('--allow-active-stop-marker')
    })
  })

  it('parses decimal min-match-rate CLI values', () => {
    const options = parseWorkPaperXlsxCorpusCliArgs(['--min-match-rate', '0.75', checkedInCorpusFile()])

    expect(options.minMatchRate).toBe(0.75)
  })

  it('parses the explicit large WorkPaper materialization opt-in', () => {
    const cliOptions = parseWorkPaperXlsxCorpusCliArgs([
      '--max-file-bytes',
      String(50 * 1024 * 1024),
      '--allow-large-workpaper-materialization',
      checkedInCorpusFile(),
    ])
    const internalOptions = parseWorkPaperXlsxCorpusInternalCliArgs([
      '--internal-check-file-json',
      checkedInCorpusFile(),
      '--allow-large-workpaper-materialization',
    ])

    expect(cliOptions.allowLargeWorkPaperMaterialization).toBe(true)
    expect(internalOptions.allowLargeWorkPaperMaterialization).toBe(true)
  })

  it('rejects non-decimal min-match-rate CLI values', () => {
    expect(() => parseWorkPaperXlsxCorpusCliArgs(['--min-match-rate', '1e-1', checkedInCorpusFile()])).toThrow(
      '--min-match-rate expects a number between 0 and 1, got 1e-1',
    )
  })

  it('rejects blank string option values', () => {
    expect(() => parseWorkPaperXlsxCorpusCliArgs(['--json-out', '   ', checkedInCorpusFile()])).toThrow('Missing value for --json-out')
  })

  it('rejects blank internal file option values', () => {
    expect(() => parseWorkPaperXlsxCorpusInternalCliArgs(['--internal-check-file-json', '   '])).toThrow(
      'Missing value for --internal-check-file-json',
    )
  })

  it('allows the checked-in fixture corpus directory while the corpus stop marker is active', () => {
    withTempCorpus((tempDir) => {
      const stopMarkerPath = join(tempDir, 'stop.md')
      writeFileSync(stopMarkerPath, 'stop')

      const result = spawnSync('bun', [checkerScriptPath(), '--corpus-run-stop-marker', stopMarkerPath, checkedInCorpusDir()], {
        encoding: 'utf8',
      })

      expect(result.status).toBe(0)
      expect(JSON.parse(result.stdout)).toMatchObject({
        summary: {
          totalFiles: 2,
          filesProcessed: 2,
          ok: 2,
        },
      })
    })
  })

  it('allows a single-file CLI debugger check while the corpus stop marker is active', () => {
    withTempCorpus((tempDir) => {
      const stopMarkerPath = join(tempDir, 'stop.md')
      writeFileSync(stopMarkerPath, 'stop')

      const result = spawnSync('bun', [checkerScriptPath(), '--corpus-run-stop-marker', stopMarkerPath, checkedInCorpusFile()], {
        encoding: 'utf8',
      })

      expect(result.status).toBe(0)
      expect(JSON.parse(result.stdout)).toMatchObject({
        summary: {
          totalFiles: 1,
          filesProcessed: 1,
          ok: 1,
        },
      })
    })
  })

  it('reports actionable mismatch samples with workbook, sheet, address, formula, expected, and actual values', () => {
    withTempCorpus((corpusDir) => {
      writeWorkbook(
        join(corpusDir, 'mismatch.xlsx'),
        writeSimpleXlsxWorkbook({
          sheets: [
            {
              name: 'Sheet1',
              cells: [
                { address: 'A1', row: 0, col: 0, value: 1 },
                { address: 'B1', row: 0, col: 1, formula: 'A1+1', value: 99 },
              ],
            },
          ],
        }),
      )

      const result = runWorkPaperXlsxCorpus([corpusDir])

      expect(result.summary).toMatchObject({
        totalFiles: 1,
        filesProcessed: 1,
        ok: 0,
        formulaCells: 1,
        comparableFormulaCells: 1,
        matchingFormulaCells: 0,
        mismatchedFormulaCells: 1,
        matchRate: 0,
      })
      expect(result.files[0]).toMatchObject({
        fileName: 'mismatch.xlsx',
        status: 'mismatched',
        formulaCells: 1,
        mismatchedFormulaCells: 1,
      })
      expect(result.mismatches[0]).toMatchObject({
        fileName: 'mismatch.xlsx',
        sheetName: 'Sheet1',
        address: 'B1',
        formula: 'A1+1',
        expected: { kind: 'number', value: 99 },
        actual: { kind: 'number', value: 2 },
      })
    })
  })

  it('recalculates imported runtime snapshots before comparing formula values', () => {
    withTempCorpus((corpusDir) => {
      writeWorkbook(
        join(corpusDir, 'whole-column-sumif.xlsx'),
        writeSimpleXlsxWorkbook({
          sheets: [
            {
              name: 'Data',
              cells: [
                { address: 'A1', row: 0, col: 0, value: 'Line' },
                { address: 'B1', row: 0, col: 1, value: 'Amount' },
                { address: 'A2', row: 1, col: 0, value: 'Revenue' },
                { address: 'B2', row: 1, col: 1, value: 1000 },
                { address: 'A3', row: 2, col: 0, value: 'Revenue' },
                { address: 'B3', row: 2, col: 1, value: 1922 },
                { address: 'A4', row: 3, col: 0, value: 'Costs' },
                { address: 'B4', row: 3, col: 1, value: -10 },
              ],
            },
            {
              name: 'Summary',
              cells: [
                { address: 'A1', row: 0, col: 0, value: 'Line' },
                { address: 'B1', row: 0, col: 1, value: 'Total' },
                { address: 'A2', row: 1, col: 0, value: 'Revenue' },
                { address: 'B2', row: 1, col: 1, formula: 'SUMIF(Data!$A:$A,A2,Data!$B:$B)', value: 2922 },
              ],
            },
          ],
        }),
      )

      const result = runWorkPaperXlsxCorpus([corpusDir])

      expect(result.summary).toMatchObject({
        totalFiles: 1,
        ok: 1,
        formulaCells: 1,
        comparableFormulaCells: 1,
        matchingFormulaCells: 1,
        mismatchedFormulaCells: 0,
        skippedFormulaCells: 0,
        matchRate: 1,
      })
      expect(result.mismatches).toEqual([])
    })
  })

  it('uses the XLSX fixture corpus tolerance for tiny floating-point residuals', () => {
    withTempCorpus((corpusDir) => {
      writeWorkbook(
        join(corpusDir, 'tiny-residual.xlsx'),
        writeSimpleXlsxWorkbook({
          sheets: [
            {
              name: 'Sheet1',
              cells: [
                { address: 'A1', row: 0, col: 0, value: 0 },
                { address: 'B1', row: 0, col: 1, value: 7.33325578039512e-9 },
                { address: 'C1', row: 0, col: 2, formula: 'A1-B1', value: 1.7598722479306161e-10 },
              ],
            },
          ],
        }),
      )

      const result = runWorkPaperXlsxCorpus([corpusDir])

      expect(result.summary).toMatchObject({
        ok: 1,
        formulaCells: 1,
        comparableFormulaCells: 1,
        matchingFormulaCells: 1,
        mismatchedFormulaCells: 0,
      })
      expect(result.mismatches).toEqual([])
    })
  })

  it('skips stale cached #NAME? results when recalculation produces a concrete value', () => {
    withTempCorpus((corpusDir) => {
      writeWorkbook(
        join(corpusDir, 'stale-name-cache.xlsx'),
        writeSimpleXlsxWorkbook({
          sheets: [
            {
              name: 'Sheet1',
              cells: [{ address: 'A1', row: 0, col: 0, formula: 'IF(TRUE,"resolved","missing")', error: '#NAME?' }],
            },
          ],
        }),
      )

      const result = runWorkPaperXlsxCorpus([corpusDir])

      expect(result.summary).toMatchObject({
        ok: 1,
        formulaCells: 1,
        comparableFormulaCells: 0,
        matchingFormulaCells: 0,
        mismatchedFormulaCells: 0,
        skippedFormulaCells: 1,
        matchRate: 1,
      })
      expect(result.skippedByReason['stale-cached-name-error']).toBe(1)
      expect(result.mismatches).toEqual([])
    })
  })

  it('counts cached-less and volatile formulas as skipped instead of comparable parity failures', () => {
    withTempCorpus((corpusDir) => {
      writeWorkbook(
        join(corpusDir, 'skipped.xlsx'),
        writeSimpleXlsxWorkbook({
          sheets: [
            {
              name: 'Sheet1',
              cells: [
                { address: 'A1', row: 0, col: 0, formula: 'NOW()', value: 46_127 },
                { address: 'B1', row: 0, col: 1, formula: 'A1+1' },
                { address: 'C1', row: 0, col: 2, formula: 'IMAGE("https://example.com/proof.png")', error: '#VALUE!' },
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
        formulaCells: 3,
        comparableFormulaCells: 0,
        matchingFormulaCells: 0,
        mismatchedFormulaCells: 0,
        skippedFormulaCells: 3,
        matchRate: 1,
      })
      expect(result.files[0]?.skippedFormulaCells).toBe(3)
      expect(result.skippedByReason).toEqual({
        'missing-cached-result': 1,
        'stale-cached-result': 0,
        'stale-cached-name-error': 0,
        'unsupported-cached-result-type': 0,
        'volatile-or-environment-dependent-formula': 2,
      })
    })
  })

  it('falls back to formula audit for empty-cache formulas', () => {
    withTempCorpus((corpusDir) => {
      writeWorkbook(
        join(corpusDir, 'empty-cache-formula.xlsx'),
        writeSimpleXlsxWorkbook({
          sheets: [
            {
              name: 'Sheet1',
              cells: [
                { address: 'A1', row: 0, col: 0, value: 1 },
                { address: 'B1', row: 0, col: 1, formula: 'A1+1' },
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
        formulaCells: 1,
        comparableFormulaCells: 0,
        matchingFormulaCells: 0,
        mismatchedFormulaCells: 0,
        skippedFormulaCells: 1,
        matchRate: 1,
      })
      expect(result.skippedByReason['missing-cached-result']).toBe(1)
      expect(result.mismatches).toEqual([])
    })
  })

  it('skips cached formulas that transitively depend on volatile formulas', () => {
    withTempCorpus((corpusDir) => {
      writeWorkbook(
        join(corpusDir, 'volatile-dependents.xlsx'),
        writeSimpleXlsxWorkbook({
          sheets: [
            {
              name: 'Sheet1',
              cells: [
                { address: 'A1', row: 0, col: 0, formula: 'RAND()', value: 0.25 },
                { address: 'B1', row: 0, col: 1, formula: 'A1:A2*2', value: 0.5 },
                { address: 'C1', row: 0, col: 2, formula: 'B1+1', value: 1.5 },
                { address: 'A2', row: 1, col: 0, value: 2 },
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
        formulaCells: 3,
        comparableFormulaCells: 0,
        matchingFormulaCells: 0,
        mismatchedFormulaCells: 0,
        skippedFormulaCells: 3,
        matchRate: 1,
      })
      expect(result.skippedByReason).toEqual({
        'missing-cached-result': 0,
        'stale-cached-result': 0,
        'stale-cached-name-error': 0,
        'unsupported-cached-result-type': 0,
        'volatile-or-environment-dependent-formula': 3,
      })
    })
  })
})

function checkedInCorpusDir(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../../packages/workpaper/fixtures/xlsx-corpus')
}

function checkedInCorpusFile(): string {
  return join(checkedInCorpusDir(), 'issue-8-production-regressions.xlsx')
}

function checkerScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../check-workpaper-xlsx-corpus.ts')
}

function withTempCorpus(run: (corpusDir: string) => void): void {
  const corpusDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-xlsx-corpus-'))
  try {
    run(corpusDir)
  } finally {
    rmSync(corpusDir, { recursive: true, force: true })
  }
}

function writeWorkbook(path: string, workbook: Uint8Array): void {
  writeFileSync(path, workbook)
}

function deterministicBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length)
  for (let index = 0; index < bytes.length; index += 1) {
    bytes[index] = (index * 31 + 17) & 0xff
  }
  return bytes
}

function buildIssueRegressionWorkbookBytes(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Summary',
        dimension: { s: { r: 0, c: 0 }, e: { r: 13, c: 3 } },
        cells: [
          textCell('A1', 0, 0, 'Metric'),
          textCell('B1', 0, 1, 'Value'),
          textCell('C1', 0, 2, 'Lookup key'),
          textCell('D1', 0, 3, 'Lookup value'),
          textCell('A2', 1, 0, 'Deposits'),
          { address: 'B2', row: 1, col: 1, formula: 'SUMIFS(Activity!$B$2:$B$4,Activity!$A$2:$A$4,"Deposit")', value: 3500 },
          textCell('A3', 2, 0, 'Deposit check'),
          { address: 'B3', row: 2, col: 1, formula: 'IF(ABS(B2-3500)<0.01,"PASS","FAIL")', value: 'PASS' },
          { address: 'C3', row: 2, col: 2, formula: '1/0', error: '#DIV/0!' },
          textCell('A4', 3, 0, 'Activity rows'),
          { address: 'B4', row: 3, col: 1, formula: 'COUNTA(Activity!$A$2:$A$4)', value: 3 },
          textCell('A5', 4, 0, 'Internal link'),
          { address: 'B5', row: 4, col: 1, formula: 'HYPERLINK("#\'Summary\'!A1","Go to Summary")', value: 'Go to Summary' },
          textCell('A6', 5, 0, 'Formatted date'),
          { address: 'B6', row: 5, col: 1, formula: 'TEXT(46127,"mm.dd.yy")', value: '04.15.26' },
          textCell('A7', 6, 0, 'Day'),
          { address: 'B7', row: 6, col: 1, formula: 'DAY(46127)', value: 15 },
          textCell('A8', 7, 0, 'Workday'),
          { address: 'B8', row: 7, col: 1, formula: 'WORKDAY(46127,2)', value: 46_129 },
          textCell('A14', 13, 0, 'Bank lookup'),
          textCell('C14', 13, 2, 'txn-123'),
          {
            address: 'D14',
            row: 13,
            col: 3,
            formula: 'IFERROR(INDEX(Bank!$B$2:$B$31,MATCH(C14,Bank!$D$2:$D$31,0)),"")',
            value: '2026-04-01',
          },
        ],
      },
      {
        name: 'Activity',
        cells: [
          textCell('A1', 0, 0, 'Type'),
          textCell('B1', 0, 1, 'Amount'),
          textCell('A2', 1, 0, 'Deposit'),
          { address: 'B2', row: 1, col: 1, value: 3500 },
          textCell('A3', 2, 0, 'Fee'),
          { address: 'B3', row: 2, col: 1, value: -18.5 },
          textCell('A4', 3, 0, 'Withdrawal'),
          { address: 'B4', row: 3, col: 1, value: -250 },
        ],
      },
      {
        name: 'Bank',
        dimension: { s: { r: 0, c: 0 }, e: { r: 30, c: 3 } },
        cells: [
          textCell('A1', 0, 0, 'Date label'),
          textCell('B1', 0, 1, 'Date'),
          textCell('C1', 0, 2, 'Description'),
          textCell('D1', 0, 3, 'Transaction ID'),
          textCell('A2', 1, 0, 'Posted'),
          textCell('B2', 1, 1, '2026-04-01'),
          textCell('C2', 1, 2, 'Deposit'),
          textCell('D2', 1, 3, 'txn-123'),
        ],
      },
    ],
  })
}

function textCell(address: string, row: number, col: number, value: string): SimpleXlsxCell {
  return { address, row, col, value }
}
