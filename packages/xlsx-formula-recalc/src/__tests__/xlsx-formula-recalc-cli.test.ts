import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

import { describe, expect, it } from 'vitest'

import { runXlsxFormulaRecalcCli, runXlsxFormulaRecalcCliAsync } from '../cli-api.js'
import {
  buildWorkbookCompatibilityReport,
  buildWorkbookCompatibilityReportFromFile,
  runWorkbookCompatibilityReportCli,
} from '../workbook-compatibility-report.js'
import {
  buildExternalLinkRangeCacheWorkbook,
  buildExternalSourceWorkbook,
  buildManyFormulaCacheWorkbook,
  buildMissingFormulaCacheWorkbook,
  buildProviderBackedRiskWorkbook,
  buildStaleFormulaCacheWorkbook,
  packageVersion,
  readCachedFormulaValue,
  readCliErrorSummary,
  readCliInspectionSummary,
  readCliSummary,
  readExternalLinkCacheCellValue,
  readFileBytes,
  readGeneratedWorkflow,
  readNoSheetJsChildOutput,
  readWorkbookCompatibilityReport,
  requireRecord,
  requireRecordArray,
} from './xlsx-formula-recalc-cli-test-helpers.js'

describe('xlsx-recalc CLI', () => {
  it('runs a one-command demo that writes a recalculated XLSX and prints proof JSON', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-'))
    try {
      const outputPath = join(tempDir, 'demo.recalculated.xlsx')
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli(['--demo', '--out', outputPath, '--json'], {
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      const summary = readCliSummary(stdout)
      expect(summary.mode).toBe('demo')
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.recalculationCompleted).toBe(true)
      expect(summary.expectedValueMatched).toBe(true)
      expect(summary.expectedReadback).toEqual({ 'Summary!B2': 72_000 })
      expect(summary.excelParity).toBe('not_proven')
      expect(summary).not.toHaveProperty('star')
      expect(summary).not.toHaveProperty('watchReleases')
      expect(summary).not.toHaveProperty('adoptionBlocker')
      expect(summary.reads['Summary!B2']?.value).toBe(72_000)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('keeps human CLI output focused on recalculation results', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-human-'))
    try {
      const outputPath = join(tempDir, 'demo.recalculated.xlsx')
      let stdout = ''

      const exitCode = runXlsxFormulaRecalcCli(['--demo', '--out', outputPath, '--read', 'Summary!B2'], {
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      expect(stdout).toContain('Recalculated generated demo workbook ->')
      expect(stdout).toContain('Summary!B2:')
      expect(stdout).not.toContain('star or bookmark')
      expect(stdout).not.toContain('adoption blocker')
      expect(stdout).not.toContain('Watch formula')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('uses streaming-native for async file-to-file recalculation by default', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-native-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--out', outputPath, '--read', 'Sheet1!B2', '--json'], {
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      const summary = readCliSummary(stdout)
      expect(summary.reads['Sheet1!B2']?.value).toBe(20)
      expect(summary.diagnostics.engineMode).toBe('streaming-native')
      expect(summary.diagnostics.fallbackUsed).toBe(false)
      expect(readCachedFormulaValue(readFileSync(outputPath), 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('does not load SheetJS xlsx for streaming-native file-to-file recalculation', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-native-no-sheetjs-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      const cliApiUrl = pathToFileURL(join(process.cwd(), 'packages/xlsx-formula-recalc/src/cli-api.ts')).href
      const script = `
void (async () => {
const { createRequire } = require('node:module')
const requireForCache = createRequire(process.cwd() + '/package.json')
const loadedXlsxModules = () =>
  Object.keys(requireForCache.cache).filter((path) => /(?:^|[\\\\/])xlsx(?:[\\\\/]|$)|[\\\\/]\\.pnpm[\\\\/]xlsx@/u.test(path))
const before = loadedXlsxModules()
const { runXlsxFormulaRecalcCliAsync } = await import(${JSON.stringify(cliApiUrl)})
let stdout = ''
let stderr = ''
const exitCode = await runXlsxFormulaRecalcCliAsync(
  [
    ${JSON.stringify(inputPath)},
    '--out',
    ${JSON.stringify(outputPath)},
    '--read',
    'Sheet1!B2',
    '--engine',
    'streaming-native',
    '--fallback-policy',
    'error',
    '--json',
  ],
  {
    stdout: (text) => {
      stdout += text
    },
    stderr: (text) => {
      stderr += text
    },
  },
)
process.stdout.write(JSON.stringify({ exitCode, stderr, before, after: loadedXlsxModules(), summary: JSON.parse(stdout) }) + '\\n')
})().catch((error) => {
  console.error(error)
  process.exit(1)
})
`
      const result = spawnSync('pnpm', ['exec', 'tsx', '--eval', script], {
        cwd: process.cwd(),
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'pipe'],
      })

      expect(result.status, result.stderr).toBe(0)
      const output = readNoSheetJsChildOutput(result.stdout)
      expect(output.exitCode, output.stderr).toBe(0)
      expect(output.before).toEqual([])
      expect(output.after).toEqual([])
      expect(output.summary.diagnostics?.engineMode).toBe('streaming-native')
      expect(output.summary.diagnostics?.fallbackUsed).toBe(false)
      expect(output.summary.reads['Sheet1!B2']?.value).toBe(20)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses synchronous file-to-file recalculation so callers use the file-backed native path', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-sync-file-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stderr = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--out', outputPath, '--read', 'Sheet1!B2', '--json'], {
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(existsSync(outputPath)).toBe(false)
      expect(stderr).toContain('runXlsxFormulaRecalcCliAsync')
      expect(stderr).toContain('file-backed streaming-native engine')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses synchronous file inspection so callers use the file-backed native inspector', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-sync-inspect-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stderr = ''

      const exitCode = runXlsxFormulaRecalcCli([inputPath, '--inspect', '--json'], {
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(stderr).toContain('runXlsxFormulaRecalcCliAsync')
      expect(stderr).toContain('file-backed streaming-native inspector')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses synchronous file inspection before reading input bytes', () => {
    let stderr = ''

    const exitCode = runXlsxFormulaRecalcCli(['/tmp/bilig-missing-large.xlsx', '--inspect', '--json'], {
      stderr: (text) => {
        stderr += text
      },
    })

    expect(exitCode).toBe(1)
    expect(stderr).toContain('runXlsxFormulaRecalcCliAsync')
    expect(stderr).not.toContain('ENOENT')
  })

  it('refuses WorkPaper engine selection on the primary async file CLI', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-workpaper-engine-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''
      let stderr = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--out', outputPath, '--engine', 'workpaper', '--json'], {
        stdout: (text) => {
          stdout += text
        },
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(stderr).toBe('')
      expect(existsSync(outputPath)).toBe(false)
      const summary = readCliErrorSummary(stdout)
      expect(summary.commandSucceeded).toBe(false)
      expect(summary.recalculationCompleted).toBe(false)
      expect(summary.error).toContain('no longer loads or exports WorkPaper')
      expect(summary.error).toContain('@bilig/workpaper')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses WorkPaper fallback policy on the primary async file CLI', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-workpaper-fallback-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''
      let stderr = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--out', outputPath, '--fallback-policy', 'workpaper', '--json'], {
        stdout: (text) => {
          stdout += text
        },
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(stderr).toBe('')
      expect(existsSync(outputPath)).toBe(false)
      const summary = readCliErrorSummary(stdout)
      expect(summary.commandSucceeded).toBe(false)
      expect(summary.recalculationCompleted).toBe(false)
      expect(summary.error).toContain('no longer loads or exports WorkPaper')
      expect(summary.error).toContain('@bilig/workpaper')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('rejects the retired timeout option before invoking the native file API', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-timeout-'))
    try {
      const inputPath = join(tempDir, 'native.xlsx')
      const outputPath = join(tempDir, 'native.recalculated.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''
      let stderr = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--timeout-ms', '1000', '--out', outputPath, '--json'], {
        stdout: (text) => {
          stdout += text
        },
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(stderr).toBe('')
      expect(existsSync(outputPath)).toBe(false)
      const summary = readCliErrorSummary(stdout)
      expect(summary.commandSucceeded).toBe(false)
      expect(summary.recalculationCompleted).toBe(false)
      expect(summary.error).toContain('does not support --timeout-ms')
      expect(summary.error).toContain('@bilig/workpaper')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('inspects workbook formula cells through streaming-native before writing an output file', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-inspect-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--inspect', '--json'], {
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(join(tempDir, 'stale-cache.recalculated.xlsx'))).toBe(false)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.schemaVersion).toBe('xlsx-cache-doctor.v1')
      expect(summary.mode).toBe('file')
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.inspectionCompleted).toBe(true)
      expect(summary.recalculationCompleted).toBe(true)
      expect(summary.excelParity).toBe('not_proven')
      expect(summary.formulaCellCount).toBe(1)
      expect(summary.inspectedFormulaCellCount).toBe(1)
      expect(summary.uninspectedFormulaCellCount).toBe(0)
      expect(summary.inspectionLimit).toBe(2000)
      expect(summary.staleCachedFormulaCount).toBe(1)
      expect(summary.cacheStatusSummary).toEqual({
        inspected: 1,
        stale: 1,
        fresh: 0,
        missingCache: 0,
        unsupportedRecalculation: 0,
      })
      expect(summary.suggestedReads).toEqual(['Sheet1!B2'])
      expect(summary.formulas[0]).toMatchObject({
        target: 'Sheet1!B2',
        formula: '=A2*10',
        cachedValue: 999,
        literalRecalculatedValue: 20,
        cacheStatus: 'stale',
        staleCachedValue: true,
      })
      expect(JSON.parse(stdout).diagnostics).toMatchObject({ engineMode: 'streaming-native', fallbackUsed: false })
      expect(JSON.parse(stdout)).not.toHaveProperty('nextStep')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('runs xlsx-cache-doctor as the default streaming-native inspection command', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-cli-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(join(tempDir, 'stale-cache.recalculated.xlsx'))).toBe(false)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.schemaVersion).toBe('xlsx-cache-doctor.v1')
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.inspectionCompleted).toBe(true)
      expect(summary.staleCachedFormulaCount).toBe(1)
      expect(summary.cacheStatusSummary.stale).toBe(1)
      expect(summary.uninspectedFormulaCellCount).toBe(0)
      expect(summary.suggestedReads).toEqual(['Sheet1!B2'])
      expect(JSON.parse(stdout).diagnostics).toMatchObject({ engineMode: 'streaming-native', fallbackUsed: false })
      expect(JSON.parse(stdout)).not.toHaveProperty('nextStep')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('runs a cache-doctor demo that proves a stale cached formula value', () => {
    let stdout = ''

    const exitCode = runXlsxFormulaRecalcCli(['--demo', '--json'], {
      commandName: 'xlsx-cache-doctor',
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const summary = readCliInspectionSummary(stdout)
    expect(summary.schemaVersion).toBe('xlsx-cache-doctor.v1')
    expect(summary.mode).toBe('demo')
    expect(summary.commandSucceeded).toBe(true)
    expect(summary.inspectionCompleted).toBe(true)
    expect(summary.formulaCellCount).toBe(1)
    expect(summary.inspectedFormulaCellCount).toBe(1)
    expect(summary.uninspectedFormulaCellCount).toBe(0)
    expect(summary.staleCachedFormulaCount).toBe(1)
    expect(summary.cacheStatusSummary).toEqual({
      inspected: 1,
      stale: 1,
      fresh: 0,
      missingCache: 0,
      unsupportedRecalculation: 0,
    })
    expect(summary.suggestedReads).toEqual(['Summary!B2'])
    expect(summary.formulas[0]).toMatchObject({
      target: 'Summary!B2',
      formula: '=Inputs!B2*Inputs!B3',
      cachedValue: 60_000,
      literalRecalculatedValue: 72_000,
      cacheStatus: 'stale',
      staleCachedValue: true,
    })
  })

  it('prints a workbook compatibility report without claiming an Excel compatibility score', () => {
    let stdout = ''

    const exitCode = runWorkbookCompatibilityReportCli(['--demo', '--json'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    const report = readWorkbookCompatibilityReport(stdout)
    expect(report.schemaVersion).toBe('bilig-workbook-compatibility-report.v1')
    expect(report.verified).toBe(true)
    expect(report.risk.level).toBe('high')
    expect(report.workbook.formulaCellCount).toBe(3)
    expect(report.findings.unsupportedFunctions).toEqual([{ name: 'CUBEVALUE', count: 1 }])
    expect(report.findings.volatileFunctions).toEqual([{ name: 'NOW', count: 1 }])
    expect(report.findings.staleCachedFormulas.count).toBe(1)
    expect(report.findings.missingCachedFormulaValues.count).toBe(1)
    expect(report.findings.unsupportedRecalculations.count).toBe(1)
    expect(report.recalculationCompleted).toBe(false)
    expect(report.excelParity).toBe('not_proven')
    expect(report.limitations).toContain('It is not an Excel compatibility certification.')
    expect(report.limitations).toContain(
      'It scans workbook package metadata and formula caches; use xlsx-cache-doctor for native recalculation proof.',
    )
    expect(stdout).not.toMatch(/compatibilityScore|excelCompatibilityPercent/u)
  })

  it('prints a file-backed workbook compatibility report without full workbook materialization', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'workbook-compatibility-file-native-'))
    try {
      const inputPath = join(tempDir, 'provider-backed-risk.xlsx')
      writeFileSync(inputPath, buildProviderBackedRiskWorkbook())
      let stdout = ''

      const exitCode = runWorkbookCompatibilityReportCli([inputPath, '--json'], {
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      const report = readWorkbookCompatibilityReport(stdout)
      expect(report.schemaVersion).toBe('bilig-workbook-compatibility-report.v1')
      expect(report.verified).toBe(true)
      expect(report.cacheInspection.inspectionLimit).toBe(2000)
      expect(report.cacheInspection.uninspectedFormulaCellCount).toBe(0)
      expect(report.workbook.formulaCellCount).toBe(5)
      expect(report.diagnostics).toMatchObject({
        engineMode: 'streaming-native',
        fallbackUsed: false,
        inputBytes: buildProviderBackedRiskWorkbook().byteLength,
        sheetCount: 1,
        targetRowCount: 5,
        editCount: 0,
        readCount: 5,
        patchedCacheCount: 0,
        unsupportedReason: 'unsupported functions: GOOGLEFINANCE (1), IMPORTDATA (1), IMPORTHTML (1), IMPORTRANGE (1), TRANSLATE (1)',
      })
      expect(report.diagnostics.maxObservedRssBytes).toBeGreaterThan(0)
      expect(report.diagnostics.phaseRssPeaks.length).toBeGreaterThan(0)
      expect(report.diagnostics.formulaCounts).toMatchObject({
        scannedFormulaCellCount: 5,
        targetedFormulaCellCount: 5,
        evaluatedFormulaCellCount: 0,
        patchedFormulaCacheCount: 0,
        unsupportedFormulaCellCount: 5,
        nativeKernelFormulaCellCount: 0,
        nativeKernelBatchCount: 0,
      })
      expect(report.findings.unsupportedFunctions).toEqual([
        { name: 'GOOGLEFINANCE', count: 1 },
        { name: 'IMPORTDATA', count: 1 },
        { name: 'IMPORTHTML', count: 1 },
        { name: 'IMPORTRANGE', count: 1 },
        { name: 'TRANSLATE', count: 1 },
      ])
      expect(report.risk.level).toBe('high')
      expect(report.recalculationCompleted).toBe(false)
      expect(stdout).not.toMatch(/compatibilityScore|excelCompatibilityPercent/u)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('keeps workbook compatibility human output explicit about the trust boundary', () => {
    let stdout = ''

    const exitCode = runWorkbookCompatibilityReportCli(['--demo'], {
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('Workbook analyzed. Risk level: HIGH')
    expect(stdout).toContain('Unsupported functions: CUBEVALUE (1)')
    expect(stdout).toContain('It is not an Excel compatibility certification.')
    expect(stdout).not.toContain('compatibility score')
  })

  it('flags provider-backed workbook formulas as high-risk unsupported functions', () => {
    const report = buildWorkbookCompatibilityReport(buildProviderBackedRiskWorkbook(), {
      fileName: 'provider-backed-risk.xlsx',
    })

    expect(report.risk.level).toBe('high')
    expect(report.findings.unsupportedFunctions).toEqual([
      { name: 'GOOGLEFINANCE', count: 1 },
      { name: 'IMPORTDATA', count: 1 },
      { name: 'IMPORTHTML', count: 1 },
      { name: 'IMPORTRANGE', count: 1 },
      { name: 'TRANSLATE', count: 1 },
    ])
    expect(report.risk.reasons.join('\n')).toContain('unsupported functions:')
  })

  it('counts external workbook links at workbook reference grain', () => {
    const report = buildWorkbookCompatibilityReport(buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx'), {
      fileName: 'external-link-risk.xlsx',
    })

    expect(report.findings.externalLinks.count).toBe(1)
    expect(report.findings.externalLinks.unresolvedCount).toBe(0)
    expect(report.risk.level).toBe('high')
    expect(report.risk.reasons).toContain('external workbook links: 1')
  })

  it('keeps workbook compatibility byte-buffer reports small-workbook only', () => {
    expect(() => buildWorkbookCompatibilityReport(new Uint8Array(1_000_001), { fileName: 'large.xlsx' })).toThrow(
      /buildWorkbookCompatibilityReport is small-workbook only/u,
    )
  })

  it('counts large workbook compatibility external workbook CLI inputs without byte hydration', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'workbook-compatibility-external-limit-'))
    try {
      const externalPath = join(tempDir, 'large-external.xlsx')
      writeFileSync(externalPath, Buffer.alloc(1_000_001))
      let stdout = ''
      let stderr = ''

      const exitCode = runWorkbookCompatibilityReportCli(['--demo', '--external-workbook', externalPath, '--json'], {
        stdout: (text) => {
          stdout += text
        },
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(0)
      expect(stderr).toBe('')
      expect(requireRecord(requireRecord(JSON.parse(stdout))['input'])['externalWorkbookCount']).toBe(1)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('raises risk when inspection limits leave formulas unchecked', () => {
    const report = buildWorkbookCompatibilityReport(buildManyFormulaCacheWorkbook(), {
      fileName: 'limited-inspection.xlsx',
      inspectLimit: 50,
    })

    expect(report.cacheInspection).toMatchObject({
      inspectedFormulaCellCount: 50,
      uninspectedFormulaCellCount: 10,
      inspectionLimit: 50,
    })
    expect(report.risk.level).toBe('medium')
    expect(report.risk.reasons).toContain('uninspected formula cells: 10')
  })

  it('keeps file-backed workbook compatibility inspection limits consistent with the bytes API', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'workbook-compatibility-file-limit-'))
    try {
      const inputPath = join(tempDir, 'limited-inspection.xlsx')
      writeFileSync(inputPath, buildManyFormulaCacheWorkbook())

      const report = buildWorkbookCompatibilityReportFromFile(inputPath, {
        inspectLimit: 50,
      })

      expect(report.cacheInspection).toMatchObject({
        inspectedFormulaCellCount: 50,
        uninspectedFormulaCellCount: 10,
        inspectionLimit: 50,
      })
      expect(report.risk.level).toBe('medium')
      expect(report.risk.reasons).toContain('uninspected formula cells: 10')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('separates missing cached formula values from stale cached values', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-missing-cache-'))
    try {
      const inputPath = join(tempDir, 'missing-cache.xlsx')
      writeFileSync(inputPath, buildMissingFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.staleCachedFormulaCount).toBe(0)
      expect(summary.cacheStatusSummary).toEqual({
        inspected: 1,
        stale: 0,
        fresh: 0,
        missingCache: 1,
        unsupportedRecalculation: 0,
      })
      expect(summary.formulas[0]).toMatchObject({
        target: 'Sheet1!B2',
        formula: '=A2*10',
        literalRecalculatedValue: 20,
        cacheStatus: 'missing-cache',
        staleCachedValue: null,
      })
      expect(summary.formulas[0]).not.toHaveProperty('cachedValue')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('prints a ready-to-commit GitHub Actions workflow for cache doctor adoption', () => {
    let stdout = ''

    const exitCode = runXlsxFormulaRecalcCli(
      [
        '--print-github-action',
        'fixtures/pricing model.xlsx',
        '--fail-on-stale',
        'false',
        '--inspect-limit',
        '50',
        '--json-output',
        '${{ runner.temp }}/custom-cache-doctor.json',
        '--markdown-output',
        '${{ runner.temp }}/custom-cache-doctor.md',
        '--package-version',
        '0.124.1',
        '--workflow-name',
        'workbook cache doctor',
      ],
      {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      },
    )

    expect(exitCode).toBe(0)
    expect(stdout).toContain('name: "workbook cache doctor"')
    expect(stdout).toContain('pull_request:')
    expect(stdout).toContain('- "**/*.xlsx"')
    expect(stdout).toContain('fetch-depth: 0')
    expect(stdout).toContain('uses: actions/setup-node@v6')
    expect(stdout).toContain('node-version: "22"')
    expect(stdout).toContain('package-manager-cache: false')
    expect(stdout).toContain('uses: proompteng/bilig@v1')
    expect(stdout).toContain('workbooks: "fixtures/pricing model.xlsx"')
    expect(stdout).toContain('changed-files-only: "true"')
    expect(stdout).toContain('inspect-limit: "50"')
    expect(stdout).toContain('json-output: "${{ runner.temp }}/custom-cache-doctor.json"')
    expect(stdout).toContain('markdown-output: "${{ runner.temp }}/custom-cache-doctor.md"')
    expect(stdout).toContain('package-version: "0.124.1"')
    expect(stdout).toContain('fail-on-stale: "false"')
    expect(stdout).toContain('name: xlsx-cache-doctor-report')

    const workflow = readGeneratedWorkflow(stdout)
    const jobs = requireRecord(workflow['jobs'])
    const job = requireRecord(jobs['inspect-xlsx-formula-caches'])
    const steps = requireRecordArray(job['steps'])
    const checkout = steps[0] ?? {}
    const setupNode = steps[1] ?? {}
    const cacheDoctor = steps[2] ?? {}
    const cacheDoctorInputs = requireRecord(cacheDoctor['with'])
    const uploadArtifact = steps[3] ?? {}

    expect(workflow['name']).toBe('workbook cache doctor')
    expect(workflow['on']).toEqual({
      pull_request: { paths: ['**/*.xlsx'] },
      workflow_dispatch: null,
    })
    expect(workflow['permissions']).toEqual({ contents: 'read' })
    expect(checkout).toMatchObject({ uses: 'actions/checkout@v5', with: { 'fetch-depth': 0 } })
    expect(setupNode).toMatchObject({
      uses: 'actions/setup-node@v6',
      with: { 'node-version': '22', 'package-manager-cache': false },
    })
    expect(cacheDoctor).toMatchObject({ id: 'cache-doctor', uses: 'proompteng/bilig@v1' })
    expect(cacheDoctorInputs).toEqual({
      workbooks: 'fixtures/pricing model.xlsx',
      'changed-files-only': 'true',
      'package-version': '0.124.1',
      'inspect-limit': '50',
      'json-output': '${{ runner.temp }}/custom-cache-doctor.json',
      'markdown-output': '${{ runner.temp }}/custom-cache-doctor.md',
      'fail-on-stale': 'false',
    })
    expect(uploadArtifact).toMatchObject({
      uses: 'actions/upload-artifact@v4',
      if: 'always()',
      with: {
        name: 'xlsx-cache-doctor-report',
        path: '${{ steps.cache-doctor.outputs.json }}\n${{ steps.cache-doctor.outputs.markdown }}\n',
      },
    })
  })

  it('requires a workbook path when printing a GitHub Actions workflow', () => {
    let stderr = ''

    const exitCode = runXlsxFormulaRecalcCli(['--print-github-action'], {
      commandName: 'xlsx-cache-doctor',
      stderr: (text) => {
        stderr += text
      },
    })

    expect(exitCode).toBe(1)
    expect(stderr).toContain('Expected workbook path or glob after --print-github-action')
  })

  it('allows generated GitHub Actions workflows to scan all matching workbooks', () => {
    let stdout = ''

    const exitCode = runXlsxFormulaRecalcCli(['--print-github-action', '**/*.xlsx', '--changed-files-only', 'false'], {
      commandName: 'xlsx-cache-doctor',
      stdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('workbooks: "**/*.xlsx"')
    expect(stdout).toContain('changed-files-only: "false"')
    expect(stdout).toContain('inspect-limit: "2000"')
    expect(stdout).toContain(`package-version: "${packageVersion}"`)
    expect(stdout).toContain('json-output: "${{ runner.temp }}/xlsx-cache-doctor.json"')
    expect(stdout).toContain('markdown-output: "${{ runner.temp }}/xlsx-cache-doctor.md"')
    expect(stdout).toContain('fail-on-stale: "false"')
  })

  it('keeps xlsx-cache-doctor in recalculation mode when readback output is explicit', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-recalc-cli-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      const outputPath = join(tempDir, 'stale-cache.fixed.xlsx')
      writeFileSync(inputPath, buildStaleFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--read', 'Sheet1!B2', '--out', outputPath, '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      const summary = readCliSummary(stdout)
      expect(summary.commandSucceeded).toBe(true)
      expect(summary.recalculationCompleted).toBe(true)
      expect(summary.reads['Sheet1!B2']?.value).toBe(20)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('keeps default file inspection bounded so large JSON reports do not collect every formula', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-default-bounded-'))
    try {
      const inputPath = join(tempDir, 'many-formulas.xlsx')
      writeFileSync(inputPath, buildManyFormulaCacheWorkbook({ formulaCount: 2005 }))
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.formulaCellCount).toBe(2005)
      expect(summary.inspectedFormulaCellCount).toBe(2000)
      expect(summary.uninspectedFormulaCellCount).toBe(5)
      expect(summary.inspectionLimit).toBe(2000)
      expect(summary.suggestedReads).not.toContain('Sheet1!B2006')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('checks every formula when a caller explicitly requests all formulas', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-all-formulas-'))
    try {
      const inputPath = join(tempDir, 'many-formulas.xlsx')
      writeFileSync(inputPath, buildManyFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--inspect-limit', 'all', '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.formulaCellCount).toBe(60)
      expect(summary.inspectedFormulaCellCount).toBe(60)
      expect(summary.uninspectedFormulaCellCount).toBe(0)
      expect(summary.inspectionLimit).toBe('all')
      expect(summary.staleCachedFormulaCount).toBe(1)
      expect(summary.cacheStatusSummary.stale).toBe(1)
      expect(summary.suggestedReads).toContain('Sheet1!B61')
      expect(summary.formulas.find((formula) => formula.target === 'Sheet1!B61')).toMatchObject({
        formula: '=A61*10',
        cachedValue: 999,
        literalRecalculatedValue: 600,
        cacheStatus: 'stale',
        staleCachedValue: true,
      })
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('reports uninspected formulas when a caller sets an explicit inspection limit', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-cache-doctor-limited-formulas-'))
    try {
      const inputPath = join(tempDir, 'many-formulas.xlsx')
      writeFileSync(inputPath, buildManyFormulaCacheWorkbook())
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync([inputPath, '--inspect-limit', '50', '--json'], {
        commandName: 'xlsx-cache-doctor',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(exitCode).toBe(0)
      const summary = readCliInspectionSummary(stdout)
      expect(summary.formulaCellCount).toBe(60)
      expect(summary.inspectedFormulaCellCount).toBe(50)
      expect(summary.uninspectedFormulaCellCount).toBe(10)
      expect(summary.inspectionLimit).toBe(50)
      expect(summary.staleCachedFormulaCount).toBe(0)
      expect(summary.cacheStatusSummary).toMatchObject({ inspected: 50, stale: 0 })
      expect(summary.suggestedReads).not.toContain('Sheet1!B61')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('hydrates external-link caches from companion workbook paths through streaming-native', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-external-'))
    try {
      const inputPath = join(tempDir, 'model.xlsx')
      const companionPath = join(tempDir, 'uploaded-rates.xlsx')
      const outputPath = join(tempDir, 'model.recalculated.xlsx')
      writeFileSync(inputPath, buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx', { lookupFormulas: false }))
      writeFileSync(companionPath, buildExternalSourceWorkbook([20, 30, 40]))
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync(
        [
          inputPath,
          '--external-workbook-target',
          companionPath,
          'file:///tmp/rates.xlsx',
          '--read',
          'Model!C1',
          '--out',
          outputPath,
          '--json',
        ],
        {
          stdout: (text) => {
            stdout += text
          },
        },
      )

      expect(exitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      const summary = readCliSummary(stdout)
      expect(summary.externalWorkbooks).toBe(1)
      expect(summary.reads['Model!C1']?.value).toBe(180)
      expect(summary.diagnostics?.externalWorkbookHydration).toMatchObject({
        externalWorkbookCount: 1,
        refreshedBookIndices: [1],
        refreshedCellCount: 3,
        references: [
          expect.objectContaining({
            status: 'refreshed',
            matchKind: 'exact-target',
            matchedFileName: 'uploaded-rates.xlsx',
            matchedTarget: 'file:///tmp/rates.xlsx',
          }),
        ],
      })
      expect(readExternalLinkCacheCellValue(readFileBytes(outputPath), 'B2')).toBe('20')
      expect(readExternalLinkCacheCellValue(readFileBytes(outputPath), 'B4')).toBe('40')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('rejects large formula-recalc external workbook CLI inputs', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-external-limit-'))
    try {
      const externalPath = join(tempDir, 'large-external.xlsx')
      writeFileSync(externalPath, Buffer.alloc(1_000_001))
      let stderr = ''

      const exitCode = runXlsxFormulaRecalcCli(['--demo', '--external-workbook', externalPath, '--json'], {
        stderr: (text) => {
          stderr += text
        },
      })

      expect(exitCode).toBe(1)
      expect(stderr).toContain('external workbook byte input is small-workbook only')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('preserves cached external-link values when CLI companion workbook paths are ambiguous through streaming-native', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-cli-ambiguous-'))
    try {
      const inputPath = join(tempDir, 'model.xlsx')
      const firstCompanionPath = join(tempDir, 'one', 'rates.xlsx')
      const secondCompanionPath = join(tempDir, 'two', 'rates.xlsx')
      const outputPath = join(tempDir, 'model.recalculated.xlsx')
      mkdirSync(join(tempDir, 'one'))
      mkdirSync(join(tempDir, 'two'))
      writeFileSync(inputPath, buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx', { lookupFormulas: false }))
      writeFileSync(firstCompanionPath, buildExternalSourceWorkbook([20, 30, 40]))
      writeFileSync(secondCompanionPath, buildExternalSourceWorkbook([200, 300, 400]))
      let stdout = ''

      const exitCode = await runXlsxFormulaRecalcCliAsync(
        [
          inputPath,
          '--external-workbook',
          firstCompanionPath,
          '--external-workbook',
          secondCompanionPath,
          '--read',
          'Model!C1',
          '--out',
          outputPath,
          '--json',
        ],
        {
          stdout: (text) => {
            stdout += text
          },
        },
      )

      expect(exitCode).toBe(0)
      const summary = readCliSummary(stdout)
      expect(summary.externalWorkbooks).toBe(2)
      expect(summary.reads['Model!C1']?.value).toBe(120)
      expect(summary.warnings).toContain(
        'Some supplied external workbook companions matched ambiguously; existing external-link cache values were preserved.',
      )
      expect(summary.diagnostics?.externalWorkbookHydration).toMatchObject({
        skippedAmbiguousMatchCount: 1,
        references: [
          expect.objectContaining({
            status: 'skipped-ambiguous-match',
            candidateCount: 2,
          }),
        ],
      })
      expect(readExternalLinkCacheCellValue(readFileBytes(outputPath), 'B2')).toBe('10')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})
