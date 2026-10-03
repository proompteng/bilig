import type { WorkbookSnapshot } from '@bilig/protocol'
import type { WorkbookCompatibilityReport } from '@bilig/xlsx/workbook-compatibility-report'
import { describe, expect, it } from 'vitest'
import {
  formulaClinicHelpText,
  parseFormulaClinicCliArgs,
  runFormulaClinicCli,
  type FormulaClinicImportXlsx,
} from '../formula-clinic-cli.js'

describe('formula clinic CLI', () => {
  it('parses the workbook path, requested cells, and numeric options', () => {
    expect(
      parseFormulaClinicCliArgs([
        'reduced.xlsx',
        '--cells',
        "Summary!B2,'Input Sheet'!C3",
        '--formula-samples',
        '3',
        '--timeout-ms',
        '5000',
      ]),
    ).toEqual({
      cells: [
        { sheetName: 'Summary', a1: 'B2' },
        { sheetName: 'Input Sheet', a1: 'C3' },
      ],
      evaluationTimeoutMs: 5000,
      filePath: 'reduced.xlsx',
      help: false,
      maxFormulaSamples: 3,
    })
  })

  it('prints a Markdown report with formula samples and readback', () => {
    let stdout = ''
    const importXlsx: FormulaClinicImportXlsx = () => ({
      snapshot: clinicWorkbookSnapshot(),
      sheetNames: ['Summary'],
      warnings: ['shared formula expanded'],
    })

    const exitCode = runFormulaClinicCli({
      argv: ['reduced.xlsx', '--cells', 'Summary!B2'],
      importXlsx,
      packageVersion: '0.0.0-test',
      readFile: () => new Uint8Array([1, 2, 3]),
      statFileSizeBytes: () => 3,
      writeStdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toContain('# Bilig formula clinic report')
    expect(stdout).toContain('- Package: `@bilig/workpaper@0.0.0-test`')
    expect(stdout).toContain('- Status: imported')
    expect(stdout).toContain('- shared formula expanded')
    expect(stdout).toContain('- `Summary!B2`: `A2*3`')
    expect(stdout).toContain('- `Summary!B2`: value `21`, formula `=A2*3`')
    expect(stdout).toContain('- [ ] This reduced case is public')
  })

  it('uses native file-backed preflight for large files before reading XLSX bytes', () => {
    let stdout = ''
    let readFileCalled = false
    let importCalled = false
    let reportPath = ''

    const exitCode = runFormulaClinicCli({
      argv: ['large.xlsx', '--cells', 'Summary!B2', '--formula-samples', '4'],
      buildWorkbookCompatibilityReportFromFile: (inputPath, options) => {
        reportPath = inputPath
        expect(options).toEqual({ fileName: 'large.xlsx', inspectLimit: 4 })
        return largeNativeCompatibilityReport()
      },
      importXlsx: () => {
        importCalled = true
        throw new Error('should not import large XLSX through WorkPaper')
      },
      packageVersion: '0.0.0-test',
      readFile: () => {
        readFileCalled = true
        throw new Error('should not read large XLSX bytes')
      },
      statFileSizeBytes: () => 1_000_001,
      writeStdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(reportPath).toBe('large.xlsx')
    expect(readFileCalled).toBe(false)
    expect(importCalled).toBe(false)
    expect(stdout).toContain('- Status: native-preflight')
    expect(stdout).toContain('- Sheets: `Summary`')
    expect(stdout).toContain('- Formula cells: 1200')
    expect(stdout).toContain('- Engine mode: `streaming-native`')
    expect(stdout).toContain('- Fallback used: `false`')
    expect(stdout).toContain('- Max observed RSS bytes: 123456')
    expect(stdout).toContain('- Phase RSS peaks: 1')
    expect(stdout).toContain('- Risk: `medium`')
    expect(stdout).toContain('- Inspected formula cells: 4')
    expect(stdout).toContain('- Scanned formula cells: 1200')
    expect(stdout).toContain('- Targeted formula cells: 4')
    expect(stdout).toContain('- Patched formula caches: 0')
    expect(stdout).toContain('- Unsupported function references: 2')
    expect(stdout).toContain('- Unsupported reason: unsupported functions: CUBESET (2)')
    expect(stdout).toContain('skipped WorkPaper readback because this workbook is above the small-workbook clinic limit')
  })

  it('returns a failed report when import throws', () => {
    let stdout = ''
    const exitCode = runFormulaClinicCli({
      argv: ['broken.xlsx'],
      importXlsx: () => {
        throw new Error('Invalid workbook')
      },
      packageVersion: '0.0.0-test',
      readFile: () => new Uint8Array([1]),
      statFileSizeBytes: () => 1,
      writeStdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(1)
    expect(stdout).toContain('- Status: failed')
    expect(stdout).toContain('- Error: Invalid workbook')
  })

  it('prints help without requiring an importer', () => {
    let stdout = ''
    const exitCode = runFormulaClinicCli({
      argv: ['--help'],
      importXlsx: () => {
        throw new Error('should not import')
      },
      writeStdout: (text) => {
        stdout += text
      },
    })

    expect(exitCode).toBe(0)
    expect(stdout).toBe(formulaClinicHelpText())
  })
})

function clinicWorkbookSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: 'Clinic' },
    sheets: [
      {
        id: 1,
        name: 'Summary',
        order: 0,
        cells: [
          { address: 'A2', value: 7 },
          { address: 'B2', formula: 'A2*3' },
        ],
      },
    ],
  }
}

function largeNativeCompatibilityReport(): WorkbookCompatibilityReport {
  return {
    schemaVersion: 'bilig-workbook-compatibility-report.v1',
    verified: true,
    input: {
      fileName: 'large.xlsx',
      externalWorkbookCount: 0,
      inspectLimit: 4,
    },
    workbook: {
      sheetCount: 1,
      sheetNames: ['Summary'],
      nonEmptyCellCount: 2400,
      formulaCellCount: 1200,
      definedNameCount: 0,
      tableCount: 0,
      pivotTableCount: 0,
      chartCount: 0,
      macroModuleCount: 0,
    },
    findings: {
      unsupportedFunctions: [{ name: 'CUBESET', count: 2 }],
      externalLinks: {
        count: 3,
        unresolvedCount: 0,
        refreshedCount: 0,
      },
      macroModules: {
        count: 0,
        byteLength: 0,
      },
      volatileFunctions: [],
      pivotTables: {
        count: 0,
        unsupportedCount: 0,
        cacheOnlyCount: 0,
      },
      staleCachedFormulas: {
        count: 0,
      },
      missingCachedFormulaValues: {
        count: 0,
      },
      unsupportedRecalculations: {
        count: 2,
      },
      warnings: ['native preflight warning'],
    },
    risk: {
      level: 'medium',
      reasons: ['unsupported function references'],
    },
    cacheInspection: {
      inspectedFormulaCellCount: 4,
      uninspectedFormulaCellCount: 1196,
      inspectionLimit: 4,
      suggestedReads: ['Summary!B2'],
    },
    diagnostics: {
      engineMode: 'streaming-native',
      fallbackUsed: false,
      inputBytes: 1_000_001,
      phaseRssPeaks: [{ phase: 'file-api:formula-cache', rssBytes: 123456 }],
      maxObservedRssBytes: 123456,
      sheetCount: 1,
      targetRowCount: 4,
      editCount: 0,
      readCount: 4,
      formulaCounts: {
        scannedFormulaCellCount: 1200,
        targetedFormulaCellCount: 4,
        evaluatedFormulaCellCount: 0,
        patchedFormulaCacheCount: 0,
        unsupportedFormulaCellCount: 2,
        nativeKernelFormulaCellCount: 0,
        nativeKernelBatchCount: 0,
      },
      patchedCacheCount: 0,
      unsupportedReason: 'unsupported functions: CUBESET (2)',
    },
    commandSucceeded: true,
    inspectionCompleted: true,
    recalculationCompleted: false,
    excelParity: 'not_proven',
    limitations: [],
    next: {
      docs: 'https://example.test/docs',
      command: 'workbook-compatibility-report large.xlsx --json',
    },
  }
}
