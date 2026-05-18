import type { WorkbookSnapshot } from '@bilig/protocol'
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

  it('prints a paste-ready report with formula samples and readback', () => {
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
    expect(stdout).toContain('- Package: `@bilig/headless@0.0.0-test`')
    expect(stdout).toContain('- Status: imported')
    expect(stdout).toContain('- shared formula expanded')
    expect(stdout).toContain('- `Summary!B2`: `A2*3`')
    expect(stdout).toContain('- `Summary!B2`: value `21`, formula `=A2*3`')
    expect(stdout).toContain('- [ ] This reduced case is public')
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
