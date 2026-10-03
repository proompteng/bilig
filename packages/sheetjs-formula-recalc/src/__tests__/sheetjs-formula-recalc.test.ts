import { existsSync, mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { recalculateSheetjsWorkbook } from '../index.js'
import { WorkPaper, exportXlsx } from '@bilig/workpaper/xlsx'
import { runXlsxFormulaRecalcCli } from '@bilig/xlsx-formula-recalc/cli-api'

describe('@bilig/sheetjs-formula-recalc', () => {
  it('re-exports the native XLSX recalculation boundary with a SheetJS-named API', async () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })

    const sourceBytes = exportXlsx(sourceWorkbook.exportSnapshot())
    sourceWorkbook.dispose()

    const result = await recalculateSheetjsWorkbook(sourceBytes, {
      fileName: 'sheetjs-pricing.xlsx',
      edits: [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ],
      reads: ['Summary!B2'],
    })

    expect(readNumber(result.reads['Summary!B2'])).toBe(72_000)
    expect(result.warnings).toEqual([])
  })

  it('supports a sheetjs-recalc command name for help and demo proof output', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'sheetjs-formula-recalc-cli-'))
    try {
      let help = ''
      const helpExitCode = runXlsxFormulaRecalcCli(['--help'], {
        commandName: 'sheetjs-recalc',
        stdout: (text) => {
          help += text
        },
      })
      expect(helpExitCode).toBe(0)
      expect(help).toContain('Usage: sheetjs-recalc')

      let stdout = ''
      const outputPath = join(tempDir, 'sheetjs-demo.recalculated.xlsx')
      const demoExitCode = runXlsxFormulaRecalcCli(['--demo', '--out', outputPath, '--json'], {
        commandName: 'sheetjs-recalc',
        stdout: (text) => {
          stdout += text
        },
      })

      expect(demoExitCode).toBe(0)
      expect(existsSync(outputPath)).toBe(true)
      expect(readNumber(JSON.parse(stdout).reads['Summary!B2'])).toBe(72_000)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})

function readNumber(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'number') {
    return value.value
  }
  throw new Error(`Expected numeric cell value, received ${JSON.stringify(value)}`)
}
