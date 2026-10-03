import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelPackageOpenSaveOracle } from '@bilig/excel-fixtures'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { WorkPaper, exportXlsx, importXlsx, recalculateXlsx } from '@bilig/workpaper/xlsx'

const inspectedFormulaCells = ['B1', 'C1', 'D1', 'E1', 'F1', 'G1'] as const

describe('macOS Desktop Excel unsupported formula cache oracle', () => {
  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'pins native save and full-rebuild semantics for unsupported cached formulas',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = mkdtempSync(join(tmpdir(), 'bilig-xlsx-unsupported-cache-oracle-'))
      try {
        const sourceBytes = buildUnsupportedFormulaCacheXlsx()
        expect(readFormulaState(sourceBytes, 'C1')).toMatchObject({ cachedValue: '0' })
        expect(readFormulaState(sourceBytes, 'E1')).toMatchObject({ cachedValue: 'stale' })
        expect(readFormulaState(sourceBytes, 'F1')).toMatchObject({ cachedValue: '99' })
        expect(readFormulaState(sourceBytes, 'G1')).toMatchObject({ cachedValue: 'stale' })

        const nativePath = join(tempDir, 'unsupported-cache-native.xlsx')
        writeFileSync(nativePath, sourceBytes)
        const excelNativeSave = runMacosExcelPackageOpenSaveOracle({
          workbookPath: nativePath,
          calculationPolicy: 'none',
          saveWorkbook: true,
          timeoutMs: 120_000,
          updateLinks: 'never',
        })
        expect(excelNativeSave.excelVersion).toMatch(/^\d+\./u)

        const excelNativeBytes = new Uint8Array(readFileSync(nativePath))
        const headlessNativeBytes = exportXlsx(importXlsx(sourceBytes, 'unsupported-cache-native.xlsx').snapshot)
        expect(readFormulaCacheStates(headlessNativeBytes)).toEqual(readFormulaCacheStates(excelNativeBytes))

        const fullRebuildPath = join(tempDir, 'unsupported-cache-full-rebuild.xlsx')
        writeFileSync(fullRebuildPath, sourceBytes)
        const excelFullRebuild = runMacosExcelInspectionOracle({
          workbookPath: fullRebuildPath,
          worksheetName: 'Model',
          formulaCells: [],
          inspectCells: [...inspectedFormulaCells],
          saveWorkbook: true,
          timeoutMs: 120_000,
          updateLinks: 'never',
        })
        expect(excelFullRebuild.excelVersion).toMatch(/^\d+\./u)
        expect(excelFullRebuild.cells).toEqual([
          expect.objectContaining({ address: 'B1', value: { kind: 'error', value: String(ErrorCode.Name) } }),
          expect.objectContaining({ address: 'C1', value: { kind: 'error', value: String(ErrorCode.Name) } }),
          expect.objectContaining({ address: 'D1', value: { kind: 'error', value: String(ErrorCode.Field) } }),
          expect.objectContaining({ address: 'E1', value: { kind: 'error', value: String(ErrorCode.Field) } }),
          expect.objectContaining({ address: 'F1', value: { kind: 'error', value: String(ErrorCode.Name) } }),
          expect.objectContaining({ address: 'G1', value: { kind: 'string', value: 'unknown' } }),
        ])

        const recalculated = recalculateXlsx(sourceBytes, {
          fileName: 'unsupported-cache-full-rebuild.xlsx',
          reads: inspectedFormulaCells.map((address) => `Model!${address}`),
        })
        expect(errorCell(recalculated.reads['Model!B1'])).toBe(ErrorCode.Name)
        expect(errorCell(recalculated.reads['Model!C1'])).toBe(ErrorCode.Name)
        expect(errorCell(recalculated.reads['Model!D1'])).toBe(ErrorCode.Field)
        expect(errorCell(recalculated.reads['Model!E1'])).toBe(ErrorCode.Field)
        expect(errorCell(recalculated.reads['Model!F1'])).toBe(ErrorCode.Name)
        expect(stringCell(recalculated.reads['Model!G1'])).toBe('unknown')

        const excelFullRebuildBytes = new Uint8Array(readFileSync(fullRebuildPath))
        expect(readFormulaValueStates(recalculated.xlsx)).toEqual(readFormulaValueStates(excelFullRebuildBytes))
      } finally {
        rmSync(tempDir, { recursive: true, force: true })
      }
    },
    180_000,
  )
})

function buildUnsupportedFormulaCacheXlsx(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Model: [['AAPL', 0, 0, '', '', 0, '']],
  })
  try {
    let bytes = exportXlsx(workbook.exportSnapshot())
    bytes = setWorkbookCalcPr(bytes, '<calcPr calcMode="manual" calcOnSave="1" calcCompleted="0" fullCalcOnLoad="0"/>')
    const cells: Readonly<Record<string, string>> = {
      B1: '<c r="B1"><f>_xldudf_WISEPRICE(A1,&quot;Shares Outstanding&quot;)</f><v>14935800000</v></c>',
      C1: '<c r="C1"><f>B1/1000000</f><v>0</v></c>',
      D1: '<c r="D1" t="str"><f>_FV(A1,&quot;Ticker symbol&quot;,TRUE)</f><v>AAPL</v></c>',
      E1: '<c r="E1" t="str"><f>D1&amp;&quot; ok&quot;</f><v>stale</v></c>',
      F1: '<c r="F1"><f>UNKNOWNFUNC(42)</f><v>99</v></c>',
      G1: '<c r="G1" t="str"><f>IFERROR(F1,&quot;unknown&quot;)</f><v>stale</v></c>',
    }
    for (const [address, cellXml] of Object.entries(cells)) {
      bytes = replaceCellXml(bytes, 'xl/worksheets/sheet1.xml', address, cellXml)
    }
    return bytes
  } finally {
    workbook.dispose()
  }
}

function stringCell(value: unknown): string {
  if (typeof value === 'object' && value !== null && 'tag' in value && value.tag === ValueTag.String && 'value' in value) {
    return String(value.value)
  }
  throw new Error(`Expected string cell value, received ${JSON.stringify(value)}`)
}

function errorCell(value: unknown): ErrorCode {
  if (typeof value === 'object' && value !== null && 'tag' in value && value.tag === ValueTag.Error && 'code' in value) {
    const code = value.code
    if (typeof code === 'number') {
      return code
    }
  }
  throw new Error(`Expected error cell value, received ${JSON.stringify(value)}`)
}

function replaceCellXml(bytes: Uint8Array, sheetPath: string, address: string, replacement: string): Uint8Array {
  const zip = unzipSync(bytes)
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const pattern = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<\\/c>`, 'u')
  if (!pattern.test(sheetXml)) {
    throw new Error(`Missing cell XML for ${address}`)
  }
  zip[sheetPath] = strToU8(sheetXml.replace(pattern, replacement))
  return zipSync(zip)
}

function setWorkbookCalcPr(bytes: Uint8Array, calcPrXml: string): Uint8Array {
  const zip = unzipSync(bytes)
  const workbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  if (/<calcPr\b[\s\S]*?\/>/u.test(workbookXml)) {
    zip['xl/workbook.xml'] = strToU8(workbookXml.replace(/<calcPr\b[\s\S]*?\/>/u, calcPrXml))
    return zipSync(zip)
  }
  zip['xl/workbook.xml'] = strToU8(workbookXml.replace('</workbook>', `${calcPrXml}</workbook>`))
  return zipSync(zip)
}

function readFormulaCacheStates(bytes: Uint8Array): Readonly<Record<string, FormulaCacheState>> {
  return Object.fromEntries(inspectedFormulaCells.map((address) => [address, readFormulaState(bytes, address)]))
}

function readFormulaValueStates(bytes: Uint8Array): Readonly<Record<string, Omit<FormulaCacheState, 'formula'>>> {
  return Object.fromEntries(
    inspectedFormulaCells.map((address) => {
      const state = readFormulaState(bytes, address)
      return [address, { cachedValue: state.cachedValue, cellType: state.cellType }]
    }),
  )
}

interface FormulaCacheState {
  readonly formula: string | null
  readonly cachedValue: string | null
  readonly cellType: string | null
}

function readFormulaState(bytes: Uint8Array, address: string): FormulaCacheState {
  const sheetXml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
  const cellXml = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<\\/c>`, 'u').exec(sheetXml)?.[0]
  if (!cellXml) {
    throw new Error(`Missing cell XML for ${address}`)
  }
  return {
    formula: /<f[^>]*>([\s\S]*?)<\/f>/u.exec(cellXml)?.[1] ?? null,
    cachedValue: /<v>([\s\S]*?)<\/v>/u.exec(cellXml)?.[1] ?? null,
    cellType: /\bt="([^"]*)"/u.exec(cellXml)?.[1] ?? null,
  }
}
