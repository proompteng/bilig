import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { strFromU8, unzipSync } from 'fflate'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
import { isMacosExcelInstalled, runMacosExcelInspectionOracle, runMacosExcelStructuralOperationOracle } from '@bilig/excel-fixtures'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'
import { createExcelAccessibleTempDir, removeMacosExcelTestDir } from './macos-excel-oracle-test-utils.js'
import { buildWorkbookWithWorksheetControl } from './worksheet-control-artifact-fixture.js'

describe('macOS Desktop Excel worksheet control artifacts oracle', () => {
  it('preserves imported worksheet form controls through WorkPaper export after a structural row insert', () => {
    const source = buildWorkbookWithWorksheetControl()
    const workpaper = WorkPaper.buildFromSnapshot(importXlsx(source, 'worksheet-control-source.xlsx').snapshot)
    try {
      const sheet = workpaper.getSheetId('Model')
      if (sheet === undefined) {
        throw new Error('Expected Model sheet to be available')
      }
      workpaper.addRows(sheet, 0, 1)

      const exported = exportXlsx(workpaper.exportSnapshot())
      const reimported = importXlsx(exported, 'worksheet-control-headless-roundtrip.xlsx')

      expect(controlCountMetrics(exported)).toEqual(controlCountMetrics(source))
      expect(controlStructuralMetrics(exported)).toMatchObject({
        controlAnchorRows: [4, 5],
        vmlAnchors: ['6, 11, 4, 5, 11, 7, 5, 12'],
      })
      expect(reimported.snapshot.workbook.metadata?.controlArtifacts?.parts.map((part) => part.path).toSorted()).toEqual([
        'xl/ctrlProps/ctrlProp1.xml',
        'xl/drawings/vmlDrawing1.vml',
      ])
      expect(reimported.snapshot.sheets[0]?.metadata?.controlArtifacts?.controlsXml).toContain('macro="[0]!WriteHelloWorld"')
    } finally {
      workpaper.dispose()
    }
  })

  it.runIf(process.env.BILIG_EXCEL_ORACLE_RUN === '1')(
    'matches Desktop Excel worksheet form control anchors after structural row inserts',
    () => {
      if (!isMacosExcelInstalled()) {
        throw new Error('BILIG_EXCEL_ORACLE_RUN=1 requires /Applications/Microsoft Excel.app')
      }

      const tempDir = createExcelAccessibleTempDir('bilig-headless-excel-control-artifacts-oracle-')
      try {
        const source = buildWorkbookWithWorksheetControl()
        const excelWorkbookPath = join(tempDir, 'excel-control-artifacts-source.xlsx')
        writeFileSync(excelWorkbookPath, source)
        const excelResult = runMacosExcelStructuralOperationOracle({
          workbookPath: excelWorkbookPath,
          worksheetName: 'Model',
          operations: [{ kind: 'insertRows', range: '1:1' }],
          inspectCells: ['A1', 'A2'],
          saveWorkbook: true,
          timeoutMs: 90_000,
        })
        expect(excelResult.cells.map((cell) => cell.value)).toEqual([
          { kind: 'string', value: '' },
          { kind: 'string', value: 'Control fixture' },
        ])
        const excelTruth = new Uint8Array(readFileSync(excelWorkbookPath))
        expect(controlCountMetrics(excelTruth)).toEqual(controlCountMetrics(source))

        const workpaper = WorkPaper.buildFromSnapshot(importXlsx(source, 'headless-control-artifacts-source.xlsx').snapshot)
        try {
          const sheet = workpaper.getSheetId('Model')
          if (sheet === undefined) {
            throw new Error('Expected Model sheet to be available')
          }
          workpaper.addRows(sheet, 0, 1)

          const headlessPath = join(tempDir, 'headless-control-artifacts.xlsx')
          writeFileSync(headlessPath, exportXlsx(workpaper.exportSnapshot()))
          const headlessExcel = runMacosExcelInspectionOracle({
            workbookPath: headlessPath,
            worksheetName: 'Model',
            formulaCells: [],
            inspectCells: ['A1', 'A2'],
            saveWorkbook: true,
            timeoutMs: 90_000,
          })
          expect(headlessExcel.cells).toEqual(excelResult.cells)

          const headlessTruth = new Uint8Array(readFileSync(headlessPath))
          expect(controlStructuralMetrics(headlessTruth)).toEqual(controlStructuralMetrics(excelTruth))
          const reimported = importXlsx(headlessTruth, 'headless-control-artifacts-truth.xlsx')
          expect(reimported.snapshot.sheets[0]?.metadata?.controlArtifacts?.controlsXml).toContain('macro="[0]!WriteHelloWorld"')
        } finally {
          workpaper.dispose()
        }
      } finally {
        removeMacosExcelTestDir(tempDir)
      }
    },
    120_000,
  )
})

function controlCountMetrics(bytes: Uint8Array): {
  controls: number
  controlPropertyParts: number
  macroAssignments: string[]
  vmlDrawingParts: number
} {
  const zip = unzipSync(bytes)
  const controlXml = Object.entries(zip)
    .filter(([path]) => path.startsWith('xl/worksheets/') && path.endsWith('.xml'))
    .map(([, content]) => strFromU8(content))
    .join('')
  const controls = controlXml.match(/<control\b[^>]*(?:\/>|>[\s\S]*?<\/control>)/gu) ?? []
  return {
    controls: controls.length,
    controlPropertyParts: Object.keys(zip).filter((path) => path.startsWith('xl/ctrlProps/')).length,
    macroAssignments: controls.flatMap((control) => {
      const macro = /\bmacro=(["'])([\s\S]*?)\1/u.exec(control)?.[2]
      return macro ? [macro] : []
    }),
    vmlDrawingParts: Object.keys(zip).filter((path) => path.includes('vmlDrawing')).length,
  }
}

function controlStructuralMetrics(bytes: Uint8Array): {
  controls: number
  controlPropertyParts: number
  macroAssignments: string[]
  vmlDrawingParts: number
  controlAnchorRows: number[]
  vmlAnchors: string[]
} {
  const zip = unzipSync(bytes)
  const worksheetXml = Object.entries(zip)
    .filter(([path]) => path.startsWith('xl/worksheets/') && path.endsWith('.xml'))
    .map(([, content]) => strFromU8(content))
    .join('')
  const vmlXml = Object.entries(zip)
    .filter(([path]) => path.startsWith('xl/drawings/') && path.endsWith('.vml'))
    .map(([, content]) => strFromU8(content))
    .join('')
  const controlAnchorRows = [...worksheetXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?row>(\d+)<\/(?:[A-Za-z_][\w.-]*:)?row>/gu)]
    .map((match) => Number(match[1]))
    .filter((value) => Number.isInteger(value))
  const vmlAnchors = [...vmlXml.matchAll(/<x:Anchor>([\s\S]*?)<\/x:Anchor>/gu)].map((match) =>
    (match[1] ?? '').replace(/\s+/gu, ' ').trim(),
  )
  return {
    ...controlCountMetrics(bytes),
    controlAnchorRows,
    vmlAnchors,
  }
}
