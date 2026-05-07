import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workspace = mkdtempSync(join(tmpdir(), 'bilig-workpaper-persistence-'))
const savePath = join(workspace, 'workpaper.json')

try {
  const workbook = WorkPaper.buildFromSheets(
    {
      Plan: [
        ['Month', 'Bookings', 'Churn', 'Net MRR'],
        ['January', 12000, 800, '=B2-C2'],
        ['February', 15000, 900, '=B3-C3'],
        ['March', 18000, 1200, '=B4-C4'],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Quarter net MRR', '=SUM(Plan!D2:D4)'],
        ['Annualized run rate', '=B2*12'],
        ['Expansion-adjusted ARR', null],
      ],
    },
    { maxRows: 1000, maxColumns: 64, useColumnIndex: true },
  )

  workbook.addNamedExpression('ExpansionRatePercent', 8)
  const summarySheet = requireSheet(workbook, 'Summary')
  workbook.setCellContents({ sheet: summarySheet, row: 3, col: 1 }, '=B3*(100+ExpansionRatePercent)/100')

  const beforeSave = readSummary(workbook)
  writeFileSync(savePath, serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true })))

  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync(savePath, 'utf8')))
  const planSheet = requireSheet(restored, 'Plan')
  restored.setCellContents({ sheet: planSheet, row: 3, col: 1 }, 21000)

  const afterRestoreAndEdit = readSummary(restored)
  const persistedAgain = exportWorkPaperDocument(restored, { includeConfig: true })
  const output = {
    beforeSave,
    afterRestoreAndEdit,
    persistedSheets: persistedAgain.sheets.map((sheet) => sheet.name),
    persistedNamedExpressions: restored.listNamedExpressions(),
    saveFileBytes: readFileSync(savePath).byteLength,
  }

  assertOutput(output)
  console.log(JSON.stringify(output, null, 2))
} finally {
  rmSync(workspace, { recursive: true, force: true })
}

function readSummary(workpaper) {
  const summarySheet = requireSheet(workpaper, 'Summary')
  return {
    quarterNetMrr: readNumber(workpaper, summarySheet, 1, 1, 'quarter net MRR'),
    annualizedRunRate: readNumber(workpaper, summarySheet, 2, 1, 'annualized run rate'),
    expansionAdjustedArr: readNumber(workpaper, summarySheet, 3, 1, 'expansion adjusted ARR'),
  }
}

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be a number, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function assertOutput(output) {
  const expected = {
    beforeSave: {
      quarterNetMrr: 42100,
      annualizedRunRate: 505200,
      expansionAdjustedArr: 545616,
    },
    afterRestoreAndEdit: {
      quarterNetMrr: 45100,
      annualizedRunRate: 541200,
      expansionAdjustedArr: 584496,
    },
    persistedSheets: ['Plan', 'Summary'],
    persistedNamedExpressions: ['ExpansionRatePercent'],
  }

  const comparable = {
    beforeSave: output.beforeSave,
    afterRestoreAndEdit: output.afterRestoreAndEdit,
    persistedSheets: output.persistedSheets,
    persistedNamedExpressions: output.persistedNamedExpressions,
  }

  if (JSON.stringify(comparable) !== JSON.stringify(expected) || output.saveFileBytes <= 0) {
    throw new Error(`Unexpected persistence result: ${JSON.stringify(output)}`)
  }
}
