import { WorkPaper } from '@bilig/headless'

const csvInput = [
  ['Product', 'Q1', 'Q2', 'Q3', 'Q4'],
  ['Widget A', 100, 150, 200, 250],
  ['Widget B', 80, 90, 100, 110],
  ['Widget C', 300, 310, 320, 330],
]

const workbook = WorkPaper.buildFromSheets({
  Data: csvInput,
})

const dataSheet = requireSheet(workbook, 'Data')

// Add one formula-backed summary cell
workbook.setCellContents({ sheet: dataSheet, row: 4, col: 0 }, 'Total Q1')
workbook.setCellContents({ sheet: dataSheet, row: 4, col: 1 }, '=SUM(B2:B4)')

// Read the display/result back
const totalQ1Cell = workbook.getCellValue({ sheet: dataSheet, row: 4, col: 1 })

const output = {
  success: true,
  totalQ1: totalQ1Cell.value,
}

assertSummary(output)
console.log(JSON.stringify(output, null, 2))

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function assertSummary(summary) {
  const expected = {
    success: true,
    totalQ1: 480,
  }

  if (JSON.stringify(summary) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected WorkPaper result: ${JSON.stringify(summary)}`)
  }
}
