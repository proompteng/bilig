import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import * as XLSX from 'xlsx'
import { recalculateXlsx } from 'xlsx-formula-recalc'

const exampleDir = dirname(fileURLToPath(import.meta.url))
const outputDir = join(exampleDir, 'dist')
const sourcePath = join(outputDir, 'stackoverflow-63085785-source.xlsx')
const outputPath = join(outputDir, 'stackoverflow-63085785-recalculated.xlsx')

mkdirSync(outputDir, { recursive: true })

const workbook = XLSX.utils.book_new()
const worksheet = {
  A1: { t: 'n', v: 1 },
  B1: { t: 'n', v: 2 },
  C1: { t: 'n', f: 'A1+B1', v: 3 },
  '!ref': 'A1:C1',
}
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1')

const sourceBytes = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
writeFileSync(sourcePath, sourceBytes)

const editedWorkbook = XLSX.read(sourceBytes, { type: 'buffer' })
editedWorkbook.Sheets.Sheet1.A1.v = 3

const staleValueBeforeRecalc = editedWorkbook.Sheets.Sheet1.C1.v
const editedBytes = XLSX.write(editedWorkbook, { bookType: 'xlsx', type: 'buffer' })
const recalculated = recalculateXlsx(editedBytes, {
  fileName: 'stackoverflow-63085785.xlsx',
  reads: ['Sheet1!C1'],
})
const recalculatedValue = readNumberCell(recalculated.reads['Sheet1!C1'], 'Sheet1!C1')

writeFileSync(outputPath, Buffer.from(recalculated.xlsx))

const proof = {
  question: 'https://stackoverflow.com/questions/63085785/how-to-recalculate-all-formulas-in-excel-file-through-javascript',
  existingLibrary: 'SheetJS / xlsx',
  formula: 'Sheet1!C1 = A1 + B1',
  edit: 'Sheet1!A1: 1 -> 3',
  staleValueBeforeRecalc,
  recalculatedValue,
  sourceXlsx: sourcePath,
  outputXlsx: outputPath,
  verified: staleValueBeforeRecalc === 3 && recalculatedValue === 5,
}

if (!proof.verified) {
  throw new Error(`Stack Overflow SheetJS proof failed: ${JSON.stringify(proof, null, 2)}`)
}

console.log(JSON.stringify(proof, null, 2))

function readNumberCell(cell, target) {
  if (cell && typeof cell === 'object' && typeof cell.value === 'number') {
    return cell.value
  }
  throw new Error(`Expected numeric readback at ${target}, got ${JSON.stringify(cell)}`)
}
