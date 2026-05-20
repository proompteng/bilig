import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import ExcelJS from 'exceljs'
import { recalculateExceljsWorkbook } from 'exceljs-formula-recalc'

const exampleDir = dirname(fileURLToPath(import.meta.url))
const outputDir = join(exampleDir, 'dist')
const sourcePath = join(outputDir, 'stackoverflow-44199441-source.xlsx')
const outputPath = join(outputDir, 'stackoverflow-44199441-recalculated.xlsx')

mkdirSync(outputDir, { recursive: true })

const sourceWorkbook = new ExcelJS.Workbook()
const sourceSheet = sourceWorkbook.addWorksheet('Sheet1')
sourceSheet.getCell('A1').value = 1
sourceSheet.getCell('B1').value = 2
sourceSheet.getCell('C1').value = {
  formula: 'A1+B1',
  result: 3,
}

const sourceBytes = Buffer.from(await sourceWorkbook.xlsx.writeBuffer())
writeFileSync(sourcePath, sourceBytes)

const workbook = new ExcelJS.Workbook()
await workbook.xlsx.load(sourceBytes)

const sheet = workbook.getWorksheet('Sheet1')
sheet.getCell('A1').value = 3

const staleFormulaCell = sheet.getCell('C1').value
const staleValueBeforeRecalc = readFormulaResult(staleFormulaCell, 'ExcelJS stale Sheet1!C1')

const recalculated = await recalculateExceljsWorkbook(workbook, {
  reads: ['Sheet1!C1'],
})
const recalculatedValue = readNumberCell(recalculated.reads['Sheet1!C1'], 'Sheet1!C1')
const patchedFormulaCell = workbook.getWorksheet('Sheet1').getCell('C1').value
const patchedExceljsResult = readFormulaResult(patchedFormulaCell, 'ExcelJS patched Sheet1!C1')

writeFileSync(outputPath, Buffer.from(await workbook.xlsx.writeBuffer()))

const proof = {
  question: 'https://stackoverflow.com/questions/44199441/get-computed-value-of-excel-sheet-cell-in-node-js',
  existingLibrary: 'ExcelJS',
  formula: 'Sheet1!C1 = A1 + B1',
  edit: 'Sheet1!A1: 1 -> 3',
  staleValueBeforeRecalc,
  recalculatedValue,
  patchedExceljsResult,
  sourceXlsx: sourcePath,
  outputXlsx: outputPath,
  verified: staleValueBeforeRecalc === 3 && recalculatedValue === 5 && patchedExceljsResult === 5,
}

if (!proof.verified) {
  throw new Error(`Stack Overflow ExcelJS proof failed: ${JSON.stringify(proof, null, 2)}`)
}

console.log(JSON.stringify(proof, null, 2))

function readNumberCell(cell, target) {
  if (cell && typeof cell === 'object' && typeof cell.value === 'number') {
    return cell.value
  }
  throw new Error(`Expected numeric readback at ${target}, got ${JSON.stringify(cell)}`)
}

function readFormulaResult(cell, target) {
  if (cell && typeof cell === 'object' && typeof cell.result === 'number') {
    return cell.result
  }
  throw new Error(`Expected ExcelJS formula result at ${target}, got ${JSON.stringify(cell)}`)
}
