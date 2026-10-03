import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { exportXlsxSourceLiteralPatches, readXlsxTargetCell, writeSimpleXlsxWorkbook } from '@bilig/xlsx'
import { recalculateXlsx } from '@bilig/xlsx-formula-recalc'

const exampleDir = dirname(fileURLToPath(import.meta.url))
const outputDir = join(exampleDir, 'dist')
const sourcePath = join(outputDir, 'stackoverflow-63085785-source.xlsx')
const outputPath = join(outputDir, 'stackoverflow-63085785-recalculated.xlsx')

mkdirSync(outputDir, { recursive: true })

const sourceBytes = writeSimpleXlsxWorkbook({
  sheets: [
    {
      name: 'Sheet1',
      cells: [
        { address: 'A1', row: 0, col: 0, value: 1 },
        { address: 'B1', row: 0, col: 1, value: 2 },
        { address: 'C1', row: 0, col: 2, formula: 'A1+B1', value: 3 },
      ],
    },
  ],
})
writeFileSync(sourcePath, sourceBytes)

const staleValueBeforeRecalc = readNativeNumberCell(readXlsxTargetCell(sourceBytes, 'Sheet1', 'C1'), 'Sheet1!C1')
const editedBytes = exportXlsxSourceLiteralPatches({
  source: sourceBytes,
  patches: [{ sheetName: 'Sheet1', address: 'A1', value: 3 }],
})
const recalculated = recalculateXlsx(editedBytes, {
  fileName: 'stackoverflow-63085785.xlsx',
  reads: ['Sheet1!C1'],
})
const recalculatedValue = readNumberCell(recalculated.reads['Sheet1!C1'], 'Sheet1!C1')

writeFileSync(outputPath, Buffer.from(recalculated.xlsx))

const proof = {
  question: 'https://stackoverflow.com/questions/63085785/how-to-recalculate-all-formulas-in-excel-file-through-javascript',
  existingLibrary: '@bilig/xlsx source-preserving literal patch path',
  formula: 'Sheet1!C1 = A1 + B1',
  edit: 'Sheet1!A1: 1 -> 3',
  staleValueBeforeRecalc,
  recalculatedValue,
  sourceXlsx: sourcePath,
  outputXlsx: outputPath,
  verified: staleValueBeforeRecalc === 3 && recalculatedValue === 5,
}

if (!proof.verified) {
  throw new Error(`Stack Overflow native XLSX proof failed: ${JSON.stringify(proof, null, 2)}`)
}

console.log(JSON.stringify(proof, null, 2))

function readNumberCell(cell, target) {
  if (cell && typeof cell === 'object' && typeof cell.value === 'number') {
    return cell.value
  }
  throw new Error(`Expected numeric readback at ${target}, got ${JSON.stringify(cell)}`)
}

function readNativeNumberCell(cell, target) {
  if (cell && typeof cell.value === 'number') {
    return cell.value
  }
  throw new Error(`Expected native numeric readback at ${target}, got ${JSON.stringify(cell)}`)
}
