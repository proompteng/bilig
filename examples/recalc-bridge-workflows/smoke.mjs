import { mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { exportXlsxSourceLiteralPatches, readXlsxTargetCell } from '@bilig/xlsx'
import ExcelJS from 'exceljs'
import { recalculateExceljsWorkbook } from '@bilig/exceljs-formula-recalc'
import { recalculateXlsx } from '@bilig/xlsx-formula-recalc'
import XlsxPopulate from 'xlsx-populate'

const exampleDir = dirname(fileURLToPath(import.meta.url))
const outputDir = join(exampleDir, 'dist')
const sourcePath = join(outputDir, 'bridge-source.xlsx')
const biligXlsxPath = join(outputDir, 'bridge-bilig-xlsx-recalculated.xlsx')
const xlsxPopulatePath = join(outputDir, 'bridge-xlsx-populate-recalculated.xlsx')
const exceljsPath = join(outputDir, 'bridge-exceljs-recalculated.xlsx')

mkdirSync(outputDir, { recursive: true })

const sourceXlsx = await createSourceXlsx()
writeFileSync(sourcePath, sourceXlsx)

const biligXlsx = runBiligXlsxBridge(sourceXlsx)
writeFileSync(biligXlsxPath, biligXlsx.xlsx)

const xlsxPopulate = await runXlsxPopulateBridge(sourceXlsx)
writeFileSync(xlsxPopulatePath, xlsxPopulate.xlsx)

const exceljs = await runExcelJsBridge(sourceXlsx)
writeFileSync(exceljsPath, exceljs.xlsx)

const output = {
  sourceXlsx: sourcePath,
  workflows: {
    biligXlsx: {
      staleCachedValueBeforeRecalc: biligXlsx.staleCachedValue,
      recalculatedValue: biligXlsx.recalculatedValue,
      outputXlsx: biligXlsxPath,
    },
    xlsxPopulate: {
      staleCachedValueBeforeRecalc: xlsxPopulate.staleCachedValue,
      recalculatedValue: xlsxPopulate.recalculatedValue,
      outputXlsx: xlsxPopulatePath,
    },
    exceljs: {
      staleCachedValueBeforeRecalc: exceljs.staleCachedValue,
      recalculatedValue: exceljs.recalculatedValue,
      patchedExceljsFormulaResult: exceljs.patchedExceljsFormulaResult,
      outputXlsx: exceljsPath,
    },
  },
  checks: {
    biligXlsx: biligXlsx.verified,
    xlsxPopulate: xlsxPopulate.verified,
    exceljs: exceljs.verified,
    allExpectedValuesMatched: biligXlsx.verified && xlsxPopulate.verified && exceljs.verified,
  },
}

if (!output.checks.allExpectedValuesMatched) {
  throw new Error(`Recalc bridge workflow failed: ${JSON.stringify(output, null, 2)}`)
}

console.log(JSON.stringify(output, null, 2))

async function createSourceXlsx() {
  const workbook = new ExcelJS.Workbook()
  const inputs = workbook.addWorksheet('Inputs')
  inputs.getCell('A1').value = 'Metric'
  inputs.getCell('B1').value = 'Value'
  inputs.getCell('A2').value = 'Units'
  inputs.getCell('B2').value = 40
  inputs.getCell('A3').value = 'List price'
  inputs.getCell('B3').value = 1200

  const summary = workbook.addWorksheet('Summary')
  summary.getCell('A1').value = 'Metric'
  summary.getCell('B1').value = 'Value'
  summary.getCell('A2').value = 'List revenue'
  summary.getCell('B2').value = {
    formula: 'Inputs!B2*Inputs!B3',
    result: 48000,
  }

  return Buffer.from(await workbook.xlsx.writeBuffer())
}

function runBiligXlsxBridge(sourceXlsxBytes) {
  const staleCachedValue = readNativeNumberCell(readXlsxTargetCell(sourceXlsxBytes, 'Summary', 'B2'), 'Summary!B2')
  const editedBytes = exportXlsxSourceLiteralPatches({
    source: sourceXlsxBytes,
    patches: [
      { sheetName: 'Inputs', address: 'B2', value: 48 },
      { sheetName: 'Inputs', address: 'B3', value: 1500 },
    ],
  })

  const recalculated = recalculateXlsx(editedBytes, {
    fileName: 'bridge-bilig-xlsx.xlsx',
    reads: ['Summary!B2'],
  })
  const recalculatedValue = readNumberCell(recalculated.reads['Summary!B2'], 'Summary!B2')

  return {
    staleCachedValue,
    recalculatedValue,
    xlsx: Buffer.from(recalculated.xlsx),
    verified: staleCachedValue === 48000 && recalculatedValue === 72000,
  }
}

async function runXlsxPopulateBridge(sourceXlsxBytes) {
  const workbook = await XlsxPopulate.fromDataAsync(sourceXlsxBytes)
  workbook.sheet('Inputs').cell('B2').value(48)
  workbook.sheet('Inputs').cell('B3').value(1500)
  const staleCachedValue = workbook.sheet('Summary').cell('B2').value()
  const editedBytes = await workbook.outputAsync('nodebuffer')

  const recalculated = recalculateXlsx(editedBytes, {
    fileName: 'bridge-xlsx-populate.xlsx',
    reads: ['Summary!B2'],
  })
  const recalculatedValue = readNumberCell(recalculated.reads['Summary!B2'], 'Summary!B2')

  return {
    staleCachedValue,
    recalculatedValue,
    xlsx: Buffer.from(recalculated.xlsx),
    verified: staleCachedValue === 48000 && recalculatedValue === 72000,
  }
}

async function runExcelJsBridge(sourceXlsxBytes) {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(sourceXlsxBytes)
  workbook.getWorksheet('Inputs').getCell('B2').value = 48
  workbook.getWorksheet('Inputs').getCell('B3').value = 1500

  const staleFormulaCell = workbook.getWorksheet('Summary').getCell('B2').value
  const recalculated = await recalculateExceljsWorkbook(workbook, {
    reads: ['Summary!B2'],
  })
  const recalculatedValue = readNumberCell(recalculated.reads['Summary!B2'], 'Summary!B2')
  const patchedFormulaCell = workbook.getWorksheet('Summary').getCell('B2').value
  const buffer = await workbook.xlsx.writeBuffer()

  return {
    staleCachedValue: readFormulaResult(staleFormulaCell, 'ExcelJS stale Summary!B2'),
    recalculatedValue,
    patchedExceljsFormulaResult: readFormulaResult(patchedFormulaCell, 'ExcelJS patched Summary!B2'),
    xlsx: Buffer.from(buffer),
    verified:
      readFormulaResult(staleFormulaCell, 'ExcelJS stale Summary!B2') === 48000 &&
      recalculatedValue === 72000 &&
      readFormulaResult(patchedFormulaCell, 'ExcelJS patched Summary!B2') === 72000,
  }
}

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

function readFormulaResult(cell, target) {
  if (cell && typeof cell === 'object' && typeof cell.result === 'number') {
    return cell.result
  }
  throw new Error(`Expected ExcelJS formula result at ${target}, got ${JSON.stringify(cell)}`)
}
