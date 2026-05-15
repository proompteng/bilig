import { WorkPaper } from '@bilig/headless'
import {
  address,
  buildCrossSheetAggregateSheets,
  buildCrossSheetDashboardSheets,
  buildCrossSheetScalarFanoutSheets,
  buildFormulaGridSheet,
  buildIndexMatchExactSheet,
  buildIndexReferenceSheet,
  buildRectangularBlockFormulaRows,
  buildRectangularBlockFormulaSheet,
  buildSparseWideSheet,
  range,
  textLookupKey,
} from './workpaper-benchmark-fixtures.js'
import {
  HyperFormula,
  measureHyperFormulaBuildFromSheets,
  measureHyperFormulaMutationSample,
  measureMutationSample,
  measureWorkPaperBuildFromSheets,
  normalizeHyperFormulaValue,
  normalizeWorkPaperValue,
  toHyperFormulaSheet,
  type BenchmarkSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { HYPERFORMULA_LICENSE_KEY } from './benchmark-workpaper-vs-hyperformula.js'

export function measureWorkPaperCrossSheetDashboardBuildSample(sheetCount: number, rowCount: number): BenchmarkSample {
  return measureWorkPaperBuildFromSheets(buildCrossSheetDashboardSheets(sheetCount, rowCount), (workbook) => {
    const summarySheetId = workbook.getSheetId('Summary')!
    return {
      leadingDataTotal: normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, 0, 0))),
      sheetCount: workbook.countSheets(),
      terminalDataTotal: normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, sheetCount - 1, 2))),
    }
  })
}

export function measureHyperFormulaCrossSheetDashboardBuildSample(sheetCount: number, rowCount: number): BenchmarkSample {
  return measureHyperFormulaBuildFromSheets(toHyperFormulaSheets(buildCrossSheetDashboardSheets(sheetCount, rowCount)), (workbook) => {
    const summarySheetId = workbook.getSheetId('Summary')!
    return {
      leadingDataTotal: normalizeHyperFormulaValue(workbook.getCellValue(address(summarySheetId, 0, 0))),
      sheetCount: workbook.countSheets(),
      terminalDataTotal: normalizeHyperFormulaValue(workbook.getCellValue(address(summarySheetId, sheetCount - 1, 2))),
    }
  })
}

export function measureWorkPaperCrossSheetScalarFanoutSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets(buildCrossSheetScalarFanoutSheets(rowCount))
  const dataSheetId = workbook.getSheetId('Data')!
  const summarySheetId = workbook.getSheetId('Summary')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(dataSheetId, 0, 0), 99),
    () => ({
      leadingValue: normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, 0, 0))),
      sheetCount: workbook.countSheets(),
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, rowCount - 1, 0))),
    }),
  )
}

export function measureHyperFormulaCrossSheetScalarFanoutSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(toHyperFormulaSheets(buildCrossSheetScalarFanoutSheets(rowCount)), {
    licenseKey: HYPERFORMULA_LICENSE_KEY,
  })
  const dataSheetId = workbook.getSheetId('Data')!
  const summarySheetId = workbook.getSheetId('Summary')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(dataSheetId, 0, 0), 99),
    () => ({
      leadingValue: normalizeHyperFormulaValue(workbook.getCellValue(address(summarySheetId, 0, 0))),
      sheetCount: workbook.countSheets(),
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(summarySheetId, rowCount - 1, 0))),
    }),
  )
}

export function measureWorkPaperCrossSheetAggregateSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets(buildCrossSheetAggregateSheets(rowCount))
  const dataSheetId = workbook.getSheetId('Data')!
  const summarySheetId = workbook.getSheetId('Summary')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(dataSheetId, 0, 0), 99),
    () => ({
      leadingSum: normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, 0, 0))),
      sheetCount: workbook.countSheets(),
      terminalSum: normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, rowCount - 1, 0))),
    }),
  )
}

export function measureHyperFormulaCrossSheetAggregateSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(toHyperFormulaSheets(buildCrossSheetAggregateSheets(rowCount)), {
    licenseKey: HYPERFORMULA_LICENSE_KEY,
  })
  const dataSheetId = workbook.getSheetId('Data')!
  const summarySheetId = workbook.getSheetId('Summary')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(dataSheetId, 0, 0), 99),
    () => ({
      leadingSum: normalizeHyperFormulaValue(workbook.getCellValue(address(summarySheetId, 0, 0))),
      sheetCount: workbook.countSheets(),
      terminalSum: normalizeHyperFormulaValue(workbook.getCellValue(address(summarySheetId, rowCount - 1, 0))),
    }),
  )
}

export function measureWorkPaperCrossSheetDashboardRecalcSample(sheetCount: number, rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets(buildCrossSheetDashboardSheets(sheetCount, rowCount))
  const dataSheetId = workbook.getSheetId('Data1')!
  const summarySheetId = workbook.getSheetId('Summary')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(dataSheetId, 0, 1), 999),
    () => ({
      leadingDataTotal: normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, 0, 1))),
      sheetCount: workbook.countSheets(),
      terminalDataTotal: normalizeWorkPaperValue(workbook.getCellValue(address(summarySheetId, sheetCount - 1, 2))),
    }),
  )
}

export function measureHyperFormulaCrossSheetDashboardRecalcSample(sheetCount: number, rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(toHyperFormulaSheets(buildCrossSheetDashboardSheets(sheetCount, rowCount)), {
    licenseKey: HYPERFORMULA_LICENSE_KEY,
  })
  const dataSheetId = workbook.getSheetId('Data1')!
  const summarySheetId = workbook.getSheetId('Summary')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(dataSheetId, 0, 1), 999),
    () => ({
      leadingDataTotal: normalizeHyperFormulaValue(workbook.getCellValue(address(summarySheetId, 0, 1))),
      sheetCount: workbook.countSheets(),
      terminalDataTotal: normalizeHyperFormulaValue(workbook.getCellValue(address(summarySheetId, sheetCount - 1, 2))),
    }),
  )
}

export function measureWorkPaperRectangularBatchEditSample(rowCount: number, inputCols: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildRectangularBlockFormulaSheet(rowCount, inputCols) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          for (let col = 0; col < inputCols; col += 1) {
            workbook.setCellContents(address(sheetId, row, col), (row + 1) * (col + 2))
          }
        }
      }),
    () => ({
      leadingSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, inputCols))),
      terminalSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, inputCols))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureHyperFormulaRectangularBatchEditSample(rowCount: number, inputCols: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildRectangularBlockFormulaSheet(rowCount, inputCols)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          for (let col = 0; col < inputCols; col += 1) {
            workbook.setCellContents(address(sheetId, row, col), (row + 1) * (col + 2))
          }
        }
      }),
    () => ({
      leadingSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, inputCols))),
      terminalSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, inputCols))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureWorkPaperBatchClearRectangularBlockSample(rowCount: number, inputCols: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildRectangularBlockFormulaSheet(rowCount, inputCols) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          for (let col = 0; col < inputCols; col += 1) {
            workbook.setCellContents(address(sheetId, row, col), null)
          }
        }
      }),
    () => ({
      leadingSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, inputCols))),
      terminalSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, inputCols))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureHyperFormulaBatchClearRectangularBlockSample(rowCount: number, inputCols: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildRectangularBlockFormulaSheet(rowCount, inputCols)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          for (let col = 0; col < inputCols; col += 1) {
            workbook.setCellContents(address(sheetId, row, col), null)
          }
        }
      }),
    () => ({
      leadingSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, inputCols))),
      terminalSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, inputCols))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureWorkPaperAppendFormulaRowsSample(rowCount: number, inputCols: number, appendCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildRectangularBlockFormulaSheet(rowCount, inputCols) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        workbook.addRows(sheetId, rowCount, appendCount)
        workbook.setCellContents(address(sheetId, rowCount, 0), buildRectangularBlockFormulaRows(appendCount, inputCols, rowCount + 1))
      }),
    () => ({
      appendedLeadingSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount, inputCols))),
      appendedTerminalSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount + appendCount - 1, inputCols))),
      dimensions: workbook.getSheetDimensions(sheetId),
    }),
  )
}

export function measureHyperFormulaAppendFormulaRowsSample(rowCount: number, inputCols: number, appendCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildRectangularBlockFormulaSheet(rowCount, inputCols)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () =>
      workbook.batch(() => {
        workbook.addRows(sheetId, [rowCount, appendCount])
        workbook.setCellContents(
          address(sheetId, rowCount, 0),
          toHyperFormulaSheet(buildRectangularBlockFormulaRows(appendCount, inputCols, rowCount + 1)),
        )
      }),
    () => ({
      appendedLeadingSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount, inputCols))),
      appendedTerminalSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount + appendCount - 1, inputCols))),
      dimensions: workbook.getSheetDimensions(sheetId),
    }),
  )
}

export function measureWorkPaperSparseWideRangeReadSample(rowCount: number, colCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildSparseWideSheet(rowCount, colCount) })
  const sheetId = workbook.getSheetId('Bench')!
  const middleCol = Math.floor(colCount / 2)
  return measureMutationSample(
    workbook,
    () => workbook.getRangeValues(range(sheetId, 0, 0, rowCount - 1, colCount - 1)),
    (values) => {
      const lastRow = values.at(-1)
      return {
        emptyValue: normalizeWorkPaperValue(values[0]?.[1]),
        middleValue: normalizeWorkPaperValue(lastRow?.[middleCol]),
        readCols: values[0]?.length ?? 0,
        readRows: values.length,
        terminalValue: normalizeWorkPaperValue(lastRow?.at(-1)),
        topLeftValue: normalizeWorkPaperValue(values[0]?.[0]),
      }
    },
  )
}

export function measureHyperFormulaSparseWideRangeReadSample(rowCount: number, colCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildSparseWideSheet(rowCount, colCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const middleCol = Math.floor(colCount / 2)
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.getRangeValues(range(sheetId, 0, 0, rowCount - 1, colCount - 1)),
    (values) => {
      const lastRow = values.at(-1)
      return {
        emptyValue: normalizeHyperFormulaValue(values[0]?.[1]),
        middleValue: normalizeHyperFormulaValue(lastRow?.[middleCol]),
        readCols: values[0]?.length ?? 0,
        readRows: values.length,
        terminalValue: normalizeHyperFormulaValue(lastRow?.at(-1)),
        topLeftValue: normalizeHyperFormulaValue(values[0]?.[0]),
      }
    },
  )
}

export function measureWorkPaperFormulaGridRangeReadSample(rowCount: number, inputCols: number, formulaCols: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildFormulaGridSheet(rowCount, inputCols, formulaCols) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.getRangeValues(range(sheetId, 0, inputCols, rowCount - 1, inputCols + formulaCols - 1)),
    (values) => {
      const lastRow = values.at(-1)
      return {
        leadingFormulaValue: normalizeWorkPaperValue(values[0]?.[0]),
        readCols: values[0]?.length ?? 0,
        readRows: values.length,
        terminalFormulaValue: normalizeWorkPaperValue(lastRow?.at(-1)),
      }
    },
  )
}

export function measureHyperFormulaFormulaGridRangeReadSample(rowCount: number, inputCols: number, formulaCols: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildFormulaGridSheet(rowCount, inputCols, formulaCols)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.getRangeValues(range(sheetId, 0, inputCols, rowCount - 1, inputCols + formulaCols - 1)),
    (values) => {
      const lastRow = values.at(-1)
      return {
        leadingFormulaValue: normalizeHyperFormulaValue(values[0]?.[0]),
        readCols: values[0]?.length ?? 0,
        readRows: values.length,
        terminalFormulaValue: normalizeHyperFormulaValue(lastRow?.at(-1)),
      }
    },
  )
}

export function measureWorkPaperIndexMatchExactSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildIndexMatchExactSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), textLookupKey(rowCount - 1)),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaIndexMatchExactSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildIndexMatchExactSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), textLookupKey(rowCount - 1)),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureWorkPaperIndexReferenceSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildIndexReferenceSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), rowCount - 1),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaIndexReferenceSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildIndexReferenceSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), rowCount - 1),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

function toHyperFormulaSheets(
  sheets: Record<string, ReadonlyArray<ReadonlyArray<unknown>>>,
): Record<string, ReturnType<typeof toHyperFormulaSheet>> {
  return Object.fromEntries(Object.entries(sheets).map(([sheetName, sheet]) => [sheetName, toHyperFormulaSheet(sheet)]))
}
