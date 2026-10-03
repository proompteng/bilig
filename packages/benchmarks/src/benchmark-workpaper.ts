import { performance } from 'node:perf_hooks'
import type { RecalcMetrics } from '@bilig/protocol'
import { WorkPaper, type WorkPaperSheet } from '../../workpaper/src/work-paper.js'
import { measureMemory, sampleMemory, type MemoryMeasurement } from './metrics.js'
import {
  address,
  buildDenseLiteralSheet,
  buildDynamicArraySheet,
  buildFormulaGridSheet,
  buildFormulaChainRow,
  buildLookupSheet,
  buildValueFormulaRows,
} from './workpaper-benchmark-fixtures.js'

export type WorkPaperBenchmarkScenario =
  | 'workpaper-build'
  | 'workpaper-single-edit'
  | 'workpaper-batch-edit'
  | 'workpaper-fresh-range-value-write'
  | 'workpaper-existing-range-value-write'
  | 'workpaper-range-read'
  | 'workpaper-range-value-block'
  | 'workpaper-range-formula-read'
  | 'workpaper-range-serialized-read'
  | 'workpaper-lookup'
  | 'workpaper-dynamic-array'
  | 'workpaper-calculate-formula'

export interface WorkPaperBenchmarkResult {
  scenario: WorkPaperBenchmarkScenario
  elapsedMs: number
  memory: MemoryMeasurement
  details: Record<string, number | boolean | string>
  metrics?: RecalcMetrics
}

export async function runWorkPaperBenchmarkSuite(): Promise<WorkPaperBenchmarkResult[]> {
  return [
    runWorkPaperBuildBenchmark(),
    runWorkPaperSingleEditBenchmark(),
    runWorkPaperBatchEditBenchmark(),
    runWorkPaperFreshRangeValueWriteBenchmark(),
    runWorkPaperExistingRangeValueWriteBenchmark(),
    runWorkPaperRangeReadBenchmark(),
    runWorkPaperRangeValueBlockBenchmark(),
    runWorkPaperRangeFormulaReadBenchmark(),
    runWorkPaperRangeSerializedReadBenchmark(),
    runWorkPaperLookupBenchmark(),
    runWorkPaperDynamicArrayBenchmark(),
    runWorkPaperCalculateFormulaBenchmark(),
  ]
}

export function runWorkPaperBuildBenchmark(rows = 160, cols = 24): WorkPaperBenchmarkResult {
  const sheet = buildDenseLiteralSheet(rows, cols)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const workbook = WorkPaper.buildFromSheets({ Bench: sheet })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const sheetId = workbook.getSheetId('Bench')!

  return {
    scenario: 'workpaper-build',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rows,
      cols,
      materializedCells: rows * cols,
      width: workbook.getSheetDimensions(sheetId).width,
      height: workbook.getSheetDimensions(sheetId).height,
    },
    metrics: workbook.getStats().lastMetrics,
  }
}

export function runWorkPaperSingleEditBenchmark(downstreamCount = 2_000): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: [buildFormulaChainRow(downstreamCount)],
  })
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setCellContents(address(sheetId, 0, 0), 99)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    scenario: 'workpaper-single-edit',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      downstreamCount,
      changeCount: changes.length,
      terminalFormula: workbook.getCellFormula(address(sheetId, 0, downstreamCount)) ?? '',
    },
    metrics: workbook.getStats().lastMetrics,
  }
}

export function runWorkPaperBatchEditBenchmark(editCount = 500): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildValueFormulaRows(editCount),
  })
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.batch(() => {
    for (let row = 0; row < editCount; row += 1) {
      workbook.setCellContents(address(sheetId, row, 0), row * 3)
    }
  })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    scenario: 'workpaper-batch-edit',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      editCount,
      changeCount: changes.length,
      sampleFormulaValue: JSON.stringify(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
    },
    metrics: workbook.getStats().lastMetrics,
  }
}

export function runWorkPaperFreshRangeValueWriteBenchmark(rows = 400, cols = 20): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({ Bench: [] })
  const sheetId = workbook.getSheetId('Bench')!
  const values = buildDenseNumericRangeValues(rows, cols)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setSheetRangeValues(sheetId, 0, 0, values)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const dimensions = workbook.getSheetDimensions(sheetId)

  return {
    scenario: 'workpaper-fresh-range-value-write',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rows,
      cols,
      materializedCells: rows * cols,
      changeCount: changes.length,
      width: dimensions.width,
      height: dimensions.height,
      terminalValue: JSON.stringify(workbook.getCellValue(address(sheetId, rows - 1, cols - 1))),
    },
    metrics: workbook.getStats().lastMetrics,
  }
}

export function runWorkPaperRangeReadBenchmark(rows = 240, cols = 24): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildDenseLiteralSheet(rows, cols),
  })
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const values = workbook.getRangeValues({
    start: address(sheetId, 0, 0),
    end: address(sheetId, rows - 1, cols - 1),
  })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    scenario: 'workpaper-range-read',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rows,
      cols,
      materializedCells: rows * cols,
      readRows: values.length,
      readCols: values[0]?.length ?? 0,
    },
  }
}

export function runWorkPaperRangeValueBlockBenchmark(rowCount = 240, inputCols = 8, formulaCols = 16): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildFormulaGridSheet(rowCount, inputCols, formulaCols),
  })
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const block = workbook.getRangeValueBlock({
    start: address(sheetId, 0, inputCols),
    end: address(sheetId, rowCount - 1, inputCols + formulaCols - 1),
  })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const terminalOffset = (block.rowCount - 1) * block.colCount + block.colCount - 1

  return {
    scenario: 'workpaper-range-value-block',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rowCount,
      inputCols,
      formulaCols,
      materializedCells: rowCount * formulaCols,
      readRows: block.rowCount,
      readCols: block.colCount,
      terminalTag: block.tags[terminalOffset] ?? 0,
      terminalNumber: block.numbers[terminalOffset] ?? 0,
    },
  }
}

export function runWorkPaperRangeFormulaReadBenchmark(rowCount = 20_000): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildSparseTailFormulaSheet(rowCount),
  })
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const formulas = workbook.getRangeFormulas({
    start: address(sheetId, 0, 0),
    end: address(sheetId, rowCount, 1),
  })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    scenario: 'workpaper-range-formula-read',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rowCount: rowCount + 1,
      cols: 2,
      materializedCells: (rowCount + 1) * 2,
      readRows: formulas.length,
      readCols: formulas[0]?.length ?? 0,
      terminalFormula: formulas[rowCount]?.[1] ?? '',
    },
  }
}

export function runWorkPaperRangeSerializedReadBenchmark(rowCount = 20_000): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildSparseTailFormulaSheet(rowCount),
  })
  const sheetId = workbook.getSheetId('Bench')!
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const serialized = workbook.getRangeSerialized({
    start: address(sheetId, 0, 0),
    end: address(sheetId, rowCount, 1),
  })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    scenario: 'workpaper-range-serialized-read',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rowCount: rowCount + 1,
      cols: 2,
      materializedCells: (rowCount + 1) * 2,
      readRows: serialized.length,
      readCols: serialized[0]?.length ?? 0,
      terminalFormula: String(serialized[rowCount]?.[1] ?? ''),
    },
  }
}

export function runWorkPaperLookupBenchmark(rowCount = 5_000): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets(
    {
      Bench: buildLookupSheet(rowCount),
    },
    {
      useColumnIndex: true,
    },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const targetAddress = address(sheetId, 0, 3)
  const formulaAddress = address(sheetId, 0, 4)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setCellContents(targetAddress, rowCount)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    scenario: 'workpaper-lookup',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rowCount,
      changeCount: changes.length,
      formulaValue: JSON.stringify(workbook.getCellValue(formulaAddress)),
    },
    metrics: workbook.getStats().lastMetrics,
  }
}

export function runWorkPaperExistingRangeValueWriteBenchmark(rows = 400, cols = 20): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildDenseNumericRangeValues(rows, cols) })
  const sheetId = workbook.getSheetId('Bench')!
  const values = buildDenseNumericRangeValues(rows, cols, 10_000)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setSheetRangeValues(sheetId, 0, 0, values)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  const dimensions = workbook.getSheetDimensions(sheetId)

  return {
    scenario: 'workpaper-existing-range-value-write',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rows,
      cols,
      materializedCells: rows * cols,
      changeCount: changes.length,
      width: dimensions.width,
      height: dimensions.height,
      terminalValue: JSON.stringify(workbook.getCellValue(address(sheetId, rows - 1, cols - 1))),
    },
    metrics: workbook.getStats().lastMetrics,
  }
}

function buildDenseNumericRangeValues(rows: number, cols: number, offset = 0): number[][] {
  return Array.from({ length: rows }, (_rowValue, row) => Array.from({ length: cols }, (_colValue, col) => (row + 1) * (col + 1) + offset))
}

function buildSparseTailFormulaSheet(rowCount: number): WorkPaperSheet {
  const rows: Array<Array<null | string>> = Array.from({ length: rowCount }, () => [null, null])
  rows.push([null, `=SUM(A1:A${rowCount})`])
  return rows
}

export function runWorkPaperDynamicArrayBenchmark(rowCount = 750): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildDynamicArraySheet(rowCount),
  })
  const sheetId = workbook.getSheetId('Bench')!
  const thresholdAddress = address(sheetId, 0, 1)
  const spillAnchor = address(sheetId, 0, 2)
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const changes = workbook.setCellContents(thresholdAddress, rowCount - 10)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()

  return {
    scenario: 'workpaper-dynamic-array',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      rowCount,
      changeCount: changes.length,
      spillIsArray: workbook.isCellPartOfArray(spillAnchor),
      spillHeight: workbook.getSheetDimensions(sheetId).height,
    },
    metrics: workbook.getStats().lastMetrics,
  }
}

export function runWorkPaperCalculateFormulaBenchmark(iterationCount = 2_000): WorkPaperBenchmarkResult {
  const workbook = WorkPaper.buildEmpty()
  const formulas = [
    '=SUM(1,2,3)',
    '=IF(TRUE,"yes","no")',
    '=PMT(0.06/12,12,100000)',
    '=ROUND(SQRT(121),2)',
    '=CONCATENATE("baz","-","bar")',
    '=LEN("foo")+LEN("bars")',
    '=MIN(100,120,220)+MAX(100,120,220)',
  ]
  const memoryBefore = sampleMemory()
  const started = performance.now()
  let sampleValue = ''
  for (let index = 0; index < iterationCount; index += 1) {
    sampleValue = JSON.stringify(workbook.calculateFormula(formulas[index % formulas.length]!))
  }
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  workbook.dispose()

  return {
    scenario: 'workpaper-calculate-formula',
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    details: {
      formulaCount: formulas.length,
      iterationCount,
      sampleValue,
    },
  }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  console.log(JSON.stringify(await runWorkPaperBenchmarkSuite(), null, 2))
}
