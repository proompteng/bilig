import { WorkPaper } from '@bilig/headless'
import { address } from './workpaper-benchmark-fixtures.js'
import {
  buildSlidingAggregateSheet,
  buildConditionalAggregationSheet,
  buildLookupSheet,
  buildApproxLookupSheet,
  buildMixedFrontierSheet,
  buildOverlappingAggregateSheet,
  buildStructuralColumnSheet,
  buildParserCacheMixedTemplateSheet,
  buildParserCacheTemplateSheet,
  buildMixedContentSheet,
  buildValueFormulaRows,
  buildBatchMultiColumnRows,
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

export function measureWorkPaperSuspendedBatchSingleColumnEditSample(editCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildValueFormulaRows(editCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation()
      for (let row = 0; row < editCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 7)
      }
      return workbook.resumeEvaluation()
    },
    () => ({
      sampleFormulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureHyperFormulaSuspendedBatchSingleColumnEditSample(editCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildValueFormulaRows(editCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation()
      for (let row = 0; row < editCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 7)
      }
      return workbook.resumeEvaluation()
    },
    () => ({
      sampleFormulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
      width: workbook.getSheetDimensions(sheetId).width,
    }),
  )
}

export function measureWorkPaperSuspendedBatchMultiColumnEditSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildBatchMultiColumnRows(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation()
      for (let row = 0; row < rowCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 3)
        workbook.setCellContents(address(sheetId, row, 1), row * 5)
      }
      return workbook.resumeEvaluation()
    },
    () => ({
      sampleSumValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 2))),
      sampleProductValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 3))),
    }),
  )
}

export function measureHyperFormulaSuspendedBatchMultiColumnEditSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildBatchMultiColumnRows(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation()
      for (let row = 0; row < rowCount; row += 1) {
        workbook.setCellContents(address(sheetId, row, 0), row * 3)
        workbook.setCellContents(address(sheetId, row, 1), row * 5)
      }
      return workbook.resumeEvaluation()
    },
    () => ({
      sampleSumValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 2))),
      sampleProductValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 3))),
    }),
  )
}

export function measureWorkPaperBatchSingleColumnUndoSample(editCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildValueFormulaRows(editCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => {
      workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
        }
      })
      return workbook.undo()
    },
    (undoChanges) => ({
      undoChangeCount: undoChanges.length,
      restoredValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 0))),
      restoredFormulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
    }),
  )
}

export function measureHyperFormulaBatchSingleColumnUndoSample(editCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildValueFormulaRows(editCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => {
      workbook.batch(() => {
        for (let row = 0; row < editCount; row += 1) {
          workbook.setCellContents(address(sheetId, row, 0), row * 3)
        }
      })
      return workbook.undo()
    },
    (undoChanges) => ({
      undoChangeCount: undoChanges.length,
      restoredValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 0))),
      restoredFormulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, editCount - 1, 1))),
    }),
  )
}

export function measureWorkPaperRebuildAndRecalculateSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildParserCacheTemplateSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.rebuildAndRecalculate(),
    () => ({
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 4))),
      dimensions: workbook.getSheetDimensions(sheetId),
    }),
  )
}

export function measureHyperFormulaRebuildAndRecalculateSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildParserCacheTemplateSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.rebuildAndRecalculate(),
    () => ({
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 4))),
      dimensions: workbook.getSheetDimensions(sheetId),
    }),
  )
}

export function measureWorkPaperConfigToggleSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildLookupSheet(rowCount) }, { useColumnIndex: false })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.updateConfig({ useColumnIndex: true }),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaConfigToggleSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY, useColumnIndex: false },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.updateConfig({ useColumnIndex: true }),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureWorkPaperStructuralInsertRowsSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildOverlappingAggregateSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.addRows(sheetId, Math.floor(rowCount / 2), 1),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount, 1))),
    }),
  )
}

export function measureHyperFormulaStructuralInsertRowsSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildOverlappingAggregateSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.addRows(sheetId, [Math.floor(rowCount / 2), 1]),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount, 1))),
    }),
  )
}

export function measureWorkPaperStructuralDeleteRowsSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildOverlappingAggregateSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.removeRows(sheetId, Math.floor(rowCount / 2), 1),
    () => {
      const dimensions = workbook.getSheetDimensions(sheetId)
      return {
        dimensions,
        terminalSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, Math.max(dimensions.height - 1, 0), 1))),
      }
    },
  )
}

export function measureHyperFormulaStructuralDeleteRowsSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildOverlappingAggregateSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.removeRows(sheetId, [Math.floor(rowCount / 2), 1]),
    () => {
      const dimensions = workbook.getSheetDimensions(sheetId)
      return {
        dimensions,
        terminalSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, Math.max(dimensions.height - 1, 0), 1))),
      }
    },
  )
}

export function measureWorkPaperStructuralMoveRowsSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildOverlappingAggregateSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  const start = Math.floor(rowCount / 2)
  return measureMutationSample(
    workbook,
    () => workbook.moveRows(sheetId, start, 1, 0),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      headValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 0))),
    }),
  )
}

export function measureHyperFormulaStructuralMoveRowsSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildOverlappingAggregateSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  const start = Math.floor(rowCount / 2)
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.moveRows(sheetId, start, 1, 0),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      headValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 0))),
    }),
  )
}

export function measureWorkPaperStructuralInsertColumnsSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildStructuralColumnSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.addColumns(sheetId, 1, 1),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalFormula: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 4))),
    }),
  )
}

export function measureHyperFormulaStructuralInsertColumnsSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildStructuralColumnSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.addColumns(sheetId, [1, 1]),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalFormula: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 4))),
    }),
  )
}

export function measureWorkPaperStructuralDeleteColumnsSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildStructuralColumnSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.removeColumns(sheetId, 1, 1),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 0))),
    }),
  )
}

export function measureHyperFormulaStructuralDeleteColumnsSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildStructuralColumnSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.removeColumns(sheetId, [1, 1]),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 0))),
    }),
  )
}

export function measureWorkPaperStructuralMoveColumnsSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildStructuralColumnSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.moveColumns(sheetId, 1, 1, 0),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      headValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 0))),
      terminalFormula: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 3))),
    }),
  )
}

export function measureHyperFormulaStructuralMoveColumnsSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildStructuralColumnSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.moveColumns(sheetId, 1, 1, 0),
    () => ({
      dimensions: workbook.getSheetDimensions(sheetId),
      headValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 0))),
      terminalFormula: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 3))),
    }),
  )
}

export function measureWorkPaperOverlappingAggregateSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildOverlappingAggregateSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 1))),
    }),
  )
}

export function measureHyperFormulaOverlappingAggregateSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildOverlappingAggregateSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 1))),
    }),
  )
}

export function measureWorkPaperSlidingAggregateSample(rowCount: number, window: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildSlidingAggregateSheet(rowCount, window),
  })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 1))),
      leadingSum: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 1))),
    }),
  )
}

export function measureHyperFormulaSlidingAggregateSample(rowCount: number, window: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildSlidingAggregateSheet(rowCount, window)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 1))),
      leadingSum: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 1))),
    }),
  )
}

export function measureWorkPaperConditionalAggregationSample(rowCount: number, formulaCopies: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildConditionalAggregationSheet(rowCount, formulaCopies),
  })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 1), rowCount * 2),
    () => ({
      sumifValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
      countifValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4 + formulaCopies))),
    }),
  )
}

export function measureHyperFormulaConditionalAggregationSample(rowCount: number, formulaCopies: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildConditionalAggregationSheet(rowCount, formulaCopies)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 1), rowCount * 2),
    () => ({
      sumifValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
      countifValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4 + formulaCopies))),
    }),
  )
}

export function measureWorkPaperConditionalAggregationCriteriaEditSample(rowCount: number, formulaCopies: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({
    Bench: buildConditionalAggregationSheet(rowCount, formulaCopies),
  })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), 'B'),
    () => ({
      sumifValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
      countifValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4 + formulaCopies))),
    }),
  )
}

export function measureHyperFormulaConditionalAggregationCriteriaEditSample(rowCount: number, formulaCopies: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildConditionalAggregationSheet(rowCount, formulaCopies)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 3), 'B'),
    () => ({
      sumifValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
      countifValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4 + formulaCopies))),
    }),
  )
}

export function measureWorkPaperParserCacheTemplateSample(rowCount: number): BenchmarkSample {
  return measureWorkPaperBuildFromSheets({ Bench: buildParserCacheTemplateSheet(rowCount) }, (workbook) => {
    const sheetId = workbook.getSheetId('Bench')!
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 4))),
    }
  })
}

export function measureHyperFormulaParserCacheTemplateSample(rowCount: number): BenchmarkSample {
  return measureHyperFormulaBuildFromSheets({ Bench: toHyperFormulaSheet(buildParserCacheTemplateSheet(rowCount)) }, (workbook) => {
    const sheetId = workbook.getSheetId('Bench')!
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 4))),
    }
  })
}

export function measureWorkPaperParserCacheMixedTemplateSample(rowCount: number): BenchmarkSample {
  return measureWorkPaperBuildFromSheets({ Bench: buildParserCacheMixedTemplateSheet(rowCount) }, (workbook) => {
    const sheetId = workbook.getSheetId('Bench')!
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 4))),
    }
  })
}

export function measureHyperFormulaParserCacheMixedTemplateSample(rowCount: number): BenchmarkSample {
  return measureHyperFormulaBuildFromSheets({ Bench: toHyperFormulaSheet(buildParserCacheMixedTemplateSheet(rowCount)) }, (workbook) => {
    const sheetId = workbook.getSheetId('Bench')!
    return {
      dimensions: workbook.getSheetDimensions(sheetId),
      terminalValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 4))),
    }
  })
}

export function measureWorkPaperRebuildRuntimeFromSnapshotSample(rowCount: number): BenchmarkSample {
  const seeded = WorkPaper.buildFromSheets({
    Bench: buildMixedContentSheet(rowCount),
    Templates: buildParserCacheMixedTemplateSheet(Math.max(Math.floor(rowCount / 2), 2)),
  })
  const serializedSheets = seeded.getAllSheetsSerialized()
  seeded.dispose()
  return measureWorkPaperBuildFromSheets(serializedSheets, (workbook) => {
    const benchId = workbook.getSheetId('Bench')!
    return {
      sheetCount: workbook.countSheets(),
      benchTerminal: normalizeWorkPaperValue(workbook.getCellValue(address(benchId, rowCount - 1, 5))),
    }
  })
}

export function measureHyperFormulaRebuildRuntimeFromSnapshotSample(rowCount: number): BenchmarkSample {
  const seeded = HyperFormula.buildFromSheets(
    {
      Bench: toHyperFormulaSheet(buildMixedContentSheet(rowCount)),
      Templates: toHyperFormulaSheet(buildParserCacheMixedTemplateSheet(Math.max(Math.floor(rowCount / 2), 2))),
    },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const serializedSheets = seeded.getAllSheetsSerialized()
  seeded.destroy()
  return measureHyperFormulaBuildFromSheets(serializedSheets, (workbook) => {
    const benchId = workbook.getSheetId('Bench')!
    return {
      sheetCount: workbook.countSheets(),
      benchTerminal: normalizeHyperFormulaValue(workbook.getCellValue(address(benchId, rowCount - 1, 5))),
    }
  })
}

export function measureWorkPaperMixedFrontierSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildMixedFrontierSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalAggregate: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 2))),
      terminalFanout: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, rowCount - 1, 1))),
    }),
  )
}

export function measureHyperFormulaMixedFrontierSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildMixedFrontierSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, 0, 0), 99),
    () => ({
      terminalAggregate: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 2))),
      terminalFanout: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, rowCount - 1, 1))),
    }),
  )
}

export function measureWorkPaperIndexedLookupAfterColumnWriteSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildLookupSheet(rowCount) }, { useColumnIndex: true })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1_000),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaIndexedLookupAfterColumnWriteSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY, useColumnIndex: true },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1_000),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureWorkPaperIndexedLookupAfterBatchWriteSample(rowCount: number, editCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildLookupSheet(rowCount) }, { useColumnIndex: true })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation()
      for (let index = 0; index < editCount; index += 1) {
        const row = rowCount - index
        workbook.setCellContents(address(sheetId, row, 0), row + 10_000)
      }
      return workbook.resumeEvaluation()
    },
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaIndexedLookupAfterBatchWriteSample(rowCount: number, editCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY, useColumnIndex: true },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => {
      workbook.suspendEvaluation()
      for (let index = 0; index < editCount; index += 1) {
        const row = rowCount - index
        workbook.setCellContents(address(sheetId, row, 0), row + 10_000)
      }
      return workbook.resumeEvaluation()
    },
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureWorkPaperApproximateLookupAfterColumnWriteSample(rowCount: number): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Bench: buildApproxLookupSheet(rowCount) })
  const sheetId = workbook.getSheetId('Bench')!
  return measureMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1),
    () => ({
      formulaValue: normalizeWorkPaperValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}

export function measureHyperFormulaApproximateLookupAfterColumnWriteSample(rowCount: number): BenchmarkSample {
  const workbook = HyperFormula.buildFromSheets(
    { Bench: toHyperFormulaSheet(buildApproxLookupSheet(rowCount)) },
    { licenseKey: HYPERFORMULA_LICENSE_KEY },
  )
  const sheetId = workbook.getSheetId('Bench')!
  return measureHyperFormulaMutationSample(
    workbook,
    () => workbook.setCellContents(address(sheetId, rowCount, 0), rowCount + 1),
    () => ({
      formulaValue: normalizeHyperFormulaValue(workbook.getCellValue(address(sheetId, 0, 4))),
    }),
  )
}
