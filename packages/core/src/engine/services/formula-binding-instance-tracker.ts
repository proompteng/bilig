import { retargetFormulaInstance } from '../../formula/structural-retargeting.js'
import type { RuntimeFormula } from '../runtime-state.js'
import { getFormulaBindingFamilyShapeKey, type FormulaBindingFamilyShapeKeyCache } from './formula-binding-family-shape-key.js'
import type { CreateEngineFormulaBindingServiceArgs, FormulaOwnerPosition } from './formula-binding-service-types.js'

export interface FormulaBindingInstanceTracker {
  readonly recordFormulaInstanceNow: (
    cellIndex: number,
    source: string,
    templateId: number | undefined,
    ownerPosition?: FormulaOwnerPosition,
  ) => void
  readonly registerFormulaFamilyNow: (cellIndex: number, formula: RuntimeFormula, ownerPosition?: FormulaOwnerPosition) => void
  readonly rebuildFormulaInstancesNow: () => void
}

export function createFormulaBindingInstanceTracker(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly formulaFamilyShapeKeyCache: FormulaBindingFamilyShapeKeyCache
}): FormulaBindingInstanceTracker {
  const recordFormulaInstanceNow = (
    cellIndex: number,
    source: string,
    templateId: number | undefined,
    ownerPosition?: FormulaOwnerPosition,
  ): void => {
    if (ownerPosition) {
      const existing = args.serviceArgs.formulaInstances.get(cellIndex)
      if (existing) {
        args.serviceArgs.formulaInstances.upsert(
          retargetFormulaInstance(existing, {
            sheetName: ownerPosition.sheetName,
            row: ownerPosition.row,
            col: ownerPosition.col,
            source,
            ...(templateId !== undefined ? { templateId } : {}),
          }),
        )
        return
      }
      args.serviceArgs.formulaInstances.upsert({
        cellIndex,
        sheetName: ownerPosition.sheetName,
        row: ownerPosition.row,
        col: ownerPosition.col,
        source,
        ...(templateId !== undefined ? { templateId } : {}),
      })
      return
    }
    const sheetId = args.serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]
    const position = args.serviceArgs.state.workbook.getCellPosition(cellIndex)
    if (sheetId === undefined || !position) {
      args.serviceArgs.formulaInstances.delete(cellIndex)
      return
    }
    const sheetName = args.serviceArgs.state.workbook.getSheetNameById(sheetId)
    if (!sheetName) {
      args.serviceArgs.formulaInstances.delete(cellIndex)
      return
    }
    const existing = args.serviceArgs.formulaInstances.get(cellIndex)
    if (existing) {
      args.serviceArgs.formulaInstances.upsert(
        retargetFormulaInstance(existing, {
          sheetName,
          row: position.row,
          col: position.col,
          source,
          ...(templateId !== undefined ? { templateId } : {}),
        }),
      )
      return
    }
    args.serviceArgs.formulaInstances.upsert({
      cellIndex,
      sheetName,
      row: position.row,
      col: position.col,
      source,
      ...(templateId !== undefined ? { templateId } : {}),
    })
  }

  const registerFormulaFamilyNow = (cellIndex: number, formula: RuntimeFormula, ownerPosition?: FormulaOwnerPosition): void => {
    const sheetId = args.serviceArgs.state.workbook.cellStore.sheetIds[cellIndex]
    const position = ownerPosition ?? args.serviceArgs.state.workbook.getCellPosition(cellIndex)
    if (sheetId === undefined || !position || formula.templateId === undefined) {
      args.serviceArgs.formulaFamilies.unregisterFormula(cellIndex)
      return
    }
    if (args.serviceArgs.formulaFamilies.getMembership(cellIndex)) {
      return
    }
    args.serviceArgs.formulaFamilies.upsertFormula({
      cellIndex,
      sheetId,
      row: position.row,
      col: position.col,
      templateId: formula.templateId,
      shapeKey: getFormulaBindingFamilyShapeKey(args.formulaFamilyShapeKeyCache, formula),
    })
  }

  const rebuildFormulaInstancesNow = (): void => {
    args.serviceArgs.formulaInstances.clear()
    args.serviceArgs.state.formulas.forEach((formula, cellIndex) => {
      recordFormulaInstanceNow(cellIndex, formula.source, formula.templateId)
    })
  }

  return {
    recordFormulaInstanceNow,
    registerFormulaFamilyNow,
    rebuildFormulaInstancesNow,
  }
}
