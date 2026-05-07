import type {
  CellRangeRef,
  WorkbookAutoFilterSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookSheetProtectionSnapshot,
  WorkbookSortSnapshot,
} from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'
import type { PivotTableInput } from './runtime-state.js'
import type { WorkbookStore } from '../workbook-store.js'
import { canonicalWorkbookRangeRef } from '../workbook-range-records.js'

export function buildSetSheetProtectionOps(workbook: WorkbookStore, protection: WorkbookSheetProtectionSnapshot): EngineOp[] | null {
  const existing = workbook.getSheetProtection(protection.sheetName)
  const normalized: WorkbookSheetProtectionSnapshot = {
    sheetName: protection.sheetName,
    ...(protection.hideFormulas !== undefined ? { hideFormulas: protection.hideFormulas } : {}),
  }
  if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
    return null
  }
  return [{ kind: 'setSheetProtection', protection: normalized }]
}

export function buildSetFilterOps(workbook: WorkbookStore, sheetName: string, range: WorkbookAutoFilterSnapshot): EngineOp[] | null {
  if (workbook.getFilter(sheetName, range)) {
    return null
  }
  return [{ kind: 'setFilter', sheetName, range: structuredClone(range) }]
}

export function buildSetSortOps(
  workbook: WorkbookStore,
  sheetName: string,
  range: CellRangeRef,
  keys: WorkbookSortSnapshot['keys'],
): EngineOp[] | null {
  const existing = workbook.getSort(sheetName, range)
  const normalizedKeys = keys.map((key) => Object.assign({}, key))
  if (
    existing &&
    existing.keys.length === normalizedKeys.length &&
    existing.keys.every(
      (key, index) => key.keyAddress === normalizedKeys[index]?.keyAddress && key.direction === normalizedKeys[index]?.direction,
    )
  ) {
    return null
  }
  return [{ kind: 'setSort', sheetName, range: { ...range }, keys: normalizedKeys }]
}

export function buildSetDataValidationOps(workbook: WorkbookStore, validation: WorkbookDataValidationSnapshot): EngineOp[] | null {
  const existing = workbook.getDataValidation(validation.range.sheetName, validation.range)
  const normalized: WorkbookDataValidationSnapshot = {
    ...structuredClone(validation),
    range: canonicalWorkbookRangeRef(validation.range),
  }
  if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
    return null
  }
  return [{ kind: 'setDataValidation', validation: normalized }]
}

export function buildSetConditionalFormatOps(workbook: WorkbookStore, format: WorkbookConditionalFormatSnapshot): EngineOp[] | null {
  const normalized: WorkbookConditionalFormatSnapshot = {
    ...structuredClone(format),
    id: format.id.trim(),
    range: canonicalWorkbookRangeRef(format.range),
  }
  const existing = workbook.getConditionalFormat(normalized.id)
  if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
    return null
  }
  return [{ kind: 'upsertConditionalFormat', format: normalized }]
}

export function buildSetRangeProtectionOps(workbook: WorkbookStore, protection: WorkbookRangeProtectionSnapshot): EngineOp[] | null {
  const normalized: WorkbookRangeProtectionSnapshot = {
    id: protection.id.trim(),
    range: canonicalWorkbookRangeRef(protection.range),
    ...(protection.hideFormulas !== undefined ? { hideFormulas: protection.hideFormulas } : {}),
  }
  const existing = workbook.getRangeProtection(normalized.id)
  if (existing && JSON.stringify(existing) === JSON.stringify(normalized)) {
    return null
  }
  return [{ kind: 'upsertRangeProtection', protection: normalized }]
}

export function buildSetPivotTableOps(sheetName: string, address: string, definition: PivotTableInput): EngineOp[] {
  return [
    {
      kind: 'upsertPivotTable',
      name: definition.name.trim(),
      sheetName,
      address,
      source: { ...definition.source },
      groupBy: [...definition.groupBy],
      values: definition.values.map((value) => Object.assign({}, value)),
      rows: 1,
      cols: Math.max(definition.groupBy.length + definition.values.length, 1),
    },
  ]
}
