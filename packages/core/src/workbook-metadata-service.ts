import { Cause, Effect, Exit } from 'effect'
import { parseCellAddress } from '@bilig/formula'
import { canonicalWorkbookAddress, canonicalWorkbookRangeRef } from './workbook-range-records.js'
import {
  cloneChartRecord,
  cloneCommentThreadRecord,
  cloneConditionalFormatRecord,
  cloneDataValidationRecord,
  cloneDefinedNameRecord,
  cloneDefinedNameValue,
  cloneFilterRecord,
  cloneImageRecord,
  cloneMacroPayloadRecord,
  cloneMergeRangeRecord,
  cloneNoteRecord,
  clonePivotRecord,
  clonePropertyRecord,
  cloneRangeProtectionRecord,
  cloneSheetProtectionRecord,
  cloneShapeRecord,
  cloneSortKeyRecord,
  cloneSortRecord,
  cloneSpillRecord,
  cloneTableRecord,
  commentThreadKey,
  conditionalFormatKey,
  dataValidationKey,
  deleteRecordsBySheet,
  filterKey,
  mergeRangeKey,
  noteKey,
  rangeProtectionKey,
  rekeyRecords,
  sortKey,
  spillKey,
  tableKey,
} from './workbook-metadata-records.js'
import {
  createWorkbookMetadataRecord,
  chartKey,
  imageKey,
  macroPayloadKey,
  type WorkbookChartRecord,
  type WorkbookCommentThreadRecord,
  type WorkbookConditionalFormatRecord,
  compareDefinedNameRecords,
  definedNameKey,
  normalizeDefinedNameScope,
  type WorkbookImageRecord,
  type WorkbookMacroPayloadRecord,
  type WorkbookMergeRangeRecord,
  type WorkbookNoteRecord,
  type WorkbookRangeProtectionRecord,
  type WorkbookSheetProtectionRecord,
  pivotKey,
  shapeKey,
  type WorkbookDataValidationRecord,
  type WorkbookDefinedNameRecord,
  type WorkbookFilterRecord,
  type WorkbookFreezePaneRecord,
  type WorkbookMetadataRecord,
  type WorkbookPivotRecord,
  type WorkbookPropertyRecord,
  type WorkbookSortRecord,
  type WorkbookShapeRecord,
  type WorkbookSpillRecord,
  type WorkbookTableRecord,
} from './workbook-metadata-types.js'
import { WorkbookMetadataError, type WorkbookMetadataService } from './workbook-metadata-service-contract.js'
import { canonicalMergeRangeRef, isSingleCellMergeRange, rangeContainsAddress, rangesIntersect } from './workbook-merge-records.js'

export { WorkbookMetadataError, type WorkbookMetadataService } from './workbook-metadata-service-contract.js'

function metadataErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function renameDataValidationSourceSheet(
  record: WorkbookDataValidationRecord,
  oldSheetName: string,
  newSheetName: string,
): WorkbookDataValidationRecord {
  const cloned = cloneDataValidationRecord(record)
  if (cloned.rule.kind !== 'list' || !cloned.rule.source) {
    return cloned
  }
  switch (cloned.rule.source.kind) {
    case 'cell-ref':
    case 'range-ref':
      if (cloned.rule.source.sheetName === oldSheetName) {
        cloned.rule.source.sheetName = newSheetName
      }
      return cloned
    case 'named-range':
    case 'structured-ref':
      return cloned
  }
  return cloned
}

interface NormalizedMergeRangeRecord {
  readonly record: WorkbookMergeRangeRecord
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

function normalizeMergeRangeForOverlap(record: WorkbookMergeRangeRecord): NormalizedMergeRangeRecord {
  const start = parseCellAddress(record.startAddress, record.sheetName)
  const end = parseCellAddress(record.endAddress, record.sheetName)
  return {
    record,
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

function mergeRangesOverlap(left: NormalizedMergeRangeRecord, right: NormalizedMergeRangeRecord): boolean {
  return !(left.endRow < right.startRow || right.endRow < left.startRow || left.endCol < right.startCol || right.endCol < left.startCol)
}

function assertMergeRangesDoNotOverlap(ranges: readonly WorkbookMergeRangeRecord[]): void {
  const normalized = ranges.map(normalizeMergeRangeForOverlap).toSorted((left, right) => left.startRow - right.startRow)
  const active: NormalizedMergeRangeRecord[] = []
  for (const range of normalized) {
    for (let index = active.length - 1; index >= 0; index -= 1) {
      if (active[index]!.endRow < range.startRow) {
        active.splice(index, 1)
      }
    }
    if (active.some((entry) => mergeRangesOverlap(entry, range))) {
      throw new Error('Merged ranges cannot overlap')
    }
    active.push(range)
  }
}

function canonicalWorkbookFilterRange(range: WorkbookFilterRecord['range']): WorkbookFilterRecord['range'] {
  const normalized = canonicalWorkbookRangeRef(range)
  const criteria = range.criteria?.length ? structuredClone(range.criteria) : undefined
  return criteria ? { ...normalized, criteria } : normalized
}

function metadataEffect<Success>(message: string, run: () => Success): Effect.Effect<Success, WorkbookMetadataError> {
  return Effect.try({
    try: run,
    catch: (cause) =>
      new WorkbookMetadataError({
        message: metadataErrorMessage(message, cause),
        cause,
      }),
  })
}

export function createWorkbookMetadataService(metadata: WorkbookMetadataRecord): WorkbookMetadataService {
  const renameSheetNow = (oldSheetName: string, newSheetName: string): void => {
    rekeyRecords(metadata.freezePanes, (record) => (record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : record))
    rekeyRecords(metadata.merges, (record) =>
      record.sheetName === oldSheetName ? { ...cloneMergeRangeRecord(record), sheetName: newSheetName } : cloneMergeRangeRecord(record),
    )
    rekeyRecords(metadata.rowMetadata, (record) => (record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : record))
    rekeyRecords(metadata.columnMetadata, (record) => (record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : record))
    rekeyRecords(metadata.filters, (record) =>
      record.sheetName === oldSheetName || record.range.sheetName === oldSheetName
        ? {
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            range: {
              ...record.range,
              sheetName: record.range.sheetName === oldSheetName ? newSheetName : record.range.sheetName,
            },
          }
        : cloneFilterRecord(record),
    )
    rekeyRecords(metadata.sorts, (record) =>
      record.sheetName === oldSheetName || record.range.sheetName === oldSheetName
        ? {
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            range: {
              ...record.range,
              sheetName: record.range.sheetName === oldSheetName ? newSheetName : record.range.sheetName,
            },
            keys: record.keys.map(cloneSortKeyRecord),
          }
        : cloneSortRecord(record),
    )
    rekeyRecords(metadata.dataValidations, (record) => {
      const cloned = renameDataValidationSourceSheet(record, oldSheetName, newSheetName)
      if (cloned.range.sheetName === oldSheetName) {
        cloned.range.sheetName = newSheetName
      }
      return cloned
    })
    rekeyRecords(metadata.sheetProtections, (record) =>
      record.sheetName === oldSheetName
        ? { ...cloneSheetProtectionRecord(record), sheetName: newSheetName }
        : cloneSheetProtectionRecord(record),
    )
    rekeyRecords(metadata.conditionalFormats, (record) =>
      record.range.sheetName === oldSheetName
        ? {
            ...cloneConditionalFormatRecord(record),
            range: {
              ...record.range,
              sheetName: newSheetName,
            },
          }
        : cloneConditionalFormatRecord(record),
    )
    rekeyRecords(metadata.rangeProtections, (record) =>
      record.range.sheetName === oldSheetName
        ? {
            ...cloneRangeProtectionRecord(record),
            range: {
              ...record.range,
              sheetName: newSheetName,
            },
          }
        : cloneRangeProtectionRecord(record),
    )
    rekeyRecords(metadata.commentThreads, (record) =>
      record.sheetName === oldSheetName
        ? { ...cloneCommentThreadRecord(record), sheetName: newSheetName }
        : cloneCommentThreadRecord(record),
    )
    rekeyRecords(metadata.notes, (record) =>
      record.sheetName === oldSheetName ? { ...cloneNoteRecord(record), sheetName: newSheetName } : cloneNoteRecord(record),
    )
    rekeyRecords(metadata.tables, (record) =>
      record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : cloneTableRecord(record),
    )
    rekeyRecords(metadata.spills, (record) => (record.sheetName === oldSheetName ? { ...record, sheetName: newSheetName } : { ...record }))
    rekeyRecords(metadata.pivots, (record) =>
      record.sheetName === oldSheetName || record.source.sheetName === oldSheetName
        ? {
            ...record,
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            source: {
              ...record.source,
              sheetName: record.source.sheetName === oldSheetName ? newSheetName : record.source.sheetName,
            },
            groupBy: [...record.groupBy],
            values: record.values.map((value) => ({ ...value })),
          }
        : clonePivotRecord(record),
    )
    rekeyRecords(metadata.charts, (record) =>
      record.sheetName === oldSheetName || record.source.sheetName === oldSheetName
        ? {
            ...cloneChartRecord(record),
            sheetName: record.sheetName === oldSheetName ? newSheetName : record.sheetName,
            source: {
              ...record.source,
              sheetName: record.source.sheetName === oldSheetName ? newSheetName : record.source.sheetName,
            },
          }
        : cloneChartRecord(record),
    )
    rekeyRecords(metadata.images, (record) =>
      record.sheetName === oldSheetName ? { ...cloneImageRecord(record), sheetName: newSheetName } : cloneImageRecord(record),
    )
    rekeyRecords(metadata.shapes, (record) =>
      record.sheetName === oldSheetName ? { ...cloneShapeRecord(record), sheetName: newSheetName } : cloneShapeRecord(record),
    )
    rekeyRecords(metadata.definedNames, (record) =>
      record.scopeSheetName === oldSheetName
        ? { ...cloneDefinedNameRecord(record), scopeSheetName: newSheetName }
        : cloneDefinedNameRecord(record),
    )
  }

  const deleteSheetRecordsNow = (sheetName: string): void => {
    deleteRecordsBySheet(metadata.definedNames, sheetName, (record) => record.scopeSheetName)
    deleteRecordsBySheet(metadata.tables, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.spills, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.pivots, sheetName, (record) => record.sheetName)
    for (const [key, record] of metadata.charts.entries()) {
      if (record.sheetName === sheetName || record.source.sheetName === sheetName) {
        metadata.charts.delete(key)
      }
    }
    deleteRecordsBySheet(metadata.images, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.shapes, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.rowMetadata, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.columnMetadata, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.merges, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.filters, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.sorts, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.dataValidations, sheetName, (record) => record.range.sheetName)
    metadata.sheetProtections.delete(sheetName)
    deleteRecordsBySheet(metadata.conditionalFormats, sheetName, (record) => record.range.sheetName)
    deleteRecordsBySheet(metadata.rangeProtections, sheetName, (record) => record.range.sheetName)
    deleteRecordsBySheet(metadata.commentThreads, sheetName, (record) => record.sheetName)
    deleteRecordsBySheet(metadata.notes, sheetName, (record) => record.sheetName)
    metadata.freezePanes.delete(sheetName)
  }

  const resetNow = (): void => {
    const defaults = createWorkbookMetadataRecord()
    metadata.properties.clear()
    metadata.macroPayloads.clear()
    metadata.definedNames.clear()
    metadata.tables.clear()
    metadata.spills.clear()
    metadata.pivots.clear()
    metadata.charts.clear()
    metadata.images.clear()
    metadata.shapes.clear()
    metadata.rowMetadata.clear()
    metadata.columnMetadata.clear()
    metadata.freezePanes.clear()
    metadata.merges.clear()
    metadata.sheetProtections.clear()
    metadata.filters.clear()
    metadata.sorts.clear()
    metadata.dataValidations.clear()
    metadata.conditionalFormats.clear()
    metadata.rangeProtections.clear()
    metadata.commentThreads.clear()
    metadata.notes.clear()
    metadata.calculationSettings = defaults.calculationSettings
    metadata.volatileContext = defaults.volatileContext
  }

  return {
    renameSheet(oldSheetName, newSheetName) {
      return metadataEffect('Failed to rename workbook sheet metadata', () => renameSheetNow(oldSheetName, newSheetName))
    },
    deleteSheetRecords(sheetName) {
      return metadataEffect('Failed to delete workbook sheet metadata', () => deleteSheetRecordsNow(sheetName))
    },
    reset() {
      return metadataEffect('Failed to reset workbook metadata', resetNow)
    },
    setWorkbookProperty(key, value) {
      return metadataEffect('Failed to set workbook property', () => {
        const trimmedKey = normalizeMetadataKey(key)
        if (value === null) {
          metadata.properties.delete(trimmedKey)
          return undefined
        }
        const record: WorkbookPropertyRecord = { key: trimmedKey, value }
        metadata.properties.set(trimmedKey, record)
        return { ...record }
      })
    },
    getWorkbookProperty(key) {
      return metadataEffect('Failed to get workbook property', () => {
        const record = metadata.properties.get(normalizeMetadataKey(key))
        return record ? { ...record } : undefined
      })
    },
    listWorkbookProperties() {
      return metadataEffect('Failed to list workbook properties', () =>
        [...metadata.properties.values()].toSorted((left, right) => left.key.localeCompare(right.key)).map(clonePropertyRecord),
      )
    },
    setMacroPayload(record) {
      return metadataEffect('Failed to set macro payload metadata', () => {
        const stored: WorkbookMacroPayloadRecord = cloneMacroPayloadRecord(record)
        metadata.macroPayloads.set(macroPayloadKey(stored.kind), stored)
        return cloneMacroPayloadRecord(stored)
      })
    },
    listMacroPayloads() {
      return metadataEffect('Failed to list macro payload metadata', () =>
        [...metadata.macroPayloads.values()]
          .toSorted((left, right) => macroPayloadKey(left.kind).localeCompare(macroPayloadKey(right.kind)))
          .map(cloneMacroPayloadRecord),
      )
    },
    setCalculationSettings(settings) {
      return metadataEffect('Failed to set calculation settings', () => {
        metadata.calculationSettings = {
          compatibilityMode: 'excel-modern',
          ...settings,
        }
        return { ...metadata.calculationSettings }
      })
    },
    getCalculationSettings() {
      return metadataEffect('Failed to get calculation settings', () => ({ ...metadata.calculationSettings }))
    },
    setVolatileContext(context) {
      return metadataEffect('Failed to set volatile context', () => {
        metadata.volatileContext = { ...context }
        return { ...metadata.volatileContext }
      })
    },
    getVolatileContext() {
      return metadataEffect('Failed to get volatile context', () => ({ ...metadata.volatileContext }))
    },
    setDefinedName(name, value, scopeSheetName) {
      return metadataEffect('Failed to set defined name', () => {
        const trimmedName = name.trim()
        const normalizedScope = normalizeDefinedNameScope(scopeSheetName)
        const record: WorkbookDefinedNameRecord = {
          name: trimmedName,
          ...(normalizedScope !== undefined ? { scopeSheetName: normalizedScope } : {}),
          value: cloneDefinedNameValue(value),
        }
        metadata.definedNames.set(definedNameKey(trimmedName, normalizedScope), record)
        return cloneDefinedNameRecord(record)
      })
    },
    getDefinedName(name, scopeSheetName) {
      return metadataEffect('Failed to get defined name', () => {
        const scopedKey = definedNameKey(name, scopeSheetName)
        const record = metadata.definedNames.get(scopedKey) ?? metadata.definedNames.get(definedNameKey(name))
        return record ? cloneDefinedNameRecord(record) : undefined
      })
    },
    deleteDefinedName(name, scopeSheetName) {
      return metadataEffect('Failed to delete defined name', () => metadata.definedNames.delete(definedNameKey(name, scopeSheetName)))
    },
    listDefinedNames() {
      return metadataEffect('Failed to list defined names', () =>
        [...metadata.definedNames.values()].toSorted(compareDefinedNameRecords).map(cloneDefinedNameRecord),
      )
    },
    setTable(record) {
      return metadataEffect('Failed to set table metadata', () => {
        const stored: WorkbookTableRecord = {
          name: record.name.trim(),
          sheetName: record.sheetName,
          startAddress: record.startAddress,
          endAddress: record.endAddress,
          columnNames: [...record.columnNames],
          headerRow: record.headerRow,
          totalsRow: record.totalsRow,
        }
        metadata.tables.set(tableKey(stored.name), stored)
        return cloneTableRecord(stored)
      })
    },
    getTable(name) {
      return metadataEffect('Failed to get table metadata', () => {
        const record = metadata.tables.get(tableKey(name))
        return record ? cloneTableRecord(record) : undefined
      })
    },
    deleteTable(name) {
      return metadataEffect('Failed to delete table metadata', () => metadata.tables.delete(tableKey(name)))
    },
    listTables() {
      return metadataEffect('Failed to list table metadata', () =>
        [...metadata.tables.values()]
          .toSorted((left, right) => tableKey(left.name).localeCompare(tableKey(right.name)))
          .map(cloneTableRecord),
      )
    },
    setFreezePane(sheetName, rows, cols, options) {
      return metadataEffect('Failed to set freeze pane metadata', () => {
        const record: WorkbookFreezePaneRecord = { sheetName, rows, cols }
        if (options?.topLeftCell !== undefined) {
          record.topLeftCell = options.topLeftCell
        }
        if (options?.activePane !== undefined) {
          record.activePane = options.activePane
        }
        metadata.freezePanes.set(sheetName, record)
        return { ...record }
      })
    },
    getFreezePane(sheetName) {
      return metadataEffect('Failed to get freeze pane metadata', () => {
        const record = metadata.freezePanes.get(sheetName)
        return record ? { ...record } : undefined
      })
    },
    clearFreezePane(sheetName) {
      return metadataEffect('Failed to clear freeze pane metadata', () => metadata.freezePanes.delete(sheetName))
    },
    setMergeRange(range) {
      return metadataEffect('Failed to set merged cell metadata', () => {
        const stored = canonicalMergeRangeRef(range)
        if (isSingleCellMergeRange(stored)) {
          throw new Error('Merged ranges must include at least two cells')
        }
        const overlapping = [...metadata.merges.values()].filter((record) => rangesIntersect(record, stored))
        if (overlapping.some((record) => mergeRangeKey(record) !== mergeRangeKey(stored))) {
          throw new Error('Merged ranges cannot overlap')
        }
        metadata.merges.set(mergeRangeKey(stored), stored)
        return cloneMergeRangeRecord(stored)
      })
    },
    setMergeRanges(sheetName, ranges) {
      return metadataEffect('Failed to set merged cell metadata ranges', () => {
        const storedRanges = ranges.map((range) => canonicalMergeRangeRef({ ...range, sheetName: range.sheetName ?? sheetName }))
        const seenKeys = new Set<string>()
        for (const stored of storedRanges) {
          if (isSingleCellMergeRange(stored)) {
            throw new Error('Merged ranges must include at least two cells')
          }
          const key = mergeRangeKey(stored)
          if (seenKeys.has(key)) {
            throw new Error('Merged ranges cannot contain duplicate ranges')
          }
          seenKeys.add(key)
        }
        assertMergeRangesDoNotOverlap(storedRanges)
        deleteRecordsBySheet(metadata.merges, sheetName, (record) => record.sheetName)
        for (const stored of storedRanges) {
          metadata.merges.set(mergeRangeKey(stored), stored)
        }
        return storedRanges.toSorted((left, right) => mergeRangeKey(left).localeCompare(mergeRangeKey(right))).map(cloneMergeRangeRecord)
      })
    },
    getMergeRange(sheetName, address) {
      return metadataEffect('Failed to get merged cell metadata', () => {
        const record = [...metadata.merges.values()].find((entry) => rangeContainsAddress(entry, sheetName, address))
        return record ? cloneMergeRangeRecord(record) : undefined
      })
    },
    getMergeRangeByRange(range) {
      return metadataEffect('Failed to get merged range metadata', () => {
        const record = metadata.merges.get(mergeRangeKey(range))
        return record ? cloneMergeRangeRecord(record) : undefined
      })
    },
    clearMergeRanges(range) {
      return metadataEffect('Failed to clear merged cell metadata', () => {
        const removed: WorkbookMergeRangeRecord[] = []
        for (const [key, record] of metadata.merges.entries()) {
          if (!rangesIntersect(record, range)) {
            continue
          }
          metadata.merges.delete(key)
          removed.push(cloneMergeRangeRecord(record))
        }
        return removed.toSorted((left, right) => mergeRangeKey(left).localeCompare(mergeRangeKey(right)))
      })
    },
    listMergeRanges(sheetName) {
      return metadataEffect('Failed to list merged cell metadata', () =>
        [...metadata.merges.values()]
          .filter((record) => record.sheetName === sheetName)
          .toSorted((left, right) => mergeRangeKey(left).localeCompare(mergeRangeKey(right)))
          .map(cloneMergeRangeRecord),
      )
    },
    setSheetProtection(record) {
      return metadataEffect('Failed to set sheet protection metadata', () => {
        const stored: WorkbookSheetProtectionRecord = cloneSheetProtectionRecord({
          sheetName: record.sheetName,
          ...(record.hideFormulas !== undefined ? { hideFormulas: record.hideFormulas } : {}),
        })
        metadata.sheetProtections.set(record.sheetName, stored)
        return cloneSheetProtectionRecord(stored)
      })
    },
    getSheetProtection(sheetName) {
      return metadataEffect('Failed to get sheet protection metadata', () => {
        const record = metadata.sheetProtections.get(sheetName)
        return record ? cloneSheetProtectionRecord(record) : undefined
      })
    },
    clearSheetProtection(sheetName) {
      return metadataEffect('Failed to clear sheet protection metadata', () => metadata.sheetProtections.delete(sheetName))
    },
    setFilter(sheetName, range) {
      return metadataEffect('Failed to set filter metadata', () => {
        const storedRange = canonicalWorkbookFilterRange(range)
        const record: WorkbookFilterRecord = { sheetName, range: storedRange }
        metadata.filters.set(filterKey(sheetName, storedRange), record)
        return cloneFilterRecord(record)
      })
    },
    getFilter(sheetName, range) {
      return metadataEffect('Failed to get filter metadata', () => {
        const record = metadata.filters.get(filterKey(sheetName, range))
        return record ? cloneFilterRecord(record) : undefined
      })
    },
    deleteFilter(sheetName, range) {
      return metadataEffect('Failed to delete filter metadata', () => metadata.filters.delete(filterKey(sheetName, range)))
    },
    listFilters(sheetName) {
      return metadataEffect('Failed to list filter metadata', () =>
        [...metadata.filters.values()]
          .filter((record) => record.sheetName === sheetName)
          .toSorted((left, right) => filterKey(left.sheetName, left.range).localeCompare(filterKey(right.sheetName, right.range)))
          .map(cloneFilterRecord),
      )
    },
    setSort(sheetName, range, keys) {
      return metadataEffect('Failed to set sort metadata', () => {
        const storedRange = canonicalWorkbookRangeRef(range)
        const record: WorkbookSortRecord = {
          sheetName,
          range: storedRange,
          keys: keys.map(cloneSortKeyRecord),
        }
        metadata.sorts.set(sortKey(sheetName, storedRange), record)
        return cloneSortRecord(record)
      })
    },
    getSort(sheetName, range) {
      return metadataEffect('Failed to get sort metadata', () => {
        const record = metadata.sorts.get(sortKey(sheetName, range))
        return record ? cloneSortRecord(record) : undefined
      })
    },
    deleteSort(sheetName, range) {
      return metadataEffect('Failed to delete sort metadata', () => metadata.sorts.delete(sortKey(sheetName, range)))
    },
    listSorts(sheetName) {
      return metadataEffect('Failed to list sort metadata', () =>
        [...metadata.sorts.values()]
          .filter((record) => record.sheetName === sheetName)
          .toSorted((left, right) => sortKey(left.sheetName, left.range).localeCompare(sortKey(right.sheetName, right.range)))
          .map(cloneSortRecord),
      )
    },
    setDataValidation(record) {
      return metadataEffect('Failed to set data validation metadata', () => {
        const storedRange = canonicalWorkbookRangeRef(record.range)
        const nextRecord: WorkbookDataValidationRecord = {
          range: storedRange,
          rule: record.rule,
        }
        if (record.allowBlank !== undefined) {
          nextRecord.allowBlank = record.allowBlank
        }
        if (record.showDropdown !== undefined) {
          nextRecord.showDropdown = record.showDropdown
        }
        if (record.promptTitle !== undefined) {
          nextRecord.promptTitle = record.promptTitle
        }
        if (record.promptMessage !== undefined) {
          nextRecord.promptMessage = record.promptMessage
        }
        if (record.errorStyle !== undefined) {
          nextRecord.errorStyle = record.errorStyle
        }
        if (record.errorTitle !== undefined) {
          nextRecord.errorTitle = record.errorTitle
        }
        if (record.errorMessage !== undefined) {
          nextRecord.errorMessage = record.errorMessage
        }
        const stored: WorkbookDataValidationRecord = cloneDataValidationRecord(nextRecord)
        metadata.dataValidations.set(dataValidationKey(storedRange.sheetName, storedRange), stored)
        return cloneDataValidationRecord(stored)
      })
    },
    getDataValidation(sheetName, range) {
      return metadataEffect('Failed to get data validation metadata', () => {
        const record = metadata.dataValidations.get(dataValidationKey(sheetName, range))
        return record ? cloneDataValidationRecord(record) : undefined
      })
    },
    deleteDataValidation(sheetName, range) {
      return metadataEffect('Failed to delete data validation metadata', () =>
        metadata.dataValidations.delete(dataValidationKey(sheetName, range)),
      )
    },
    listDataValidations(sheetName) {
      return metadataEffect('Failed to list data validation metadata', () =>
        [...metadata.dataValidations.values()]
          .filter((record) => record.range.sheetName === sheetName)
          .toSorted((left, right) =>
            dataValidationKey(left.range.sheetName, left.range).localeCompare(dataValidationKey(right.range.sheetName, right.range)),
          )
          .map(cloneDataValidationRecord),
      )
    },
    setConditionalFormat(record) {
      return metadataEffect('Failed to set conditional format metadata', () => {
        const id = conditionalFormatKey(record.id)
        const nextRecord: WorkbookConditionalFormatRecord = {
          id,
          range: canonicalWorkbookRangeRef(record.range),
          rule: structuredClone(record.rule),
          style: structuredClone(record.style),
        }
        if (record.stopIfTrue !== undefined) {
          nextRecord.stopIfTrue = record.stopIfTrue
        }
        if (record.priority !== undefined) {
          nextRecord.priority = record.priority
        }
        const stored = cloneConditionalFormatRecord(nextRecord)
        metadata.conditionalFormats.set(id, stored)
        return cloneConditionalFormatRecord(stored)
      })
    },
    getConditionalFormat(id) {
      return metadataEffect('Failed to get conditional format metadata', () => {
        const record = metadata.conditionalFormats.get(conditionalFormatKey(id))
        return record ? cloneConditionalFormatRecord(record) : undefined
      })
    },
    deleteConditionalFormat(id) {
      return metadataEffect('Failed to delete conditional format metadata', () =>
        metadata.conditionalFormats.delete(conditionalFormatKey(id)),
      )
    },
    listConditionalFormats(sheetName) {
      return metadataEffect('Failed to list conditional format metadata', () =>
        [...metadata.conditionalFormats.values()]
          .filter((record) => record.range.sheetName === sheetName)
          .toSorted((left, right) => {
            const priorityCompare = (left.priority ?? Number.MAX_SAFE_INTEGER) - (right.priority ?? Number.MAX_SAFE_INTEGER)
            if (priorityCompare !== 0) {
              return priorityCompare
            }
            return left.id.localeCompare(right.id)
          })
          .map(cloneConditionalFormatRecord),
      )
    },
    setRangeProtection(record) {
      return metadataEffect('Failed to set range protection metadata', () => {
        const id = rangeProtectionKey(record.id)
        const stored: WorkbookRangeProtectionRecord = cloneRangeProtectionRecord({
          id,
          range: canonicalWorkbookRangeRef(record.range),
          ...(record.hideFormulas !== undefined ? { hideFormulas: record.hideFormulas } : {}),
        })
        metadata.rangeProtections.set(id, stored)
        return cloneRangeProtectionRecord(stored)
      })
    },
    getRangeProtection(id) {
      return metadataEffect('Failed to get range protection metadata', () => {
        const record = metadata.rangeProtections.get(rangeProtectionKey(id))
        return record ? cloneRangeProtectionRecord(record) : undefined
      })
    },
    deleteRangeProtection(id) {
      return metadataEffect('Failed to delete range protection metadata', () => metadata.rangeProtections.delete(rangeProtectionKey(id)))
    },
    listRangeProtections(sheetName) {
      return metadataEffect('Failed to list range protection metadata', () =>
        [...metadata.rangeProtections.values()]
          .filter((record) => record.range.sheetName === sheetName)
          .toSorted((left, right) => left.id.localeCompare(right.id))
          .map(cloneRangeProtectionRecord),
      )
    },
    setCommentThread(record) {
      return metadataEffect('Failed to set comment thread metadata', () => {
        const normalizedAddress = canonicalWorkbookAddress(record.sheetName, record.address)
        const stored: WorkbookCommentThreadRecord = cloneCommentThreadRecord({
          ...record,
          threadId: record.threadId.trim(),
          address: normalizedAddress,
          comments: record.comments.map((comment) => ({
            id: comment.id.trim(),
            body: comment.body.trim(),
            ...(comment.authorUserId !== undefined ? { authorUserId: comment.authorUserId } : {}),
            ...(comment.authorDisplayName !== undefined ? { authorDisplayName: comment.authorDisplayName } : {}),
            ...(comment.createdAtUnixMs !== undefined ? { createdAtUnixMs: comment.createdAtUnixMs } : {}),
          })),
        })
        metadata.commentThreads.set(commentThreadKey(record.sheetName, normalizedAddress), stored)
        return cloneCommentThreadRecord(stored)
      })
    },
    getCommentThread(sheetName, address) {
      return metadataEffect('Failed to get comment thread metadata', () => {
        const record = metadata.commentThreads.get(commentThreadKey(sheetName, address))
        return record ? cloneCommentThreadRecord(record) : undefined
      })
    },
    deleteCommentThread(sheetName, address) {
      return metadataEffect('Failed to delete comment thread metadata', () =>
        metadata.commentThreads.delete(commentThreadKey(sheetName, address)),
      )
    },
    listCommentThreads(sheetName) {
      return metadataEffect('Failed to list comment thread metadata', () =>
        [...metadata.commentThreads.values()]
          .filter((record) => record.sheetName === sheetName)
          .toSorted((left, right) =>
            commentThreadKey(left.sheetName, left.address).localeCompare(commentThreadKey(right.sheetName, right.address)),
          )
          .map(cloneCommentThreadRecord),
      )
    },
    setNote(record) {
      return metadataEffect('Failed to set note metadata', () => {
        const normalizedAddress = canonicalWorkbookAddress(record.sheetName, record.address)
        const stored: WorkbookNoteRecord = cloneNoteRecord({
          sheetName: record.sheetName,
          address: normalizedAddress,
          text: record.text.trim(),
        })
        metadata.notes.set(noteKey(record.sheetName, normalizedAddress), stored)
        return cloneNoteRecord(stored)
      })
    },
    getNote(sheetName, address) {
      return metadataEffect('Failed to get note metadata', () => {
        const record = metadata.notes.get(noteKey(sheetName, address))
        return record ? cloneNoteRecord(record) : undefined
      })
    },
    deleteNote(sheetName, address) {
      return metadataEffect('Failed to delete note metadata', () => metadata.notes.delete(noteKey(sheetName, address)))
    },
    listNotes(sheetName) {
      return metadataEffect('Failed to list note metadata', () =>
        [...metadata.notes.values()]
          .filter((record) => record.sheetName === sheetName)
          .toSorted((left, right) => noteKey(left.sheetName, left.address).localeCompare(noteKey(right.sheetName, right.address)))
          .map(cloneNoteRecord),
      )
    },
    setSpill(sheetName, address, rows, cols) {
      return metadataEffect('Failed to set spill metadata', () => {
        const normalizedAddress = canonicalWorkbookAddress(sheetName, address)
        const record: WorkbookSpillRecord = { sheetName, address: normalizedAddress, rows, cols }
        metadata.spills.set(spillKey(sheetName, normalizedAddress), record)
        return { ...record }
      })
    },
    getSpill(sheetName, address) {
      return metadataEffect('Failed to get spill metadata', () => {
        const record = metadata.spills.get(spillKey(sheetName, address))
        return record ? { ...record } : undefined
      })
    },
    deleteSpill(sheetName, address) {
      return metadataEffect('Failed to delete spill metadata', () => metadata.spills.delete(spillKey(sheetName, address)))
    },
    listSpills() {
      return metadataEffect('Failed to list spill metadata', () =>
        [...metadata.spills.values()]
          .toSorted((left, right) => `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`))
          .map(cloneSpillRecord),
      )
    },
    setPivot(record) {
      return metadataEffect('Failed to set pivot metadata', () => {
        const normalizedAddress = canonicalWorkbookAddress(record.sheetName, record.address)
        const stored: WorkbookPivotRecord = {
          ...record,
          name: record.name.trim(),
          address: normalizedAddress,
          groupBy: [...record.groupBy],
          values: record.values.map((value) => ({ ...value })),
          source: canonicalWorkbookRangeRef(record.source),
        }
        metadata.pivots.set(pivotKey(record.sheetName, normalizedAddress), stored)
        return clonePivotRecord(stored)
      })
    },
    getPivot(sheetName, address) {
      return metadataEffect('Failed to get pivot metadata', () => {
        const record = metadata.pivots.get(pivotKey(sheetName, address))
        return record ? clonePivotRecord(record) : undefined
      })
    },
    getPivotByKey(key) {
      return metadataEffect('Failed to get pivot metadata by key', () => {
        const record = metadata.pivots.get(key)
        return record ? clonePivotRecord(record) : undefined
      })
    },
    deletePivot(sheetName, address) {
      return metadataEffect('Failed to delete pivot metadata', () => metadata.pivots.delete(pivotKey(sheetName, address)))
    },
    hasPivots() {
      return metadataEffect('Failed to read pivot metadata state', () => metadata.pivots.size > 0)
    },
    listPivots() {
      return metadataEffect('Failed to list pivot metadata', () =>
        [...metadata.pivots.values()]
          .toSorted((left, right) => `${left.sheetName}!${left.address}`.localeCompare(`${right.sheetName}!${right.address}`))
          .map(clonePivotRecord),
      )
    },
    setChart(record) {
      return metadataEffect('Failed to set chart metadata', () => {
        const stored: WorkbookChartRecord = {
          id: record.id.trim(),
          sheetName: record.sheetName,
          address: canonicalWorkbookAddress(record.sheetName, record.address),
          source: canonicalWorkbookRangeRef(record.source),
          chartType: record.chartType,
          rows: record.rows,
          cols: record.cols,
          ...(record.seriesOrientation !== undefined ? { seriesOrientation: record.seriesOrientation } : {}),
          ...(record.firstRowAsHeaders !== undefined ? { firstRowAsHeaders: record.firstRowAsHeaders } : {}),
          ...(record.firstColumnAsLabels !== undefined ? { firstColumnAsLabels: record.firstColumnAsLabels } : {}),
          ...(record.title !== undefined ? { title: record.title } : {}),
          ...(record.legendPosition !== undefined ? { legendPosition: record.legendPosition } : {}),
        }
        metadata.charts.set(chartKey(stored.id), stored)
        return cloneChartRecord(stored)
      })
    },
    getChart(id) {
      return metadataEffect('Failed to get chart metadata', () => {
        const record = metadata.charts.get(chartKey(id))
        return record ? cloneChartRecord(record) : undefined
      })
    },
    deleteChart(id) {
      return metadataEffect('Failed to delete chart metadata', () => metadata.charts.delete(chartKey(id)))
    },
    listCharts() {
      return metadataEffect('Failed to list chart metadata', () =>
        [...metadata.charts.values()].toSorted((left, right) => left.id.localeCompare(right.id)).map(cloneChartRecord),
      )
    },
    setImage(record) {
      return metadataEffect('Failed to set image metadata', () => {
        const stored: WorkbookImageRecord = {
          id: record.id.trim(),
          sheetName: record.sheetName,
          address: canonicalWorkbookAddress(record.sheetName, record.address),
          sourceUrl: record.sourceUrl,
          rows: record.rows,
          cols: record.cols,
          ...(record.altText !== undefined ? { altText: record.altText } : {}),
        }
        metadata.images.set(imageKey(stored.id), stored)
        return cloneImageRecord(stored)
      })
    },
    getImage(id) {
      return metadataEffect('Failed to get image metadata', () => {
        const record = metadata.images.get(imageKey(id))
        return record ? cloneImageRecord(record) : undefined
      })
    },
    deleteImage(id) {
      return metadataEffect('Failed to delete image metadata', () => metadata.images.delete(imageKey(id)))
    },
    listImages() {
      return metadataEffect('Failed to list image metadata', () =>
        [...metadata.images.values()].toSorted((left, right) => left.id.localeCompare(right.id)).map(cloneImageRecord),
      )
    },
    setShape(record) {
      return metadataEffect('Failed to set shape metadata', () => {
        const stored: WorkbookShapeRecord = {
          id: record.id.trim(),
          sheetName: record.sheetName,
          address: canonicalWorkbookAddress(record.sheetName, record.address),
          shapeType: record.shapeType,
          rows: record.rows,
          cols: record.cols,
          ...(record.text !== undefined ? { text: record.text } : {}),
          ...(record.fillColor !== undefined ? { fillColor: record.fillColor } : {}),
          ...(record.strokeColor !== undefined ? { strokeColor: record.strokeColor } : {}),
        }
        metadata.shapes.set(shapeKey(stored.id), stored)
        return cloneShapeRecord(stored)
      })
    },
    getShape(id) {
      return metadataEffect('Failed to get shape metadata', () => {
        const record = metadata.shapes.get(shapeKey(id))
        return record ? cloneShapeRecord(record) : undefined
      })
    },
    deleteShape(id) {
      return metadataEffect('Failed to delete shape metadata', () => metadata.shapes.delete(shapeKey(id)))
    },
    listShapes() {
      return metadataEffect('Failed to list shape metadata', () =>
        [...metadata.shapes.values()].toSorted((left, right) => left.id.localeCompare(right.id)).map(cloneShapeRecord),
      )
    },
  }
}

export function runWorkbookMetadataEffect<Success, Failure>(effect: Effect.Effect<Success, Failure>): Success {
  const exit = Effect.runSyncExit(effect)
  if (Exit.isSuccess(exit)) {
    return exit.value
  }
  throw Cause.squash(exit.cause)
}

function normalizeMetadataKey(key: string): string {
  const trimmed = key.trim()
  if (trimmed.length === 0) {
    throw new Error('Workbook metadata keys must be non-empty')
  }
  return trimmed
}
