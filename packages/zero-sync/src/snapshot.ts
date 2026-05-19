import { formatAddress, parseCellAddress } from '@bilig/formula'
import type {
  CellNumberFormatRecord,
  CellStyleRecord,
  CompatibilityMode,
  SheetFormatRangeSnapshot,
  SheetMetadataSnapshot,
  SheetStyleRangeSnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookDefinedNameSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookMetadataSnapshot,
  WorkbookPropertySnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { isWorkbookSnapshot, sanitizeCellStyleRecord } from '@bilig/protocol'
import { isLiteralInput } from './mutators.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined
}

function firstString(record: Record<string, unknown>, keys: readonly string[]): string | undefined {
  for (const key of keys) {
    const value = asString(record[key])
    if (value !== undefined) {
      return value
    }
  }
  return undefined
}

function asNonNegativeNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0 ? value : undefined
}

function asSafeNonNegativeInteger(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 ? value : undefined
}

function asSafePositiveInteger(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0 ? value : undefined
}

function firstSafeNonNegativeInteger(record: Record<string, unknown>, keys: readonly string[]): number | undefined {
  for (const key of keys) {
    const value = asSafeNonNegativeInteger(record[key])
    if (value !== undefined) {
      return value
    }
  }
  return undefined
}

function asBoolean(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

function asArray(value: unknown): unknown[] {
  return Array.isArray(value) ? value : []
}

function hasOwn(record: Record<string, unknown>, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(record, key)
}

function firstRecord(record: Record<string, unknown>, keys: readonly string[]): Record<string, unknown> | undefined {
  for (const key of keys) {
    const value = record[key]
    if (isRecord(value)) {
      return value
    }
  }
  return undefined
}

interface ArrayProjectionField<T> {
  readonly present: boolean
  readonly values: readonly T[]
}

interface OptionalProjectionField<T> {
  readonly present: boolean
  readonly value: T | undefined
}

function parseArrayProjectionField<T>(
  record: Record<string, unknown>,
  key: string,
  parse: (entries: unknown[]) => readonly T[],
): ArrayProjectionField<T> {
  const present = hasOwn(record, key)
  return {
    present,
    values: present ? parse(asArray(record[key])) : [],
  }
}

function isCellNumberFormatKind(value: unknown): value is CellNumberFormatRecord['kind'] {
  return (
    value === 'general' ||
    value === 'number' ||
    value === 'currency' ||
    value === 'accounting' ||
    value === 'percent' ||
    value === 'date' ||
    value === 'time' ||
    value === 'datetime' ||
    value === 'text'
  )
}

function isCompatibilityMode(value: unknown): value is CompatibilityMode {
  return value === 'excel-modern' || value === 'odf-1.4'
}

export function createEmptyWorkbookSnapshot(documentId: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: documentId,
    },
    sheets: [
      {
        id: 1,
        name: 'Sheet1',
        order: 0,
        cells: [],
      },
    ],
  }
}

function parseAxisMetadata(entries: unknown[]): WorkbookAxisMetadataSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const start = asSafeNonNegativeInteger(entry['startIndex'])
      const count = asSafePositiveInteger(entry['count'])
      if (start === undefined || count === undefined) {
        return null
      }
      const next: WorkbookAxisMetadataSnapshot = {
        start,
        count,
      }
      const size = asNonNegativeNumber(entry['size'])
      const hiddenFlag = asBoolean(entry['hidden'])
      if (size !== undefined) {
        next.size = size
      }
      if (hiddenFlag !== undefined) {
        next.hidden = hiddenFlag
      }
      return next
    })
    .filter((entry): entry is WorkbookAxisMetadataSnapshot => entry !== null)
}

function mergeAxisMetadataEntries(
  primary: readonly WorkbookAxisMetadataSnapshot[],
  fallback: readonly WorkbookAxisMetadataSnapshot[] | undefined,
): WorkbookAxisMetadataSnapshot[] {
  if (!fallback || fallback.length === 0) {
    return [...primary]
  }
  const fallbackByKey = new Map(fallback.map((entry) => [`${String(entry.start)}:${String(entry.count)}`, entry]))
  return primary.map((entry) => {
    const preserved = fallbackByKey.get(`${String(entry.start)}:${String(entry.count)}`)
    if (!preserved) {
      return entry
    }
    return {
      ...preserved,
      ...entry,
    }
  })
}

function mergeProjectedArray<T>(projected: ArrayProjectionField<T>, fallback: readonly T[] | undefined): readonly T[] | undefined {
  if (projected.present) {
    return projected.values.length > 0 ? projected.values : undefined
  }
  return fallback && fallback.length > 0 ? fallback : undefined
}

function parseWorkbookProperties(entries: unknown[]): WorkbookPropertySnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const key = asString(entry['key'])
      const value = entry['value']
      if (!key || !isLiteralInput(value)) {
        return null
      }
      return { key, value }
    })
    .filter((entry): entry is WorkbookPropertySnapshot => entry !== null)
}

function isWorkbookDefinedNameValueSnapshot(value: unknown): value is WorkbookDefinedNameValueSnapshot {
  if (isLiteralInput(value)) {
    return true
  }
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    return false
  }
  switch (value['kind']) {
    case 'scalar':
      return isLiteralInput(value['value'])
    case 'cell-ref':
      return typeof value['sheetName'] === 'string' && typeof value['address'] === 'string'
    case 'range-ref':
      return typeof value['sheetName'] === 'string' && typeof value['startAddress'] === 'string' && typeof value['endAddress'] === 'string'
    case 'structured-ref':
      return typeof value['tableName'] === 'string' && typeof value['columnName'] === 'string'
    case 'formula':
      return typeof value['formula'] === 'string'
    default:
      return false
  }
}

function parseDefinedNames(entries: unknown[]): WorkbookDefinedNameSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const name = asString(entry['name'])
      const value = entry['value']
      if (!name || !isWorkbookDefinedNameValueSnapshot(value)) {
        return null
      }
      return { name, value }
    })
    .filter((entry): entry is WorkbookDefinedNameSnapshot => entry !== null)
}

function parseStyleRecords(entries: unknown[]): CellStyleRecord[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const id = firstString(entry, ['id', 'styleId'])
      const recordJSON = firstRecord(entry, ['recordJSON', 'styleJson'])
      if (!id || !recordJSON) {
        return null
      }
      return sanitizeCellStyleRecord(id, recordJSON)
    })
    .filter((entry): entry is CellStyleRecord => entry !== null)
}

function parseNumberFormats(entries: unknown[]): CellNumberFormatRecord[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const id = firstString(entry, ['id', 'formatId'])
      const code = asString(entry['code'])
      const kind = asString(entry['kind'])
      if (!id || !code || !isCellNumberFormatKind(kind)) {
        return null
      }
      return {
        id,
        code,
        kind,
      }
    })
    .filter((entry): entry is CellNumberFormatRecord => entry !== null)
}

function parseFreezePane(
  sheetEntry: Record<string, unknown>,
  fallback?: WorkbookFreezePaneSnapshot,
): OptionalProjectionField<WorkbookFreezePaneSnapshot> {
  const rowsPresent = hasOwn(sheetEntry, 'freezeRows')
  const colsPresent = hasOwn(sheetEntry, 'freezeCols')
  if (!rowsPresent && !colsPresent) {
    return { present: false, value: fallback }
  }
  const rows = rowsPresent ? asSafeNonNegativeInteger(sheetEntry['freezeRows']) : 0
  const cols = colsPresent ? asSafeNonNegativeInteger(sheetEntry['freezeCols']) : 0
  if (rows === undefined || cols === undefined) {
    return { present: true, value: undefined }
  }
  if (rows > 0 || cols > 0) {
    return {
      present: true,
      value: {
        rows,
        cols,
      },
    }
  }
  return { present: true, value: undefined }
}

function preserveSnapshotOnlyWorkbookMetadata(metadata: WorkbookMetadataSnapshot | undefined): WorkbookMetadataSnapshot {
  if (!metadata) {
    return {}
  }
  const {
    properties: _properties,
    definedNames: _definedNames,
    styles: _styles,
    formats: _formats,
    calculationSettings: _calculationSettings,
    volatileContext: _volatileContext,
    ...snapshotOnlyMetadata
  } = metadata
  return { ...snapshotOnlyMetadata }
}

function preserveSnapshotOnlySheetMetadata(metadata: SheetMetadataSnapshot | undefined): SheetMetadataSnapshot {
  if (!metadata) {
    return {}
  }
  const {
    rowMetadata: _rowMetadata,
    columnMetadata: _columnMetadata,
    styleRanges: _styleRanges,
    formatRanges: _formatRanges,
    freezePane: _freezePane,
    ...snapshotOnlyMetadata
  } = metadata
  return { ...snapshotOnlyMetadata }
}

function parseStyleRanges(entries: unknown[]): SheetStyleRangeSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const startRow = asSafeNonNegativeInteger(entry['startRow'])
      const endRow = asSafeNonNegativeInteger(entry['endRow'])
      const startCol = asSafeNonNegativeInteger(entry['startCol'])
      const endCol = asSafeNonNegativeInteger(entry['endCol'])
      const styleId = asString(entry['styleId'])
      if (
        startRow === undefined ||
        endRow === undefined ||
        startCol === undefined ||
        endCol === undefined ||
        endRow < startRow ||
        endCol < startCol ||
        !styleId
      ) {
        return null
      }
      return {
        range: {
          sheetName: '',
          startAddress: formatAddress(startRow, startCol),
          endAddress: formatAddress(endRow, endCol),
        },
        styleId,
      }
    })
    .filter((entry): entry is SheetStyleRangeSnapshot => entry !== null)
}

function parseFormatRanges(entries: unknown[]): SheetFormatRangeSnapshot[] {
  return entries
    .map((entry) => {
      if (!isRecord(entry)) {
        return null
      }
      const startRow = asSafeNonNegativeInteger(entry['startRow'])
      const endRow = asSafeNonNegativeInteger(entry['endRow'])
      const startCol = asSafeNonNegativeInteger(entry['startCol'])
      const endCol = asSafeNonNegativeInteger(entry['endCol'])
      const formatId = asString(entry['formatId'])
      if (
        startRow === undefined ||
        endRow === undefined ||
        startCol === undefined ||
        endCol === undefined ||
        endRow < startRow ||
        endCol < startCol ||
        !formatId
      ) {
        return null
      }
      return {
        range: {
          sheetName: '',
          startAddress: formatAddress(startRow, startCol),
          endAddress: formatAddress(endRow, endCol),
        },
        formatId,
      }
    })
    .filter((entry): entry is SheetFormatRangeSnapshot => entry !== null)
}

function parseCellCoordinates(
  cellEntry: Record<string, unknown>,
  sheetName: string,
): { address: string; rowNum: number | undefined; colNum: number | undefined } | null {
  const rowNum = asSafeNonNegativeInteger(cellEntry['rowNum'])
  const colNum = asSafeNonNegativeInteger(cellEntry['colNum'])
  const address =
    asString(cellEntry['address']) ?? (rowNum !== undefined && colNum !== undefined ? formatAddress(rowNum, colNum) : undefined)
  if (!address) {
    return null
  }
  if (rowNum !== undefined && colNum !== undefined) {
    return { address, rowNum, colNum }
  }
  try {
    const parsed = parseCellAddress(address, sheetName)
    return { address, rowNum: parsed.row, colNum: parsed.col }
  } catch {
    return { address, rowNum: undefined, colNum: undefined }
  }
}

function singleCellStyleRange(
  rowNum: number | undefined,
  colNum: number | undefined,
  styleId: string | undefined,
): SheetStyleRangeSnapshot | null {
  if (rowNum === undefined || colNum === undefined || !styleId) {
    return null
  }
  const address = formatAddress(rowNum, colNum)
  return {
    range: {
      sheetName: '',
      startAddress: address,
      endAddress: address,
    },
    styleId,
  }
}

function singleCellFormatRange(
  rowNum: number | undefined,
  colNum: number | undefined,
  formatId: string | undefined,
): SheetFormatRangeSnapshot | null {
  if (rowNum === undefined || colNum === undefined || !formatId) {
    return null
  }
  const address = formatAddress(rowNum, colNum)
  return {
    range: {
      sheetName: '',
      startAddress: address,
      endAddress: address,
    },
    formatId,
  }
}

function dedupeProjectionRanges<T>(entries: readonly T[], keyOf: (entry: T) => string): readonly T[] {
  const byKey = new Map<string, T>()
  for (const entry of entries) {
    byKey.set(keyOf(entry), entry)
  }
  return [...byKey.values()]
}

function mergeDerivedStyleRanges(
  projected: ArrayProjectionField<SheetStyleRangeSnapshot>,
  derived: readonly SheetStyleRangeSnapshot[],
  cellsPresent: boolean,
): ArrayProjectionField<SheetStyleRangeSnapshot> {
  if (projected.present || cellsPresent) {
    return {
      present: true,
      values: dedupeProjectionRanges([...projected.values, ...derived], (entry) =>
        [entry.styleId, entry.range.startAddress, entry.range.endAddress].join('\u0000'),
      ),
    }
  }
  return projected
}

function mergeDerivedFormatRanges(
  projected: ArrayProjectionField<SheetFormatRangeSnapshot>,
  derived: readonly SheetFormatRangeSnapshot[],
  cellsPresent: boolean,
): ArrayProjectionField<SheetFormatRangeSnapshot> {
  if (projected.present || cellsPresent) {
    return {
      present: true,
      values: dedupeProjectionRanges([...projected.values, ...derived], (entry) =>
        [entry.formatId, entry.range.startAddress, entry.range.endAddress].join('\u0000'),
      ),
    }
  }
  return projected
}

function withSheetMetadataFallback(
  sheetName: string,
  rowEntries: ArrayProjectionField<WorkbookAxisMetadataSnapshot>,
  columnEntries: ArrayProjectionField<WorkbookAxisMetadataSnapshot>,
  styleRanges: ArrayProjectionField<SheetStyleRangeSnapshot>,
  formatRanges: ArrayProjectionField<SheetFormatRangeSnapshot>,
  freezePane: OptionalProjectionField<WorkbookFreezePaneSnapshot>,
  fallback?: SheetMetadataSnapshot,
) {
  const next = preserveSnapshotOnlySheetMetadata(fallback)
  if (rowEntries.present) {
    if (rowEntries.values.length > 0) {
      next.rowMetadata = mergeAxisMetadataEntries(rowEntries.values, fallback?.rowMetadata)
    }
  } else if (fallback?.rowMetadata) {
    next.rowMetadata = fallback.rowMetadata
  }
  if (columnEntries.present) {
    if (columnEntries.values.length > 0) {
      next.columnMetadata = mergeAxisMetadataEntries(columnEntries.values, fallback?.columnMetadata)
    }
  } else if (fallback?.columnMetadata) {
    next.columnMetadata = fallback.columnMetadata
  }
  if (styleRanges.present) {
    if (styleRanges.values.length > 0) {
      next.styleRanges = styleRanges.values.map((entry) => ({
        ...entry,
        range: {
          ...entry.range,
          sheetName,
        },
      }))
    }
  } else if (fallback?.styleRanges) {
    next.styleRanges = fallback.styleRanges
  }
  if (formatRanges.present) {
    if (formatRanges.values.length > 0) {
      next.formatRanges = formatRanges.values.map((entry) => ({
        ...entry,
        range: {
          ...entry.range,
          sheetName,
        },
      }))
    }
  } else if (fallback?.formatRanges) {
    next.formatRanges = fallback.formatRanges
  }
  if (freezePane.present) {
    if (freezePane.value) {
      next.freezePane = freezePane.value
    }
  } else if (fallback?.freezePane) {
    next.freezePane = fallback.freezePane
  }
  return Object.keys(next).length > 0 ? next : undefined
}

export function projectWorkbookToSnapshot(value: unknown, documentId: string) {
  if (!isRecord(value)) {
    return null
  }

  const baseSnapshot = isWorkbookSnapshot(value['snapshot']) ? value['snapshot'] : createEmptyWorkbookSnapshot(documentId)
  const workbookName = asString(value['name']) ?? baseSnapshot.workbook.name ?? documentId

  const workbookMetadata = parseArrayProjectionField(value, 'workbookMetadataEntries', parseWorkbookProperties)
  const definedNames = parseArrayProjectionField(value, 'definedNames', parseDefinedNames)
  const styles = parseArrayProjectionField(value, 'styles', parseStyleRecords)
  const numberFormats = parseArrayProjectionField(value, 'numberFormats', parseNumberFormats)
  const numberFormatCodeById = new Map(numberFormats.values.map((entry) => [entry.id, entry.code]))

  const calculationSettingsRecord = isRecord(value['calculationSettings']) ? value['calculationSettings'] : null
  const calculationMode = calculationSettingsRecord ? asString(calculationSettingsRecord['mode']) : asString(value['calcMode'])
  const compatibilityMode = asString(value['compatibilityMode'])
  const recalcEpoch =
    calculationSettingsRecord?.['recalcEpoch'] !== undefined
      ? asSafeNonNegativeInteger(calculationSettingsRecord['recalcEpoch'])
      : asSafeNonNegativeInteger(value['recalcEpoch'])
  const calculationSettingsPresent = hasOwn(value, 'calculationSettings') || hasOwn(value, 'calcMode') || hasOwn(value, 'compatibilityMode')
  const recalcEpochPresent =
    hasOwn(value, 'recalcEpoch') || (calculationSettingsRecord !== null && hasOwn(calculationSettingsRecord, 'recalcEpoch'))

  const fallbackSheets = new Map(baseSnapshot.sheets.map((sheet) => [sheet.name, sheet]))
  const sheetsPresent = hasOwn(value, 'sheets')
  const projectedSheets = (sheetsPresent ? asArray(value['sheets']) : [])
    .map((sheetEntry) => {
      if (!isRecord(sheetEntry)) {
        return null
      }
      const sheetName = asString(sheetEntry['name'])
      const sortOrder = asSafeNonNegativeInteger(sheetEntry['sortOrder'])
      if (!sheetName || sortOrder === undefined) {
        return null
      }

      const fallbackSheet = fallbackSheets.get(sheetName)
      const cellsPresent = hasOwn(sheetEntry, 'cells')
      const derivedStyleRanges: SheetStyleRangeSnapshot[] = []
      const derivedFormatRanges: SheetFormatRangeSnapshot[] = []
      const cells = cellsPresent
        ? asArray(sheetEntry['cells'])
            .map((cellEntry) => {
              if (!isRecord(cellEntry)) {
                return null
              }
              const coordinates = parseCellCoordinates(cellEntry, sheetName)
              if (!coordinates) {
                return null
              }
              const { address, rowNum, colNum } = coordinates
              const styleId = asString(cellEntry['styleId'])
              const explicitFormatId = firstString(cellEntry, ['explicitFormatId', 'formatId'])
              const inputValue = cellEntry['inputValue']
              const formula = asString(cellEntry['formula'])
              const format = asString(cellEntry['format']) ?? (explicitFormatId ? numberFormatCodeById.get(explicitFormatId) : undefined)
              const nextCell: WorkbookSnapshot['sheets'][number]['cells'][number] = { address }
              if (formula) {
                nextCell.formula = formula
              } else if (isLiteralInput(inputValue)) {
                nextCell.value = inputValue
              }
              if (format) {
                nextCell.format = format
              }
              const styleRange = singleCellStyleRange(rowNum, colNum, styleId)
              if (styleRange) {
                derivedStyleRanges.push(styleRange)
              }
              const formatRange = singleCellFormatRange(rowNum, colNum, explicitFormatId)
              if (formatRange) {
                derivedFormatRanges.push(formatRange)
              }
              return nextCell
            })
            .filter((entry): entry is WorkbookSnapshot['sheets'][number]['cells'][number] => entry !== null)
        : (fallbackSheet?.cells ?? [])
      const metadata = withSheetMetadataFallback(
        sheetName,
        parseArrayProjectionField(sheetEntry, 'rowMetadata', parseAxisMetadata),
        parseArrayProjectionField(sheetEntry, 'columnMetadata', parseAxisMetadata),
        mergeDerivedStyleRanges(parseArrayProjectionField(sheetEntry, 'styleRanges', parseStyleRanges), derivedStyleRanges, cellsPresent),
        mergeDerivedFormatRanges(
          parseArrayProjectionField(sheetEntry, 'formatRanges', parseFormatRanges),
          derivedFormatRanges,
          cellsPresent,
        ),
        parseFreezePane(sheetEntry, fallbackSheet?.metadata?.freezePane),
        fallbackSheet?.metadata,
      )

      const id =
        firstSafeNonNegativeInteger(sheetEntry, ['id', 'sheetId']) ??
        (!hasOwn(sheetEntry, 'id') && !hasOwn(sheetEntry, 'sheetId') ? fallbackSheet?.id : undefined)
      const nextSheet: WorkbookSnapshot['sheets'][number] = metadata
        ? { name: sheetName, order: sortOrder, metadata, cells }
        : { name: sheetName, order: sortOrder, cells }
      if (id !== undefined) {
        nextSheet.id = id
      }
      return nextSheet
    })
    .filter((entry): entry is WorkbookSnapshot['sheets'][number] => entry !== null)

  const workbookMetadataSnapshot = preserveSnapshotOnlyWorkbookMetadata(baseSnapshot.workbook.metadata)

  const properties = mergeProjectedArray(workbookMetadata, baseSnapshot.workbook.metadata?.properties)
  if (properties) {
    workbookMetadataSnapshot.properties = [...properties]
  }
  const definedNameEntries = mergeProjectedArray(definedNames, baseSnapshot.workbook.metadata?.definedNames)
  if (definedNameEntries) {
    workbookMetadataSnapshot.definedNames = [...definedNameEntries]
  }
  const styleEntries = mergeProjectedArray(styles, baseSnapshot.workbook.metadata?.styles)
  if (styleEntries) {
    workbookMetadataSnapshot.styles = [...styleEntries]
  }
  const formatEntries = mergeProjectedArray(numberFormats, baseSnapshot.workbook.metadata?.formats)
  if (formatEntries) {
    workbookMetadataSnapshot.formats = [...formatEntries]
  }
  if ((calculationMode === 'automatic' || calculationMode === 'manual') && isCompatibilityMode(compatibilityMode)) {
    workbookMetadataSnapshot.calculationSettings = {
      ...baseSnapshot.workbook.metadata?.calculationSettings,
      mode: calculationMode,
      compatibilityMode,
    }
  } else if (calculationMode === 'automatic' || calculationMode === 'manual') {
    workbookMetadataSnapshot.calculationSettings = {
      ...baseSnapshot.workbook.metadata?.calculationSettings,
      mode: calculationMode,
    }
  } else if (!calculationSettingsPresent && baseSnapshot.workbook.metadata?.calculationSettings) {
    workbookMetadataSnapshot.calculationSettings = baseSnapshot.workbook.metadata.calculationSettings
  }
  if (recalcEpoch !== undefined) {
    workbookMetadataSnapshot.volatileContext = {
      recalcEpoch,
    }
  } else if (!recalcEpochPresent && baseSnapshot.workbook.metadata?.volatileContext) {
    workbookMetadataSnapshot.volatileContext = baseSnapshot.workbook.metadata.volatileContext
  }

  const workbook =
    Object.keys(workbookMetadataSnapshot).length > 0 ? { name: workbookName, metadata: workbookMetadataSnapshot } : { name: workbookName }

  return {
    version: 1,
    workbook,
    sheets: sheetsPresent ? projectedSheets : baseSnapshot.sheets,
  }
}
