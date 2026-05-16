import type { Database, SqlValue } from '@sqlite.org/sqlite-wasm'
import {
  ErrorCode,
  sanitizeCellStyleRecord,
  ValueTag,
  type CellSnapshot,
  type CellStyleRecord,
  type LiteralInput,
  type WorkbookAxisEntrySnapshot,
} from '@bilig/protocol'
import type {
  WorkbookLocalAuthoritativeDelta,
  WorkbookLocalAuthoritativeBase,
  WorkbookLocalProjectionOverlay,
  WorkbookLocalViewportBase,
  WorkbookLocalViewportCell,
} from './workbook-local-base.js'

interface ViewportBounds {
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'boolean' || typeof value === 'string' || isFiniteNumber(value)
}

function isErrorCode(value: unknown): value is ErrorCode {
  return (
    value === ErrorCode.None ||
    value === ErrorCode.Div0 ||
    value === ErrorCode.Ref ||
    value === ErrorCode.Value ||
    value === ErrorCode.Name ||
    value === ErrorCode.NA ||
    value === ErrorCode.Cycle ||
    value === ErrorCode.Spill ||
    value === ErrorCode.Blocked
  )
}

function parseCellSnapshotValue(value: unknown): CellSnapshot['value'] | null {
  if (!isRecord(value)) {
    return null
  }
  switch (value['tag']) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty }
    case ValueTag.Number:
      return isFiniteNumber(value['value']) ? { tag: ValueTag.Number, value: value['value'] } : null
    case ValueTag.Boolean:
      return typeof value['value'] === 'boolean' ? { tag: ValueTag.Boolean, value: value['value'] } : null
    case ValueTag.String:
      return typeof value['value'] === 'string' && isSafeNonNegativeInteger(value['stringId'])
        ? {
            tag: ValueTag.String,
            value: value['value'],
            stringId: value['stringId'],
          }
        : null
    case ValueTag.Error:
      return isErrorCode(value['code']) ? { tag: ValueTag.Error, code: value['code'] } : null
    default:
      return null
  }
}

function parseViewportCellFromRow(row: Record<string, SqlValue>): WorkbookLocalViewportCell | null {
  const address = row['address']
  const sheetName = row['sheetName']
  const rowNum = row['rowNum']
  const colNum = row['colNum']
  const valueJson = row['valueJson']
  const flags = row['flags']
  const version = row['version']
  if (
    typeof address !== 'string' ||
    typeof sheetName !== 'string' ||
    !isSafeNonNegativeInteger(rowNum) ||
    !isSafeNonNegativeInteger(colNum) ||
    typeof valueJson !== 'string' ||
    !isSafeNonNegativeInteger(flags) ||
    !isSafeNonNegativeInteger(version)
  ) {
    return null
  }
  try {
    const parsedValue = parseCellSnapshotValue(JSON.parse(valueJson) as unknown)
    if (!parsedValue) {
      return null
    }
    const snapshot: CellSnapshot = {
      sheetName,
      address,
      value: parsedValue,
      flags,
      version,
    }
    const inputJson = row['inputJson']
    if (typeof inputJson === 'string') {
      const parsedInput = JSON.parse(inputJson) as unknown
      if (isLiteralInput(parsedInput)) {
        snapshot.input = parsedInput
      }
    }
    if (typeof row['formula'] === 'string') {
      snapshot.formula = row['formula']
    }
    if (typeof row['format'] === 'string') {
      snapshot.format = row['format']
    }
    if (typeof row['styleId'] === 'string') {
      snapshot.styleId = row['styleId']
    }
    if (typeof row['numberFormatId'] === 'string') {
      snapshot.numberFormatId = row['numberFormatId']
    }
    return {
      row: rowNum,
      col: colNum,
      snapshot,
    }
  } catch {
    return null
  }
}

function parseAxisEntrySnapshot(row: Record<string, SqlValue>): WorkbookAxisEntrySnapshot | null {
  const id = row['id']
  const entryIndex = row['entryIndex']
  if (typeof id !== 'string' || !isSafeNonNegativeInteger(entryIndex)) {
    return null
  }
  const entry: WorkbookAxisEntrySnapshot = {
    id,
    index: entryIndex,
  }
  if (isSafeNonNegativeInteger(row['size'])) {
    entry.size = row['size']
  }
  if (row['hidden'] === 0 || row['hidden'] === 1) {
    entry.hidden = row['hidden'] === 1
  } else if (typeof row['hidden'] === 'boolean') {
    entry.hidden = row['hidden']
  }
  return entry
}

function parseCellStyleRecord(row: Record<string, SqlValue>): CellStyleRecord | null {
  const id = row['id']
  const recordJson = row['recordJson']
  if (typeof id !== 'string' || typeof recordJson !== 'string') {
    return null
  }
  try {
    const parsed = JSON.parse(recordJson) as unknown
    return sanitizeCellStyleRecord(id, parsed)
  } catch {
    return null
  }
}

function readSingleObjectRow(db: Database, sql: string, bind?: readonly SqlValue[]): Record<string, SqlValue> | null {
  const statement = db.prepare(sql)
  try {
    if (bind) {
      statement.bind([...bind])
    }
    if (!statement.step()) {
      return null
    }
    return statement.get({})
  } finally {
    statement.finalize()
  }
}

function clearWorkbookProjectionTables(db: Database): void {
  db.exec('DELETE FROM projection_overlay_cell')
  db.exec('DELETE FROM projection_overlay_row_axis')
  db.exec('DELETE FROM projection_overlay_column_axis')
  db.exec('DELETE FROM projection_overlay_style')
}

function replaceAuthoritativeStyles(db: Database, styles: readonly CellStyleRecord[]): void {
  db.exec('DELETE FROM authoritative_style')
  const insertStyle = db.prepare(
    `
      INSERT INTO authoritative_style (style_id, record_json)
      VALUES (?, ?)
    `,
  )
  try {
    for (const style of styles) {
      insertStyle.bind([style.id, JSON.stringify(style)])
      insertStyle.step()
      insertStyle.reset()
    }
  } finally {
    insertStyle.finalize()
  }
}

function collectCanonicalAuthoritativeSheets(base: WorkbookLocalAuthoritativeBase): WorkbookLocalAuthoritativeBase['sheets'] {
  const sheetsById = new Map<number, WorkbookLocalAuthoritativeBase['sheets'][number]>()
  base.sheets.forEach((sheet) => {
    sheetsById.set(sheet.sheetId, sheet)
  })

  const collect = (sheetId: number, sheetName: string): void => {
    if (sheetsById.has(sheetId)) {
      return
    }
    sheetsById.set(sheetId, {
      sheetId,
      name: sheetName,
      sortOrder: Number.MAX_SAFE_INTEGER,
      freezeRows: 0,
      freezeCols: 0,
    })
  }

  base.cellInputs.forEach((cell) => collect(cell.sheetId, cell.sheetName))
  base.cellRenders.forEach((cell) => collect(cell.sheetId, cell.sheetName))
  base.rowAxisEntries.forEach((entry) => collect(entry.sheetId, entry.sheetName))
  base.columnAxisEntries.forEach((entry) => collect(entry.sheetId, entry.sheetName))

  return [...sheetsById.values()].toSorted((left, right) => left.sortOrder - right.sortOrder || left.sheetId - right.sheetId)
}

function normalizeAuthoritativeBaseSheetNames(base: WorkbookLocalAuthoritativeBase): WorkbookLocalAuthoritativeBase {
  const sheets = collectCanonicalAuthoritativeSheets(base)
  const sheetNamesById = new Map(sheets.map((sheet) => [sheet.sheetId, sheet.name]))
  const resolveSheetName = (sheetId: number, fallbackName: string): string => sheetNamesById.get(sheetId) ?? fallbackName

  return {
    sheets,
    cellInputs: base.cellInputs.map((cell) => ({
      ...cell,
      sheetName: resolveSheetName(cell.sheetId, cell.sheetName),
    })),
    cellRenders: base.cellRenders.map((cell) => ({
      ...cell,
      sheetName: resolveSheetName(cell.sheetId, cell.sheetName),
    })),
    rowAxisEntries: base.rowAxisEntries.map((entry) => ({
      ...entry,
      sheetName: resolveSheetName(entry.sheetId, entry.sheetName),
    })),
    columnAxisEntries: base.columnAxisEntries.map((entry) => ({
      ...entry,
      sheetName: resolveSheetName(entry.sheetId, entry.sheetName),
    })),
    styles: base.styles,
  }
}

function insertWorkbookAuthoritativeBaseRows(db: Database, base: WorkbookLocalAuthoritativeBase, includeSheets = true): void {
  const normalizedBase = normalizeAuthoritativeBaseSheetNames(base)
  const insertSheet = db.prepare(
    `
      INSERT INTO authoritative_sheet (sheet_id, name, sort_order, freeze_rows, freeze_cols)
      VALUES (?, ?, ?, ?, ?)
    `,
  )
  const insertInput = db.prepare(
    `
      INSERT INTO authoritative_cell_input (
        sheet_id,
        sheet_name,
        address,
        row_num,
        col_num,
        input_json,
        formula,
        format
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `,
  )
  const insertRender = db.prepare(
    `
      INSERT INTO authoritative_cell_render (
        sheet_id,
        sheet_name,
        address,
        row_num,
        col_num,
        value_json,
        flags,
        version,
        style_id,
        number_format_id
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `,
  )
  const insertAxis = (tableName: 'authoritative_row_axis' | 'authoritative_column_axis') =>
    db.prepare(
      `
        INSERT INTO ${tableName} (
          sheet_id,
          sheet_name,
          axis_index,
          axis_id,
          size,
          hidden
        )
        VALUES (?, ?, ?, ?, ?, ?)
      `,
    )
  const insertRowAxis = insertAxis('authoritative_row_axis')
  const insertColumnAxis = insertAxis('authoritative_column_axis')
  try {
    if (includeSheets) {
      for (const sheet of normalizedBase.sheets) {
        insertSheet.bind([sheet.sheetId, sheet.name, sheet.sortOrder, sheet.freezeRows, sheet.freezeCols])
        insertSheet.step()
        insertSheet.reset()
      }
    }
    for (const cell of normalizedBase.cellInputs) {
      insertInput.bind([
        cell.sheetId,
        cell.sheetName,
        cell.address,
        cell.rowNum,
        cell.colNum,
        cell.input === undefined ? null : JSON.stringify(cell.input),
        cell.formula ?? null,
        cell.format ?? null,
      ])
      insertInput.step()
      insertInput.reset()
    }
    for (const cell of normalizedBase.cellRenders) {
      insertRender.bind([
        cell.sheetId,
        cell.sheetName,
        cell.address,
        cell.rowNum,
        cell.colNum,
        JSON.stringify(cell.value),
        cell.flags,
        cell.version,
        cell.styleId ?? null,
        cell.numberFormatId ?? null,
      ])
      insertRender.step()
      insertRender.reset()
    }
    for (const axis of normalizedBase.rowAxisEntries) {
      insertRowAxis.bind([
        axis.sheetId,
        axis.sheetName,
        axis.entry.index,
        axis.entry.id,
        axis.entry.size ?? null,
        axis.entry.hidden ?? null,
      ])
      insertRowAxis.step()
      insertRowAxis.reset()
    }
    for (const axis of normalizedBase.columnAxisEntries) {
      insertColumnAxis.bind([
        axis.sheetId,
        axis.sheetName,
        axis.entry.index,
        axis.entry.id,
        axis.entry.size ?? null,
        axis.entry.hidden ?? null,
      ])
      insertColumnAxis.step()
      insertColumnAxis.reset()
    }
  } finally {
    insertSheet.finalize()
    insertInput.finalize()
    insertRender.finalize()
    insertRowAxis.finalize()
    insertColumnAxis.finalize()
  }
}

function upsertAuthoritativeSheets(db: Database, sheets: WorkbookLocalAuthoritativeBase['sheets']): void {
  if (sheets.length === 0) {
    return
  }
  const updateSheet = db.prepare(
    `
      UPDATE authoritative_sheet
         SET name = ?,
             sort_order = ?,
             freeze_rows = ?,
             freeze_cols = ?
       WHERE sheet_id = ?
    `,
  )
  const insertSheet = db.prepare(
    `
      INSERT INTO authoritative_sheet (sheet_id, name, sort_order, freeze_rows, freeze_cols)
      VALUES (?, ?, ?, ?, ?)
    `,
  )
  try {
    for (const sheet of sheets) {
      updateSheet.bind([sheet.name, sheet.sortOrder, sheet.freezeRows, sheet.freezeCols, sheet.sheetId])
      updateSheet.step()
      const updated = db.changes(true, false) > 0
      updateSheet.reset()
      if (updated) {
        continue
      }
      insertSheet.bind([sheet.sheetId, sheet.name, sheet.sortOrder, sheet.freezeRows, sheet.freezeCols])
      insertSheet.step()
      insertSheet.reset()
    }
  } finally {
    updateSheet.finalize()
    insertSheet.finalize()
  }
}

function deleteAuthoritativeSheetData(db: Database, sheetIds: readonly number[]): void {
  if (sheetIds.length === 0) {
    return
  }
  const deleteFrom = (
    tableName: 'authoritative_cell_input' | 'authoritative_cell_render' | 'authoritative_row_axis' | 'authoritative_column_axis',
  ) => db.prepare(`DELETE FROM ${tableName} WHERE sheet_id = ?`)
  const statements = [
    deleteFrom('authoritative_cell_input'),
    deleteFrom('authoritative_cell_render'),
    deleteFrom('authoritative_row_axis'),
    deleteFrom('authoritative_column_axis'),
  ]
  try {
    for (const sheetId of sheetIds) {
      statements.forEach((statement) => {
        statement.bind([sheetId])
        statement.step()
        statement.reset()
      })
    }
  } finally {
    statements.forEach((statement) => statement.finalize())
  }
}

function deleteAuthoritativeSheets(db: Database, sheetIds: readonly number[]): void {
  if (sheetIds.length === 0) {
    return
  }
  const statement = db.prepare('DELETE FROM authoritative_sheet WHERE sheet_id = ?')
  try {
    for (const sheetId of sheetIds) {
      statement.bind([sheetId])
      statement.step()
      statement.reset()
    }
  } finally {
    statement.finalize()
  }
}

export function writeWorkbookAuthoritativeBase(db: Database, base: WorkbookLocalAuthoritativeBase): void {
  db.exec('DELETE FROM authoritative_cell_input')
  db.exec('DELETE FROM authoritative_cell_render')
  db.exec('DELETE FROM authoritative_row_axis')
  db.exec('DELETE FROM authoritative_column_axis')
  db.exec('DELETE FROM authoritative_style')
  db.exec('DELETE FROM authoritative_sheet')
  insertWorkbookAuthoritativeBaseRows(db, base)
  replaceAuthoritativeStyles(db, base.styles)
}

export function writeWorkbookAuthoritativeDelta(db: Database, delta: WorkbookLocalAuthoritativeDelta): void {
  if (delta.replaceAll) {
    writeWorkbookAuthoritativeBase(db, delta.base)
    return
  }
  deleteAuthoritativeSheetData(db, delta.replacedSheetIds)
  const referencedSheets = collectCanonicalAuthoritativeSheets(delta.base)
  upsertAuthoritativeSheets(db, referencedSheets)
  insertWorkbookAuthoritativeBaseRows(db, delta.base, false)
  const persistedSheetIds = new Set(referencedSheets.map((sheet) => sheet.sheetId))
  deleteAuthoritativeSheets(
    db,
    delta.replacedSheetIds.filter((sheetId) => !persistedSheetIds.has(sheetId)),
  )
  replaceAuthoritativeStyles(db, delta.base.styles)
}

export function writeWorkbookProjectionOverlay(db: Database, overlay: WorkbookLocalProjectionOverlay): void {
  clearWorkbookProjectionTables(db)

  const insertCell = db.prepare(
    `
      INSERT INTO projection_overlay_cell (
        sheet_id,
        sheet_name,
        address,
        row_num,
        col_num,
        value_json,
        flags,
        version,
        input_json,
        formula,
        format,
        style_id,
        number_format_id
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `,
  )
  const insertAxis = (tableName: 'projection_overlay_row_axis' | 'projection_overlay_column_axis') =>
    db.prepare(
      `
        INSERT INTO ${tableName} (
          sheet_id,
          sheet_name,
          axis_index,
          axis_id,
          size,
          hidden
        )
        VALUES (?, ?, ?, ?, ?, ?)
      `,
    )
  const insertStyle = db.prepare(
    `
      INSERT INTO projection_overlay_style (style_id, record_json)
      VALUES (?, ?)
    `,
  )
  const insertRowAxis = insertAxis('projection_overlay_row_axis')
  const insertColumnAxis = insertAxis('projection_overlay_column_axis')
  try {
    for (const cell of overlay.cells) {
      insertCell.bind([
        cell.sheetId,
        cell.sheetName,
        cell.address,
        cell.rowNum,
        cell.colNum,
        JSON.stringify(cell.value),
        cell.flags,
        cell.version,
        cell.input === undefined ? null : JSON.stringify(cell.input),
        cell.formula ?? null,
        cell.format ?? null,
        cell.styleId ?? null,
        cell.numberFormatId ?? null,
      ])
      insertCell.step()
      insertCell.reset()
    }
    for (const axis of overlay.rowAxisEntries) {
      insertRowAxis.bind([
        axis.sheetId,
        axis.sheetName,
        axis.entry.index,
        axis.entry.id,
        axis.entry.size ?? null,
        axis.entry.hidden ?? null,
      ])
      insertRowAxis.step()
      insertRowAxis.reset()
    }
    for (const axis of overlay.columnAxisEntries) {
      insertColumnAxis.bind([
        axis.sheetId,
        axis.sheetName,
        axis.entry.index,
        axis.entry.id,
        axis.entry.size ?? null,
        axis.entry.hidden ?? null,
      ])
      insertColumnAxis.step()
      insertColumnAxis.reset()
    }
    for (const style of overlay.styles) {
      insertStyle.bind([style.id, JSON.stringify(style)])
      insertStyle.step()
      insertStyle.reset()
    }
  } finally {
    insertCell.finalize()
    insertRowAxis.finalize()
    insertColumnAxis.finalize()
    insertStyle.finalize()
  }
}

function readViewportCells(db: Database, sql: string, bind: readonly SqlValue[]): WorkbookLocalViewportCell[] {
  const cells: WorkbookLocalViewportCell[] = []
  const statement = db.prepare(sql)
  try {
    statement.bind([...bind])
    while (statement.step()) {
      const parsed = parseViewportCellFromRow(statement.get({}))
      if (parsed) {
        cells.push(parsed)
      }
    }
  } finally {
    statement.finalize()
  }
  return cells
}

function readAxisEntries(
  db: Database,
  tableName: 'authoritative_row_axis' | 'authoritative_column_axis' | 'projection_overlay_row_axis' | 'projection_overlay_column_axis',
  sheetId: number,
  start: number,
  end: number,
): WorkbookAxisEntrySnapshot[] {
  const rows: WorkbookAxisEntrySnapshot[] = []
  const statement = db.prepare(
    `
      SELECT axis_id AS id,
             axis_index AS entryIndex,
             size,
             hidden
        FROM ${tableName}
       WHERE sheet_id = ?
         AND axis_index >= ?
         AND axis_index <= ?
       ORDER BY axis_index ASC
    `,
  )
  try {
    statement.bind([sheetId, start, end])
    while (statement.step()) {
      const entry = parseAxisEntrySnapshot(statement.get({}))
      if (entry) {
        rows.push(entry)
      }
    }
  } finally {
    statement.finalize()
  }
  return rows
}

function readStylesByIds(
  db: Database,
  tableName: 'authoritative_style' | 'projection_overlay_style',
  styleIds: ReadonlySet<string>,
): CellStyleRecord[] {
  if (styleIds.size === 0) {
    return []
  }
  const styles: CellStyleRecord[] = []
  const statement = db.prepare(
    `
      SELECT style_id AS id,
             record_json AS recordJson
        FROM ${tableName}
       WHERE style_id = ?
    `,
  )
  try {
    for (const styleId of styleIds) {
      statement.bind([styleId])
      if (statement.step()) {
        const style = parseCellStyleRecord(statement.get({}))
        if (style) {
          styles.push(style)
        }
      }
      statement.reset()
    }
  } finally {
    statement.finalize()
  }
  return styles
}

function sortViewportCells(cells: Iterable<WorkbookLocalViewportCell>): WorkbookLocalViewportCell[] {
  return [...cells].toSorted((left, right) => left.row - right.row || left.col - right.col)
}

function sortAxisEntries(entries: Iterable<WorkbookAxisEntrySnapshot>): WorkbookAxisEntrySnapshot[] {
  return [...entries].toSorted((left, right) => left.index - right.index)
}

function mergeViewportBaseAndOverlay(input: {
  readonly base: WorkbookLocalViewportBase
  readonly overlayCells: readonly WorkbookLocalViewportCell[]
  readonly overlayRowAxisEntries: readonly WorkbookAxisEntrySnapshot[]
  readonly overlayColumnAxisEntries: readonly WorkbookAxisEntrySnapshot[]
  readonly overlayStyles: readonly CellStyleRecord[]
}): WorkbookLocalViewportBase {
  const cells = new Map<string, WorkbookLocalViewportCell>()
  input.base.cells.forEach((cell) => {
    cells.set(cell.snapshot.address, cell)
  })
  input.overlayCells.forEach((cell) => {
    cells.set(cell.snapshot.address, cell)
  })

  const rowAxisEntries = new Map<number, WorkbookAxisEntrySnapshot>()
  input.base.rowAxisEntries.forEach((entry) => {
    rowAxisEntries.set(entry.index, entry)
  })
  input.overlayRowAxisEntries.forEach((entry) => {
    rowAxisEntries.set(entry.index, entry)
  })

  const columnAxisEntries = new Map<number, WorkbookAxisEntrySnapshot>()
  input.base.columnAxisEntries.forEach((entry) => {
    columnAxisEntries.set(entry.index, entry)
  })
  input.overlayColumnAxisEntries.forEach((entry) => {
    columnAxisEntries.set(entry.index, entry)
  })

  const styles = new Map<string, CellStyleRecord>()
  input.base.styles.forEach((style) => {
    styles.set(style.id, style)
  })
  input.overlayStyles.forEach((style) => {
    styles.set(style.id, style)
  })
  if (!styles.has('style-0')) {
    styles.set('style-0', { id: 'style-0' })
  }

  return {
    sheetId: input.base.sheetId,
    sheetName: input.base.sheetName,
    freezeRows: input.base.freezeRows,
    freezeCols: input.base.freezeCols,
    cells: sortViewportCells(cells.values()),
    rowAxisEntries: sortAxisEntries(rowAxisEntries.values()),
    columnAxisEntries: sortAxisEntries(columnAxisEntries.values()),
    styles: [...styles.values()],
  }
}

function readWorkbookViewportBase(db: Database, sheetName: string, viewport: ViewportBounds): WorkbookLocalViewportBase | null {
  const sheetRecord = readSingleObjectRow(
    db,
    `
      SELECT name,
             sheet_id AS sheetId,
             freeze_rows AS freezeRows,
             freeze_cols AS freezeCols
        FROM authoritative_sheet
       WHERE name = ?
    `,
    [sheetName],
  )
  if (!sheetRecord) {
    return null
  }
  const sheetId = sheetRecord['sheetId']
  if (!isSafeNonNegativeInteger(sheetId)) {
    return null
  }
  const freezeRows = isSafeNonNegativeInteger(sheetRecord['freezeRows']) ? sheetRecord['freezeRows'] : 0
  const freezeCols = isSafeNonNegativeInteger(sheetRecord['freezeCols']) ? sheetRecord['freezeCols'] : 0

  const cells = readViewportCells(
    db,
    `
      SELECT render.sheet_name AS sheetName,
             render.address AS address,
             render.row_num AS rowNum,
             render.col_num AS colNum,
             render.value_json AS valueJson,
             render.flags AS flags,
             render.version AS version,
             render.style_id AS styleId,
             render.number_format_id AS numberFormatId,
             input.input_json AS inputJson,
             input.formula AS formula,
             input.format AS format
        FROM authoritative_cell_render AS render
        LEFT JOIN authoritative_cell_input AS input
          ON input.sheet_id = render.sheet_id
         AND input.address = render.address
       WHERE render.sheet_id = ?
         AND render.row_num >= ?
         AND render.row_num <= ?
         AND render.col_num >= ?
         AND render.col_num <= ?
       ORDER BY render.row_num ASC, render.col_num ASC
    `,
    [sheetId, viewport.rowStart, viewport.rowEnd, viewport.colStart, viewport.colEnd],
  )
  const styleIds = new Set<string>(['style-0'])
  cells.forEach((cell) => {
    if (cell.snapshot.styleId) {
      styleIds.add(cell.snapshot.styleId)
    }
  })

  return {
    sheetId,
    sheetName,
    freezeRows,
    freezeCols,
    cells,
    rowAxisEntries: readAxisEntries(db, 'authoritative_row_axis', sheetId, viewport.rowStart, viewport.rowEnd),
    columnAxisEntries: readAxisEntries(db, 'authoritative_column_axis', sheetId, viewport.colStart, viewport.colEnd),
    styles: readStylesByIds(db, 'authoritative_style', styleIds),
  }
}

export function readWorkbookViewportProjection(
  db: Database,
  sheetName: string,
  viewport: ViewportBounds,
): WorkbookLocalViewportBase | null {
  const base = readWorkbookViewportBase(db, sheetName, viewport)
  if (!base) {
    return null
  }
  const sheetId = base.sheetId

  const overlayCells = readViewportCells(
    db,
    `
      SELECT sheet_name AS sheetName,
             address,
             row_num AS rowNum,
             col_num AS colNum,
             value_json AS valueJson,
             flags,
             version,
             input_json AS inputJson,
             formula,
             format,
             style_id AS styleId,
             number_format_id AS numberFormatId
        FROM projection_overlay_cell
       WHERE sheet_id = ?
         AND row_num >= ?
         AND row_num <= ?
         AND col_num >= ?
         AND col_num <= ?
       ORDER BY row_num ASC, col_num ASC
    `,
    [sheetId, viewport.rowStart, viewport.rowEnd, viewport.colStart, viewport.colEnd],
  )
  const overlayStyleIds = new Set<string>()
  overlayCells.forEach((cell) => {
    if (cell.snapshot.styleId) {
      overlayStyleIds.add(cell.snapshot.styleId)
    }
  })

  return mergeViewportBaseAndOverlay({
    base,
    overlayCells,
    overlayRowAxisEntries: readAxisEntries(db, 'projection_overlay_row_axis', sheetId, viewport.rowStart, viewport.rowEnd),
    overlayColumnAxisEntries: readAxisEntries(db, 'projection_overlay_column_axis', sheetId, viewport.colStart, viewport.colEnd),
    overlayStyles: readStylesByIds(db, 'projection_overlay_style', overlayStyleIds),
  })
}
