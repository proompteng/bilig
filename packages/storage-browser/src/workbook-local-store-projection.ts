import type { Database, SqlValue } from '@sqlite.org/sqlite-wasm'
import { ValueTag, type CellSnapshot, type CellStyleRecord, type WorkbookAxisEntrySnapshot } from '@bilig/protocol'
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

function parseCellSnapshotValue(value: unknown): CellSnapshot['value'] | null {
  if (!isRecord(value) || typeof value['tag'] !== 'number') {
    return null
  }
  switch (value['tag']) {
    case 0:
      return { tag: ValueTag.Empty }
    case 1:
      return typeof value['value'] === 'number' ? { tag: ValueTag.Number, value: value['value'] } : null
    case 2:
      return typeof value['value'] === 'boolean' ? { tag: ValueTag.Boolean, value: value['value'] } : null
    case 3:
      return typeof value['value'] === 'string'
        ? {
            tag: ValueTag.String,
            value: value['value'],
            stringId: typeof value['stringId'] === 'number' ? value['stringId'] : 0,
          }
        : null
    case 4:
      return typeof value['code'] === 'number' ? { tag: ValueTag.Error, code: value['code'] } : null
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
    typeof rowNum !== 'number' ||
    typeof colNum !== 'number' ||
    typeof valueJson !== 'string' ||
    typeof flags !== 'number' ||
    typeof version !== 'number'
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
      if (parsedInput === null || typeof parsedInput === 'boolean' || typeof parsedInput === 'number' || typeof parsedInput === 'string') {
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
  if (typeof id !== 'string' || typeof entryIndex !== 'number') {
    return null
  }
  const entry: WorkbookAxisEntrySnapshot = {
    id,
    index: entryIndex,
  }
  if (typeof row['size'] === 'number') {
    entry.size = row['size']
  }
  if (typeof row['hidden'] === 'number') {
    entry.hidden = row['hidden'] !== 0
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
    if (!isRecord(parsed)) {
      return null
    }
    return {
      ...(parsed as Omit<CellStyleRecord, 'id'>),
      id,
    }
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

function insertWorkbookAuthoritativeBaseRows(db: Database, base: WorkbookLocalAuthoritativeBase, includeSheets = true): void {
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
      for (const sheet of base.sheets) {
        insertSheet.bind([sheet.sheetId, sheet.name, sheet.sortOrder, sheet.freezeRows, sheet.freezeCols])
        insertSheet.step()
        insertSheet.reset()
      }
    }
    for (const cell of base.cellInputs) {
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
    for (const cell of base.cellRenders) {
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
    for (const axis of base.rowAxisEntries) {
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
    for (const axis of base.columnAxisEntries) {
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
  upsertAuthoritativeSheets(db, delta.base.sheets)
  insertWorkbookAuthoritativeBaseRows(db, delta.base, false)
  const persistedSheetIds = new Set(delta.base.sheets.map((sheet) => sheet.sheetId))
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
  if (typeof sheetId !== 'number') {
    return null
  }
  const freezeRows = typeof sheetRecord['freezeRows'] === 'number' ? sheetRecord['freezeRows'] : 0
  const freezeCols = typeof sheetRecord['freezeCols'] === 'number' ? sheetRecord['freezeCols'] : 0

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
