import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from './lookup.js'

interface LookupDatabaseBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue
  numberResult: (value: number) => CellValue
  isError: (value: LookupBuiltinArgument | undefined) => value is Extract<CellValue, { tag: ValueTag.Error }>
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument
  toNumber: (value: CellValue) => number | undefined
  toStringValue: (value: CellValue) => string
  requireCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue
  getRangeValue: (range: RangeBuiltinArgument, row: number, col: number) => CellValue
  matchesCriteria: (value: CellValue, criteria: CellValue) => boolean
}

type DatabaseCriteriaClause = {
  columnIndex: number
  criteria: CellValue
}

type DatabaseCriteriaRow = {
  clauses: DatabaseCriteriaClause[]
  blocked: boolean
}

type DatabaseSelection = {
  database: RangeBuiltinArgument
  matchingRows: number[]
  fieldIndex?: number
}

function normalizeHeaderLabel(value: CellValue, deps: LookupDatabaseBuiltinDeps): string {
  return deps.toStringValue(value).trim().toUpperCase()
}

function scalarFromLookupArgument(arg: LookupBuiltinArgument, deps: LookupDatabaseBuiltinDeps): CellValue {
  if (!deps.isRangeArg(arg)) {
    return arg
  }
  if (arg.refKind !== 'cells' || arg.values.length !== 1) {
    return deps.errorValue(ErrorCode.Value)
  }
  return arg.values[0] ?? { tag: ValueTag.Empty }
}

function resolveDatabaseFieldIndex(
  database: RangeBuiltinArgument,
  fieldArg: LookupBuiltinArgument,
  allowOmitted: boolean,
  deps: LookupDatabaseBuiltinDeps,
): number | undefined | CellValue {
  const field = scalarFromLookupArgument(fieldArg, deps)
  if (deps.isError(field)) {
    return field
  }
  if (field.tag === ValueTag.Empty) {
    return allowOmitted ? undefined : deps.errorValue(ErrorCode.Value)
  }

  if (field.tag === ValueTag.Number) {
    const position = Math.trunc(field.value)
    return position >= 1 && position <= database.cols ? position - 1 : deps.errorValue(ErrorCode.Value)
  }

  if (field.tag !== ValueTag.String) {
    return deps.errorValue(ErrorCode.Value)
  }

  const normalizedField = normalizeHeaderLabel(field, deps)
  if (normalizedField === '') {
    return allowOmitted ? undefined : deps.errorValue(ErrorCode.Value)
  }
  for (let col = 0; col < database.cols; col += 1) {
    if (normalizeHeaderLabel(deps.getRangeValue(database, 0, col), deps) === normalizedField) {
      return col
    }
  }
  return deps.errorValue(ErrorCode.Value)
}

function buildDatabaseCriteriaRows(
  database: RangeBuiltinArgument,
  criteria: RangeBuiltinArgument,
  deps: LookupDatabaseBuiltinDeps,
): DatabaseCriteriaRow[] | CellValue {
  if (criteria.rows < 2 || criteria.cols < 1) {
    return deps.errorValue(ErrorCode.Value)
  }

  const headerColumns: number[] = []
  const headerBlocked: boolean[] = []
  for (let col = 0; col < criteria.cols; col += 1) {
    const header = deps.getRangeValue(criteria, 0, col)
    if (header.tag === ValueTag.Error) {
      return header
    }
    const normalized = normalizeHeaderLabel(header, deps)
    if (normalized === '') {
      headerColumns.push(-1)
      headerBlocked.push(true)
      continue
    }
    let matchedColumn = -1
    for (let databaseCol = 0; databaseCol < database.cols; databaseCol += 1) {
      if (normalizeHeaderLabel(deps.getRangeValue(database, 0, databaseCol), deps) === normalized) {
        matchedColumn = databaseCol
        break
      }
    }
    headerColumns.push(matchedColumn)
    headerBlocked.push(matchedColumn < 0)
  }

  const rows: DatabaseCriteriaRow[] = []
  for (let row = 1; row < criteria.rows; row += 1) {
    const clauses: DatabaseCriteriaClause[] = []
    let blocked = false
    for (let col = 0; col < criteria.cols; col += 1) {
      const value = deps.getRangeValue(criteria, row, col)
      if (value.tag === ValueTag.Empty) {
        continue
      }
      if (value.tag === ValueTag.Error) {
        return value
      }
      if (headerBlocked[col]) {
        blocked = true
        continue
      }
      const databaseColumn = headerColumns[col]
      if (databaseColumn === undefined || databaseColumn < 0) {
        continue
      }
      clauses.push({ columnIndex: databaseColumn, criteria: value })
    }
    rows.push({ clauses, blocked })
  }
  return rows
}

function recordMatchesDatabaseCriteria(
  database: RangeBuiltinArgument,
  databaseRow: number,
  criteriaRows: readonly DatabaseCriteriaRow[],
  deps: LookupDatabaseBuiltinDeps,
): boolean {
  for (const criteriaRow of criteriaRows) {
    if (criteriaRow.blocked) {
      continue
    }
    if (
      criteriaRow.clauses.every((clause) =>
        deps.matchesCriteria(deps.getRangeValue(database, databaseRow, clause.columnIndex), clause.criteria),
      )
    ) {
      return true
    }
  }
  return false
}

function matchingDatabaseRows(
  databaseArg: LookupBuiltinArgument,
  criteriaArg: LookupBuiltinArgument,
  deps: LookupDatabaseBuiltinDeps,
): { database: RangeBuiltinArgument; matchingRows: number[] } | CellValue {
  const database = deps.requireCellRange(databaseArg)
  if (!deps.isRangeArg(database)) {
    return database
  }
  const criteria = deps.requireCellRange(criteriaArg)
  if (!deps.isRangeArg(criteria)) {
    return criteria
  }
  if (database.rows < 1 || database.cols < 1) {
    return deps.errorValue(ErrorCode.Value)
  }

  const criteriaRows = buildDatabaseCriteriaRows(database, criteria, deps)
  if (!Array.isArray(criteriaRows)) {
    return criteriaRows
  }

  const matchingRows: number[] = []
  for (let row = 1; row < database.rows; row += 1) {
    if (recordMatchesDatabaseCriteria(database, row, criteriaRows, deps)) {
      matchingRows.push(row)
    }
  }
  return { database, matchingRows }
}

function selectedDatabaseFieldValues(
  databaseArg: LookupBuiltinArgument,
  fieldArg: LookupBuiltinArgument,
  criteriaArg: LookupBuiltinArgument,
  allowOmittedField: boolean,
  deps: LookupDatabaseBuiltinDeps,
): DatabaseSelection | CellValue {
  const matches = matchingDatabaseRows(databaseArg, criteriaArg, deps)
  if ('tag' in matches) {
    return matches
  }
  const fieldIndex = resolveDatabaseFieldIndex(matches.database, fieldArg, allowOmittedField, deps)
  if (typeof fieldIndex !== 'number' && fieldIndex !== undefined) {
    return fieldIndex
  }
  return fieldIndex === undefined
    ? {
        database: matches.database,
        matchingRows: matches.matchingRows,
      }
    : {
        database: matches.database,
        matchingRows: matches.matchingRows,
        fieldIndex,
      }
}

function sampleVariance(numbers: readonly number[]): number | undefined {
  if (numbers.length < 2) {
    return undefined
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length
  return numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (numbers.length - 1)
}

function populationVariance(numbers: readonly number[]): number | undefined {
  if (numbers.length === 0) {
    return undefined
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length
  return numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / numbers.length
}

function collectSelectionNumbers(selection: DatabaseSelection, deps: LookupDatabaseBuiltinDeps): number[] {
  const values: number[] = []
  for (const row of selection.matchingRows) {
    const numeric = deps.toNumber(deps.getRangeValue(selection.database, row, selection.fieldIndex ?? 0))
    if (numeric !== undefined) {
      values.push(numeric)
    }
  }
  return values
}

export function createLookupDatabaseBuiltins(deps: LookupDatabaseBuiltinDeps): Record<string, LookupBuiltin> {
  return {
    DAVERAGE: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }

      const values = collectSelectionNumbers(selection, deps)
      if (values.length === 0) {
        return deps.errorValue(ErrorCode.Div0)
      }
      const sum = values.reduce((total, value) => total + value, 0)
      return deps.numberResult(sum / values.length)
    },
    DCOUNT: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, true, deps)
      if ('tag' in selection) {
        return selection
      }
      if (selection.fieldIndex === undefined) {
        return deps.numberResult(selection.matchingRows.length)
      }
      return deps.numberResult(collectSelectionNumbers(selection, deps).length)
    },
    DCOUNTA: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, true, deps)
      if ('tag' in selection) {
        return selection
      }
      if (selection.fieldIndex === undefined) {
        return deps.numberResult(selection.matchingRows.length)
      }
      let count = 0
      for (const row of selection.matchingRows) {
        if (deps.getRangeValue(selection.database, row, selection.fieldIndex).tag !== ValueTag.Empty) {
          count += 1
        }
      }
      return deps.numberResult(count)
    },
    DGET: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }
      if (selection.matchingRows.length !== 1) {
        return deps.errorValue(ErrorCode.Value)
      }
      return deps.getRangeValue(selection.database, selection.matchingRows[0]!, selection.fieldIndex ?? 0)
    },
    DMAX: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }
      const values = collectSelectionNumbers(selection, deps)
      return deps.numberResult(values.length === 0 ? 0 : Math.max(...values))
    },
    DMIN: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }
      const values = collectSelectionNumbers(selection, deps)
      return deps.numberResult(values.length === 0 ? 0 : Math.min(...values))
    },
    DPRODUCT: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }
      const values = collectSelectionNumbers(selection, deps)
      if (values.length === 0) {
        return deps.numberResult(0)
      }
      return deps.numberResult(values.reduce((product, value) => product * value, 1))
    },
    DSTDEV: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }
      const variance = sampleVariance(collectSelectionNumbers(selection, deps))
      return variance === undefined ? deps.errorValue(ErrorCode.Div0) : deps.numberResult(Math.sqrt(variance))
    },
    DSTDEVP: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }
      const variance = populationVariance(collectSelectionNumbers(selection, deps))
      return variance === undefined ? deps.errorValue(ErrorCode.Div0) : deps.numberResult(Math.sqrt(variance))
    },
    DSUM: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }
      return deps.numberResult(collectSelectionNumbers(selection, deps).reduce((sum, value) => sum + value, 0))
    },
    DVAR: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }
      const variance = sampleVariance(collectSelectionNumbers(selection, deps))
      return variance === undefined ? deps.errorValue(ErrorCode.Div0) : deps.numberResult(variance)
    },
    DVARP: (databaseArg, fieldArg, criteriaArg) => {
      const selection = selectedDatabaseFieldValues(databaseArg, fieldArg, criteriaArg, false, deps)
      if ('tag' in selection) {
        return selection
      }
      const variance = populationVariance(collectSelectionNumbers(selection, deps))
      return variance === undefined ? deps.errorValue(ErrorCode.Div0) : deps.numberResult(variance)
    },
  }
}
