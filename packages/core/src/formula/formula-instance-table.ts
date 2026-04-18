export interface FormulaInstanceSnapshot {
  readonly cellIndex: number
  readonly sheetName: string
  readonly row: number
  readonly col: number
  readonly source: string
  readonly templateId?: number
}

export interface FormulaInstanceTable {
  readonly upsert: (record: FormulaInstanceSnapshot) => void
  readonly get: (cellIndex: number) => FormulaInstanceSnapshot | undefined
  readonly delete: (cellIndex: number) => boolean
  readonly clear: () => void
  readonly list: () => FormulaInstanceSnapshot[]
  readonly hydrate: (records: readonly FormulaInstanceSnapshot[]) => void
}

export function createFormulaInstanceTable(): FormulaInstanceTable {
  const records = new Map<number, FormulaInstanceSnapshot>()

  return {
    upsert(record) {
      records.set(record.cellIndex, record)
    },
    get(cellIndex) {
      return records.get(cellIndex)
    },
    delete(cellIndex) {
      return records.delete(cellIndex)
    },
    clear() {
      records.clear()
    },
    list() {
      return [...records.values()].toSorted(
        (left, right) =>
          left.sheetName.localeCompare(right.sheetName) || left.row - right.row || left.col - right.col || left.cellIndex - right.cellIndex,
      )
    },
    hydrate(entries) {
      records.clear()
      entries.forEach((record) => {
        records.set(record.cellIndex, record)
      })
    },
  }
}
