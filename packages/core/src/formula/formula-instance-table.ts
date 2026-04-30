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
  const records: Array<FormulaInstanceSnapshot | undefined> = []
  let recordCount = 0

  return {
    upsert(record) {
      if (records[record.cellIndex] === undefined) {
        recordCount += 1
      }
      records[record.cellIndex] = record
    },
    get(cellIndex) {
      return records[cellIndex]
    },
    delete(cellIndex) {
      if (records[cellIndex] === undefined) {
        return false
      }
      records[cellIndex] = undefined
      recordCount -= 1
      return true
    },
    clear() {
      records.length = 0
      recordCount = 0
    },
    list() {
      if (recordCount === 0) {
        return []
      }
      const snapshots: FormulaInstanceSnapshot[] = []
      for (let cellIndex = 0; cellIndex < records.length; cellIndex += 1) {
        const record = records[cellIndex]
        if (record !== undefined) {
          snapshots.push(record)
        }
      }
      return snapshots.toSorted(
        (left, right) =>
          left.sheetName.localeCompare(right.sheetName) || left.row - right.row || left.col - right.col || left.cellIndex - right.cellIndex,
      )
    },
    hydrate(entries) {
      records.length = 0
      recordCount = 0
      entries.forEach((record) => {
        if (records[record.cellIndex] === undefined) {
          recordCount += 1
        }
        records[record.cellIndex] = record
      })
    },
  }
}
