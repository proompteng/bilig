import type { WorkbookAxisEntrySnapshot } from '@bilig/protocol'

interface SheetColumnInfo {
  index: number
  size: number | null
  hidden: boolean
}

interface SheetRowInfo {
  index: number
  size: number | null
  hidden: boolean
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function toPixelSize(value: number | undefined, unit: 'pt' | 'ch'): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    return null
  }
  if (unit === 'pt') {
    return Math.round((value * 96) / 72)
  }
  return Math.round(value * 8 + 5)
}

function toPositivePixelSize(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? Math.round(value) : null
}

export function buildColumnEntries(columns: unknown[] | undefined): WorkbookAxisEntrySnapshot[] | undefined {
  if (!Array.isArray(columns) || columns.length === 0) {
    return undefined
  }
  const entries: SheetColumnInfo[] = []
  columns.forEach((entry, index) => {
    if (!isRecord(entry)) {
      return
    }
    const size =
      typeof entry['wpx'] === 'number'
        ? toPositivePixelSize(entry['wpx'])
        : typeof entry['wch'] === 'number'
          ? toPixelSize(entry['wch'], 'ch')
          : null
    const hidden = entry['hidden'] === true
    if (size === null && !hidden) {
      return
    }
    entries.push({ index, size, hidden })
  })
  if (entries.length === 0) {
    return undefined
  }
  return entries.map(({ index, size, hidden }) => {
    const snapshot: WorkbookAxisEntrySnapshot = {
      id: `col:${index}`,
      index,
    }
    if (size !== null) {
      snapshot.size = size
    }
    if (hidden) {
      snapshot.hidden = true
    }
    return snapshot
  })
}

export function buildRowEntries(rows: unknown[] | undefined): WorkbookAxisEntrySnapshot[] | undefined {
  if (!Array.isArray(rows) || rows.length === 0) {
    return undefined
  }
  const entries: SheetRowInfo[] = []
  rows.forEach((entry, index) => {
    if (!isRecord(entry)) {
      return
    }
    const size =
      typeof entry['hpx'] === 'number'
        ? toPositivePixelSize(entry['hpx'])
        : typeof entry['hpt'] === 'number'
          ? toPixelSize(entry['hpt'], 'pt')
          : null
    const hidden = entry['hidden'] === true
    if (size === null && !hidden) {
      return
    }
    entries.push({ index, size, hidden })
  })
  if (entries.length === 0) {
    return undefined
  }
  return entries.map(({ index, size, hidden }) => {
    const snapshot: WorkbookAxisEntrySnapshot = {
      id: `row:${index}`,
      index,
    }
    if (size !== null) {
      snapshot.size = size
    }
    if (hidden) {
      snapshot.hidden = true
    }
    return snapshot
  })
}
