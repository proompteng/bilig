import type { WorkbookSnapshot } from '@bilig/protocol'

export interface RecentWorkbook {
  readonly documentId: string
  readonly name: string
  readonly serverUrl?: string
}

function recentWorkbooksKey(userId: string): string {
  return `bilig:recent-workbooks:${encodeURIComponent(userId)}`
}

export function createBlankWorkbook(): WorkbookSnapshot {
  return { version: 1, workbook: { name: 'Untitled workbook' }, sheets: [{ id: 1, name: 'Sheet1', order: 0, cells: [] }] }
}

export function duplicateWorkbook(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const copy = structuredClone(snapshot)
  copy.workbook.name = `Copy of ${snapshot.workbook.name}`
  return copy
}

export function loadRecentWorkbooks(userId: string): RecentWorkbook[] {
  try {
    const stored: unknown = JSON.parse(localStorage.getItem(recentWorkbooksKey(userId)) ?? '[]')
    if (!Array.isArray(stored)) return []
    return stored
      .filter(
        (entry): entry is RecentWorkbook =>
          typeof entry === 'object' &&
          entry !== null &&
          typeof entry.documentId === 'string' &&
          typeof entry.name === 'string' &&
          (entry.serverUrl === undefined || typeof entry.serverUrl === 'string'),
      )
      .slice(0, 12)
  } catch {
    return []
  }
}

export function rememberWorkbook(userId: string, workbook: RecentWorkbook): void {
  try {
    localStorage.setItem(
      recentWorkbooksKey(userId),
      JSON.stringify([workbook, ...loadRecentWorkbooks(userId).filter((entry) => entry.documentId !== workbook.documentId)].slice(0, 12)),
    )
  } catch {
    // Browser storage availability must not block opening a workbook.
  }
}

export function workbookBackupFile(snapshot: WorkbookSnapshot): File {
  const name = snapshot.workbook.name.replace(/[\\/:*?"<>|]/g, '_') || 'Workbook'
  return new File([JSON.stringify(snapshot)], `${name}.bilig.json`, { type: 'application/vnd.bilig.workbook+json' })
}
