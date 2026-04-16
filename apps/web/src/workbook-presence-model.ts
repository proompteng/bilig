export const WORKBOOK_PRESENCE_HEARTBEAT_MS = 10_000
export const WORKBOOK_PRESENCE_STALE_AFTER_MS = 25_000
export const WORKBOOK_PRESENCE_STALE_TICK_MS = 5_000

export interface WorkbookPresenceCoarseRow {
  readonly sessionId: string
  readonly userId: string
  readonly presenceClientId: string | null
  readonly sheetId: number | null
  readonly sheetName: string | null
  readonly address: string | null
  readonly selectionJson: unknown
  readonly updatedAt: number
}

export interface WorkbookCollaboratorPresence {
  readonly sessionId: string
  readonly userId: string
  readonly label: string
  readonly initials: string
  readonly toneIndex: number
  readonly sheetName: string
  readonly address: string
  readonly updatedAt: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeWorkbookPresenceCoarseRow(value: unknown): WorkbookPresenceCoarseRow | null {
  if (!isRecord(value)) {
    return null
  }
  const sessionId = value['sessionId']
  const userId = value['userId']
  const presenceClientId = value['presenceClientId']
  const sheetId = value['sheetId']
  const sheetName = value['sheetName']
  const address = value['address']
  const updatedAt = value['updatedAt']
  if (typeof sessionId !== 'string' || typeof userId !== 'string' || typeof updatedAt !== 'number') {
    return null
  }
  return {
    sessionId,
    userId,
    presenceClientId: typeof presenceClientId === 'string' ? presenceClientId : null,
    sheetId: typeof sheetId === 'number' ? sheetId : null,
    sheetName: typeof sheetName === 'string' ? sheetName : null,
    address: typeof address === 'string' ? address : null,
    selectionJson: value['selectionJson'],
    updatedAt,
  }
}

export function normalizeWorkbookPresenceRows(value: unknown): readonly WorkbookPresenceCoarseRow[] {
  if (!Array.isArray(value)) {
    return []
  }
  return value.flatMap((entry) => {
    const row = normalizeWorkbookPresenceCoarseRow(entry)
    return row ? [row] : []
  })
}

function titleCaseSegments(value: string): string {
  return value
    .split(/[._-]+/u)
    .filter((segment) => segment.length > 0)
    .map((segment) => segment[0]!.toUpperCase() + segment.slice(1))
    .join(' ')
}

export function formatWorkbookCollaboratorLabel(userId: string): string {
  if (userId.startsWith('guest:')) {
    return `Guest ${userId.slice(-4).toUpperCase()}`
  }
  if (userId.includes('@')) {
    return titleCaseSegments(userId.split('@', 1)[0] ?? userId)
  }
  return titleCaseSegments(userId)
}

function formatWorkbookCollaboratorInitials(label: string): string {
  const words = label
    .split(/\s+/u)
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
  if (words.length === 0) {
    return '?'
  }
  if (words.length === 1) {
    return words[0]!.slice(0, 2).toUpperCase()
  }
  return `${words[0]![0] ?? ''}${words[1]![0] ?? ''}`.toUpperCase()
}

function hashToneIndex(value: string): number {
  let hash = 0
  for (const character of value) {
    hash = (hash * 31 + character.charCodeAt(0)) >>> 0
  }
  return hash % 8
}

export function selectActiveWorkbookCollaborators(input: {
  readonly rows: readonly WorkbookPresenceCoarseRow[]
  readonly currentUserId: string
  readonly currentPresenceClientId: string
  readonly currentSessionId: string
  readonly knownSheetNames: readonly string[]
  readonly now: number
  readonly staleAfterMs?: number
}): readonly WorkbookCollaboratorPresence[] {
  const staleAfterMs = input.staleAfterMs ?? WORKBOOK_PRESENCE_STALE_AFTER_MS
  const knownSheets = new Set(input.knownSheetNames)
  return input.rows
    .filter((row) => !row.userId.startsWith('guest:'))
    .filter((row) => row.userId !== input.currentUserId)
    .filter((row) => row.presenceClientId !== input.currentPresenceClientId)
    .filter((row) => row.sessionId !== input.currentSessionId)
    .filter((row) => row.updatedAt >= input.now - staleAfterMs)
    .filter((row) => typeof row.sheetName === 'string' && knownSheets.has(row.sheetName))
    .map((row) => {
      const label = formatWorkbookCollaboratorLabel(row.userId)
      return {
        sessionId: row.sessionId,
        userId: row.userId,
        label,
        initials: formatWorkbookCollaboratorInitials(label),
        toneIndex: hashToneIndex(row.userId),
        sheetName: row.sheetName!,
        address: row.address ?? 'A1',
        updatedAt: row.updatedAt,
      } satisfies WorkbookCollaboratorPresence
    })
    .toSorted((left, right) => {
      if (left.updatedAt !== right.updatedAt) {
        return right.updatedAt - left.updatedAt
      }
      return left.label.localeCompare(right.label)
    })
}
