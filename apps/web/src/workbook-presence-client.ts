const WORKBOOK_PRESENCE_CLIENT_STORAGE_KEY = 'bilig:workbook-presence-client-id'
const WORKBOOK_PRESENCE_CLIENT_ID_PREFIX = 'presence:'

function createWorkbookPresenceClientId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `${WORKBOOK_PRESENCE_CLIENT_ID_PREFIX}${crypto.randomUUID()}`
  }
  return `${WORKBOOK_PRESENCE_CLIENT_ID_PREFIX}${Math.random().toString(36).slice(2)}`
}

function parseWorkbookPresenceClientId(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null
  }
  const trimmedValue = value.trim()
  return trimmedValue.startsWith(WORKBOOK_PRESENCE_CLIENT_ID_PREFIX) &&
    trimmedValue.length > WORKBOOK_PRESENCE_CLIENT_ID_PREFIX.length &&
    !/\s/.test(trimmedValue)
    ? trimmedValue
    : null
}

function removeStoredWorkbookPresenceClientId(): void {
  try {
    window.localStorage.removeItem(WORKBOOK_PRESENCE_CLIENT_STORAGE_KEY)
  } catch {
    // Ignore storage cleanup failures and keep presence usable.
  }
}

export function loadOrCreateWorkbookPresenceClientId(): string {
  if (typeof window === 'undefined') {
    return createWorkbookPresenceClientId()
  }
  try {
    const storedValue = parseWorkbookPresenceClientId(window.localStorage.getItem(WORKBOOK_PRESENCE_CLIENT_STORAGE_KEY))
    if (storedValue) {
      return storedValue
    }
    removeStoredWorkbookPresenceClientId()
    const nextValue = createWorkbookPresenceClientId()
    window.localStorage.setItem(WORKBOOK_PRESENCE_CLIENT_STORAGE_KEY, nextValue)
    return nextValue
  } catch {
    return createWorkbookPresenceClientId()
  }
}
