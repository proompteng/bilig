const WORKBOOK_PRESENCE_CLIENT_STORAGE_KEY = 'bilig:workbook-presence-client-id'

function createWorkbookPresenceClientId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `presence:${crypto.randomUUID()}`
  }
  return `presence:${Math.random().toString(36).slice(2)}`
}

export function loadOrCreateWorkbookPresenceClientId(): string {
  if (typeof window === 'undefined') {
    return createWorkbookPresenceClientId()
  }
  try {
    const storedValue = window.localStorage.getItem(WORKBOOK_PRESENCE_CLIENT_STORAGE_KEY)
    if (typeof storedValue === 'string' && storedValue.trim().length > 0) {
      return storedValue
    }
    const nextValue = createWorkbookPresenceClientId()
    window.localStorage.setItem(WORKBOOK_PRESENCE_CLIENT_STORAGE_KEY, nextValue)
    return nextValue
  } catch {
    return createWorkbookPresenceClientId()
  }
}
