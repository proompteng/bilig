import type { BiligRuntimeConfig, BiligZeroQueryContext } from '@bilig/zero-sync'

export interface RuntimeConfig {
  documentId: string
  persistState: boolean
  currentUserId: string
  workbookAgentEnabled: boolean
}

export function createLocalOnlyRuntimeConfig(currentUserId = 'local:user'): BiligRuntimeConfig {
  return {
    zeroCacheUrl: '/zero',
    defaultDocumentId: 'local-workbook',
    persistState: true,
    currentUserId,
    workbookAgentEnabled: false,
  }
}

export function normalizeRuntimeConfigUserId<T extends { currentUserId: string }>(
  config: T,
  session: {
    readonly userId: string
  },
): T {
  if (config.currentUserId === session.userId) {
    return config
  }
  return {
    ...config,
    currentUserId: session.userId,
  }
}

export function createZeroQueryContext(session: { readonly userId: string }): BiligZeroQueryContext {
  return {
    userID: session.userId,
  }
}

function resolvePersistState(configuredPersistState: boolean, searchParams: URLSearchParams): boolean {
  const explicitPersistState = searchParams.get('persist')
  if (explicitPersistState === null || explicitPersistState.length === 0) {
    return configuredPersistState
  }
  if (explicitPersistState === '0' || explicitPersistState === 'false') {
    return false
  }
  if (explicitPersistState === '1' || explicitPersistState === 'true') {
    return true
  }
  throw new Error(`persist query parameter must be "1", "true", "0", or "false" when set, got ${explicitPersistState}`)
}

export function resolveRuntimeConfig(config: BiligRuntimeConfig): RuntimeConfig {
  const searchParams = typeof window === 'undefined' ? new URLSearchParams() : new URLSearchParams(window.location.search)
  const explicitDocumentId = searchParams.get('document')
  const persistState = resolvePersistState(config.persistState, searchParams)

  if (explicitDocumentId) {
    return {
      documentId: explicitDocumentId,
      persistState,
      currentUserId: config.currentUserId,
      workbookAgentEnabled: config.workbookAgentEnabled === true,
    }
  }

  return {
    documentId: config.defaultDocumentId,
    persistState,
    currentUserId: config.currentUserId,
    workbookAgentEnabled: config.workbookAgentEnabled === true,
  }
}

export function resolveRemoteSyncEnabled(env: { readonly DEV?: boolean; readonly VITE_BILIG_REMOTE_SYNC?: string | undefined }): boolean {
  const configured = env.VITE_BILIG_REMOTE_SYNC
  const searchParams = typeof window === 'undefined' ? new URLSearchParams() : new URLSearchParams(window.location.search)
  const explicitPersistState = searchParams.get('persist')
  if (configured === undefined || configured.length === 0) {
    return env.DEV !== true && explicitPersistState !== '0' && explicitPersistState !== 'false'
  }
  if (configured === '1' || configured === 'true') {
    return explicitPersistState !== '0' && explicitPersistState !== 'false'
  }
  if (configured === '0' || configured === 'false') {
    return false
  }
  throw new Error(`VITE_BILIG_REMOTE_SYNC must be "1", "true", "0", or "false" when set, got ${configured}`)
}
