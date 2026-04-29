import type { BiligRuntimeConfig } from '@bilig/zero-sync'

export interface RuntimeConfig {
  documentId: string
  persistState: boolean
  currentUserId: string
  workbookAgentEnabled: boolean
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

export function resolveRuntimeConfig(config: BiligRuntimeConfig): RuntimeConfig {
  const searchParams = typeof window === 'undefined' ? new URLSearchParams() : new URLSearchParams(window.location.search)
  const explicitDocumentId = searchParams.get('document')

  if (explicitDocumentId) {
    return {
      documentId: explicitDocumentId,
      persistState: true,
      currentUserId: config.currentUserId,
      workbookAgentEnabled: config.workbookAgentEnabled === true,
    }
  }

  return {
    documentId: config.defaultDocumentId,
    persistState: config.persistState,
    currentUserId: config.currentUserId,
    workbookAgentEnabled: config.workbookAgentEnabled === true,
  }
}
