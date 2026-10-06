import type { ZeroConnectionState } from './worker-workbook-app-model.js'

export type WorkbookSaveState = 'loading' | 'read-only' | 'error' | 'saving' | 'local' | 'offline' | 'saved'

export function deriveWorkbookSaveState(input: {
  connectionStateName: ZeroConnectionState['name']
  runtimeReady: boolean
  remoteSyncAvailable: boolean
  zeroConfigured: boolean
  zeroHealthReady: boolean
  writesAllowed: boolean
  hasLocalMutationInFlight?: boolean
  pendingMutationSummary?: { readonly activeCount: number; readonly failedCount: number } | undefined
  failedPendingMutation?: unknown
}): WorkbookSaveState {
  if (!input.runtimeReady) return 'loading'
  if (!input.writesAllowed) return 'read-only'
  if (input.failedPendingMutation || (input.pendingMutationSummary?.failedCount ?? 0) > 0) return 'error'
  if (input.hasLocalMutationInFlight === true) return 'saving'
  if (!input.zeroConfigured) return 'local'
  if (input.connectionStateName === 'needs-auth' || input.connectionStateName === 'error') return 'error'
  if (input.connectionStateName === 'disconnected' || input.connectionStateName === 'closed') return 'offline'
  if ((input.pendingMutationSummary?.activeCount ?? 0) > 0) return 'saving'
  if (input.connectionStateName === 'connecting' || !input.remoteSyncAvailable || !input.zeroHealthReady) return 'local'
  return 'saved'
}
