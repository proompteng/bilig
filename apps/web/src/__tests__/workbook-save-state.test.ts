import { expect, it } from 'vitest'
import { deriveWorkbookStatusPresentation } from '../workbook-toolbar-state.js'

it.each([false, true])('does not claim an in-flight local write is saved with remote sync %s', (zeroConfigured) => {
  expect(
    deriveWorkbookStatusPresentation({
      connectionStateName: 'connected',
      runtimeReady: true,
      remoteSyncAvailable: zeroConfigured,
      zeroConfigured,
      zeroHealthReady: zeroConfigured,
      writesAllowed: true,
      hasLocalMutationInFlight: true,
      pendingMutationSummary: { activeCount: 0, failedCount: 0 },
    }),
  ).toMatchObject({ syncLabel: 'Saving…', tone: 'neutral' })
})
