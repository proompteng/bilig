import { expect, it } from 'vitest'
import { deriveWorkbookSaveState } from '../workbook-save-state.js'

const ready = {
  connectionStateName: 'connected',
  runtimeReady: true,
  remoteSyncAvailable: true,
  zeroConfigured: true,
  zeroHealthReady: true,
  writesAllowed: true,
  pendingMutationSummary: { activeCount: 0, failedCount: 0 },
} as const

it.each([false, true])('does not claim a deferred write is saved with remote sync %s', (zeroConfigured) => {
  expect(deriveWorkbookSaveState({ ...ready, zeroConfigured, hasLocalMutationInFlight: true })).toBe('saving')
})

it.each([
  { input: {}, expected: 'saved' },
  { input: { runtimeReady: false }, expected: 'loading' },
  { input: { writesAllowed: false }, expected: 'read-only' },
  { input: { zeroConfigured: false }, expected: 'local' },
  { input: { connectionStateName: 'connecting', remoteSyncAvailable: false }, expected: 'local' },
  { input: { pendingMutationSummary: { activeCount: 2, failedCount: 0 } }, expected: 'saving' },
  { input: { pendingMutationSummary: { activeCount: 0, failedCount: 1 } }, expected: 'error' },
  { input: { connectionStateName: 'error' }, expected: 'error' },
  { input: { connectionStateName: 'disconnected', pendingMutationSummary: { activeCount: 2, failedCount: 0 } }, expected: 'offline' },
] satisfies { input: Partial<Parameters<typeof deriveWorkbookSaveState>[0]>; expected: string }[])(
  'derives $expected for $input',
  ({ input, expected }) => {
    expect(deriveWorkbookSaveState({ ...ready, ...input })).toBe(expected)
  },
)
