import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import { ValueTag } from '@bilig/protocol'
import type { AuthoritativeWorkbookEventRecord } from '@bilig/zero-sync'
import { runScheduledProperty } from '@bilig/test-fuzz'
import { WorkbookWorkerRuntime } from '../worker-runtime.js'
import { createMemoryWorkbookLocalStoreFactory } from '@bilig/storage-browser'
import { runtimeSyncTrackedAddresses } from './runtime-sync-fuzz-helpers.js'

type RuntimeSyncMutationKey = 'alpha' | 'beta' | 'gamma'
type RuntimeSyncScheduledAction =
  | { kind: 'enqueue'; key: RuntimeSyncMutationKey; address: string; value: number }
  | { kind: 'submit'; key: RuntimeSyncMutationKey }
  | { kind: 'ack'; key: RuntimeSyncMutationKey; revision: number }
  | { kind: 'remote'; address: string; value: number; revision: number }

type RuntimeSyncScheduledMutation = {
  key: RuntimeSyncMutationKey
  id: string
  address: string
  value: number
  status: 'local' | 'submitted' | 'rebased' | 'acked'
}

type RuntimeSyncScheduledModel = {
  authoritative: Map<string, number>
  mutations: RuntimeSyncScheduledMutation[]
}

describe('runtime sync scheduled fuzz', () => {
  it('should keep runtime state coherent when scheduler-chosen local and authoritative actions are applied in arbitrary order', async () => {
    await runScheduledProperty({
      suite: 'web/runtime-sync/scheduled-order-parity',
      arbitrary: fc.array(runtimeSyncScheduledActionArbitrary, { minLength: 4, maxLength: 12 }),
      predicate: async ({ scheduler, value: actions }) => {
        const runtime = new WorkbookWorkerRuntime({
          localStoreFactory: createMemoryWorkbookLocalStoreFactory(),
        })
        await runtime.bootstrap({
          documentId: 'runtime-sync-scheduled-fuzz',
          replicaId: 'browser:runtime-sync-scheduled-fuzz',
          persistState: true,
        })
        const model: RuntimeSyncScheduledModel = {
          authoritative: new Map(),
          mutations: [],
        }
        let operationChain = Promise.resolve()
        const runAction = scheduler.scheduleFunction(async (action: RuntimeSyncScheduledAction): Promise<void> => {
          operationChain = operationChain.then(async () => {
            await applyRuntimeSyncScheduledAction(runtime, model, action)
            assertRuntimeSyncScheduledState(runtime, model)
            return undefined
          })
          await operationChain
        })

        try {
          const actionPromises = actions.map((action) => runAction(action))
          await scheduler.waitFor(Promise.all(actionPromises))
        } finally {
          runtime.dispose()
        }
      },
    })
  })
})

// Helpers

const runtimeSyncMutationKeys = ['alpha', 'beta', 'gamma'] as const

const runtimeSyncScheduledActionArbitrary = fc.oneof<RuntimeSyncScheduledAction>(
  fc
    .record({
      key: fc.constantFrom<RuntimeSyncMutationKey>(...runtimeSyncMutationKeys),
      address: fc.constantFrom(...runtimeSyncTrackedAddresses),
      value: fc.integer({ min: -50, max: 50 }),
    })
    .map((action) => ({
      kind: 'enqueue' as const,
      key: action.key,
      address: action.address,
      value: action.value,
    })),
  fc.constantFrom<RuntimeSyncMutationKey>(...runtimeSyncMutationKeys).map((key) => ({
    kind: 'submit' as const,
    key,
  })),
  fc
    .record({
      key: fc.constantFrom<RuntimeSyncMutationKey>(...runtimeSyncMutationKeys),
      revision: fc.integer({ min: 1, max: 20 }),
    })
    .map((action) => ({
      kind: 'ack' as const,
      key: action.key,
      revision: action.revision,
    })),
  fc
    .record({
      address: fc.constantFrom(...runtimeSyncTrackedAddresses),
      value: fc.integer({ min: -50, max: 50 }),
      revision: fc.integer({ min: 1, max: 20 }),
    })
    .map((action) => ({
      kind: 'remote' as const,
      address: action.address,
      value: action.value,
      revision: action.revision,
    })),
)

async function applyRuntimeSyncScheduledAction(
  runtime: WorkbookWorkerRuntime,
  model: RuntimeSyncScheduledModel,
  action: RuntimeSyncScheduledAction,
): Promise<void> {
  switch (action.kind) {
    case 'enqueue': {
      if (model.mutations.some((mutation) => mutation.key === action.key)) {
        return
      }
      const mutation = await runtime.enqueuePendingMutation({
        method: 'setCellValue',
        args: ['Sheet1', action.address, action.value],
      })
      model.mutations.push({
        key: action.key,
        id: mutation.id,
        address: action.address,
        value: action.value,
        status: 'local',
      })
      return
    }
    case 'submit': {
      const mutation = model.mutations.find((entry) => entry.key === action.key)
      if (!mutation || mutation.status === 'acked') {
        return
      }
      await runtime.markPendingMutationSubmitted(mutation.id)
      if (mutation.status === 'local' || mutation.status === 'rebased') {
        mutation.status = 'submitted'
      }
      return
    }
    case 'ack': {
      const mutation = model.mutations.find((entry) => entry.key === action.key)
      if (!mutation || mutation.status === 'acked') {
        return
      }
      await runtime.applyAuthoritativeEvents(
        [
          buildSetCellValueEvent({
            revision: action.revision,
            address: mutation.address,
            value: mutation.value,
            clientMutationId: mutation.id,
          }),
        ],
        action.revision,
      )
      model.authoritative.set(mutation.address, mutation.value)
      mutation.status = 'acked'
      markRemainingLocalMutationsRebased(model, mutation.id)
      return
    }
    case 'remote':
      await runtime.applyAuthoritativeEvents(
        [
          buildSetCellValueEvent({
            revision: action.revision,
            address: action.address,
            value: action.value,
          }),
        ],
        action.revision,
      )
      model.authoritative.set(action.address, action.value)
      markRemainingLocalMutationsRebased(model)
      return
  }
}

function markRemainingLocalMutationsRebased(model: RuntimeSyncScheduledModel, absorbedMutationId?: string): void {
  model.mutations.forEach((mutation) => {
    if (mutation.id === absorbedMutationId || mutation.status !== 'local') {
      return
    }
    mutation.status = 'rebased'
  })
}

function assertRuntimeSyncScheduledState(runtime: WorkbookWorkerRuntime, model: RuntimeSyncScheduledModel): void {
  const overlayValues = new Map<string, number>()
  model.mutations.forEach((mutation) => {
    if (mutation.status !== 'acked') {
      overlayValues.set(mutation.address, mutation.value)
    }
  })

  runtimeSyncTrackedAddresses.forEach((address) => {
    const expected = overlayValues.get(address) ?? model.authoritative.get(address)
    const actual = runtime.getCell('Sheet1', address).value
    if (expected === undefined) {
      if (actual.tag !== ValueTag.Empty) {
        throw new Error(`Expected ${address} to be empty, received ${JSON.stringify(actual)}`)
      }
      return
    }
    expect(actual).toEqual({ tag: ValueTag.Number, value: expected })
  })

  expect(
    runtime.listPendingMutations().map((mutation) => ({
      id: mutation.id,
      status: mutation.status,
    })),
  ).toEqual(
    model.mutations
      .filter((mutation) => mutation.status !== 'acked')
      .map((mutation) => ({
        id: mutation.id,
        status: mutation.status,
      })),
  )

  expect(
    runtime.listMutationJournalEntries().map((mutation) => ({
      id: mutation.id,
      status: mutation.status,
    })),
  ).toEqual(
    model.mutations.map((mutation) => ({
      id: mutation.id,
      status: mutation.status,
    })),
  )
}

function buildSetCellValueEvent(input: {
  revision: number
  address: string
  value: number
  clientMutationId?: string | null
}): AuthoritativeWorkbookEventRecord {
  return {
    revision: input.revision,
    clientMutationId: input.clientMutationId ?? null,
    payload: {
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: input.address,
      value: input.value,
    },
  }
}
