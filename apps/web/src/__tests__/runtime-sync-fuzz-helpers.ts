import * as fc from 'fast-check'
import { createMemoryWorkbookLocalStoreFactory } from '@bilig/storage-browser'
import { ValueTag } from '@bilig/protocol'
import type { AuthoritativeWorkbookEventRecord } from '@bilig/zero-sync'
import { WorkbookWorkerRuntime } from '../worker-runtime.js'

export type RuntimeSyncAction =
  | { kind: 'local'; address: string; value: number }
  | { kind: 'submit' }
  | { kind: 'remote'; address: string; value: number }
  | { kind: 'ack' }

type ExpectedPendingStatus = 'local' | 'submitted' | 'rebased'

export type RuntimeSyncModel = {
  nextRevision: number
  authoritative: Map<string, number>
  pending: Array<{
    id: string
    address: string
    value: number
    status: ExpectedPendingStatus
  }>
}

export const runtimeSyncTrackedAddresses = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3'] as const

export const runtimeSyncActionArbitrary = fc.oneof<RuntimeSyncAction>(
  fc
    .record({
      address: fc.constantFrom(...runtimeSyncTrackedAddresses),
      value: fc.integer({ min: -50, max: 50 }),
    })
    .map((action) => Object.assign({ kind: 'local' as const }, action)),
  fc.constant({ kind: 'submit' as const }),
  fc
    .record({
      address: fc.constantFrom(...runtimeSyncTrackedAddresses),
      value: fc.integer({ min: -50, max: 50 }),
    })
    .map((action) => Object.assign({ kind: 'remote' as const }, action)),
  fc.constant({ kind: 'ack' as const }),
)

export async function createRuntimeSyncHarness(): Promise<{
  runtime: WorkbookWorkerRuntime
  model: RuntimeSyncModel
}> {
  const runtime = new WorkbookWorkerRuntime({
    localStoreFactory: createMemoryWorkbookLocalStoreFactory(),
  })
  await runtime.bootstrap({
    documentId: 'runtime-sync-fuzz',
    replicaId: 'browser:runtime-sync-fuzz',
    persistState: true,
  })
  return {
    runtime,
    model: {
      nextRevision: 0,
      authoritative: new Map(),
      pending: [],
    },
  }
}

export async function applyRuntimeSyncAction(
  runtime: WorkbookWorkerRuntime,
  model: RuntimeSyncModel,
  action: RuntimeSyncAction,
): Promise<void> {
  switch (action.kind) {
    case 'local': {
      const mutation = await runtime.enqueuePendingMutation({
        method: 'setCellValue',
        args: ['Sheet1', action.address, action.value],
      })
      model.pending.push({
        id: mutation.id,
        address: action.address,
        value: action.value,
        status: 'local',
      })
      return
    }
    case 'submit': {
      const target = model.pending.find((mutation) => mutation.status === 'local' || mutation.status === 'rebased')
      if (!target) {
        return
      }
      await runtime.markPendingMutationSubmitted(target.id)
      target.status = 'submitted'
      return
    }
    case 'remote': {
      model.nextRevision += 1
      await runtime.applyAuthoritativeEvents(
        [
          buildSetCellValueEvent({
            revision: model.nextRevision,
            address: action.address,
            value: action.value,
          }),
        ],
        model.nextRevision,
      )
      model.authoritative.set(action.address, action.value)
      model.pending.forEach((mutation) => {
        if (mutation.status === 'local') {
          mutation.status = 'rebased'
        }
      })
      return
    }
    case 'ack': {
      const target = model.pending.find((mutation) => mutation.status === 'submitted')
      if (!target) {
        return
      }
      model.nextRevision += 1
      await runtime.applyAuthoritativeEvents(
        [
          buildSetCellValueEvent({
            revision: model.nextRevision,
            address: target.address,
            value: target.value,
            clientMutationId: target.id,
          }),
        ],
        model.nextRevision,
      )
      model.authoritative.set(target.address, target.value)
      model.pending = model.pending.filter((mutation) => mutation.id !== target.id)
      model.pending.forEach((mutation) => {
        if (mutation.status === 'local') {
          mutation.status = 'rebased'
        }
      })
      return
    }
  }
}

export function assertRuntimeSyncState(runtime: WorkbookWorkerRuntime, model: RuntimeSyncModel): void {
  const overlayValues = new Map<string, number>()
  model.pending.forEach((mutation) => {
    overlayValues.set(mutation.address, mutation.value)
  })
  runtimeSyncTrackedAddresses.forEach((address) => {
    const expected = overlayValues.get(address) ?? model.authoritative.get(address)
    const actual = runtime.getCell('Sheet1', address).value
    if (expected === undefined) {
      expectEmptyValue(actual, address)
      return
    }
    if (actual.tag !== ValueTag.Number || actual.value !== expected) {
      throw new Error(`Runtime sync mismatch at ${address}: expected ${String(expected)}, received ${JSON.stringify(actual)}`)
    }
  })
  const runtimePending = runtime.listPendingMutations().map((mutation) => ({
    id: mutation.id,
    status: mutation.status,
  }))
  const modelPending = model.pending.map((mutation) => ({
    id: mutation.id,
    status: mutation.status,
  }))
  if (JSON.stringify(runtimePending) !== JSON.stringify(modelPending)) {
    throw new Error(`Pending mutation mismatch: expected ${JSON.stringify(modelPending)}, received ${JSON.stringify(runtimePending)}`)
  }
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

function expectEmptyValue(value: ReturnType<WorkbookWorkerRuntime['getCell']>['value'], address: string): void {
  if (value.tag !== ValueTag.Empty) {
    throw new Error(`Expected ${address} to be empty, received ${JSON.stringify(value)}`)
  }
}
