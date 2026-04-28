import type { SyncState } from '@bilig/protocol'
import { assign, fromCallback, sendTo, setup } from 'xstate'
import type { WorkbookWorkerStateSnapshot } from './worker-runtime.js'
import type { ZeroConnectionState } from './worker-workbook-app-model.js'
import {
  createWorkerRuntimeSessionController,
  type CreateWorkerRuntimeSessionInput,
  type WorkerHandle,
  type WorkerRuntimeSelection,
  type WorkerRuntimeSessionController,
  type WorkerRuntimeSessionPhase,
} from './runtime-session.js'

type ConnectionStateName = ZeroConnectionState['name']

export interface WorkerRuntimeMachineContext {
  readonly sessionInput: WorkerRuntimeMachineInput
  readonly persistState: boolean
  readonly runtimeResourceId: string
  readonly runtimeResourceVersion: number
  readonly runtimeState: WorkbookWorkerStateSnapshot | null
  readonly selection: WorkerRuntimeSelection
  readonly connectionStateName: ConnectionStateName
  readonly error: string | null
}

const NON_PERSISTED_SESSION_INPUT_KEYS = ['createSession', 'createWorker', 'fetchImpl', 'perfSession', 'zero'] as const

interface WorkerRuntimeMachineResources {
  readonly controller: WorkerRuntimeSessionController
  readonly handle: WorkerHandle
}

const workerRuntimeResourcesById = new Map<string, WorkerRuntimeMachineResources>()

type WorkerRuntimeMachineEvent =
  | { type: 'retry'; persistState?: boolean }
  | { type: 'error.clear' }
  | { type: 'selection.changed'; selection: WorkerRuntimeSelection }
  | { type: 'connection.changed'; connectionStateName: ConnectionStateName }
  | { type: 'session.ready'; controller: WorkerRuntimeSessionController }
  | { type: 'session.runtime'; runtimeState: WorkbookWorkerStateSnapshot }
  | { type: 'session.selection'; selection: WorkerRuntimeSelection }
  | { type: 'session.phase'; phase: WorkerRuntimeSessionPhase }
  | { type: 'session.error'; message: string }
  | { type: 'session.failed'; message: string }

export interface WorkerRuntimeMachineInput extends CreateWorkerRuntimeSessionInput {
  readonly connectionStateName?: ConnectionStateName
  readonly createSession?: (
    input: CreateWorkerRuntimeSessionInput,
    callbacks: Parameters<typeof createWorkerRuntimeSessionController>[1],
  ) => Promise<WorkerRuntimeSessionController>
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isWorkbookWorkerStateSnapshotValue(value: unknown): value is WorkbookWorkerStateSnapshot {
  return (
    isRecord(value) &&
    typeof value['workbookName'] === 'string' &&
    Array.isArray(value['sheetNames']) &&
    isRecord(value['metrics']) &&
    typeof value['syncState'] === 'string'
  )
}

function initialConnectionStateName(input: WorkerRuntimeMachineInput): ConnectionStateName {
  return input.connectionStateName ?? (input.zero ? 'connecting' : 'closed')
}

function mapConnectionStateToRuntimeSyncState(connectionStateName: ConnectionStateName, hasZero: boolean): SyncState | null {
  if (!hasZero) {
    return 'local-only'
  }
  switch (connectionStateName) {
    case 'connected':
      return 'live'
    case 'connecting':
      return 'syncing'
    case 'disconnected':
      return 'reconnecting'
    case 'needs-auth':
    case 'error':
    case 'closed':
      return 'local-only'
  }
}

function sameWorkerRuntimeSelection(left: WorkerRuntimeSelection, right: WorkerRuntimeSelection): boolean {
  return left.sheetName === right.sheetName && left.address === right.address
}

function resolveSteadySubstate(input: {
  hasZero: boolean
  connectionStateName: ConnectionStateName
}): 'localReady' | 'live' | 'syncing' | 'offline' {
  if (!input.hasZero) {
    return 'localReady'
  }
  switch (input.connectionStateName) {
    case 'connected':
      return 'live'
    case 'connecting':
      return 'syncing'
    case 'disconnected':
    case 'needs-auth':
    case 'error':
    case 'closed':
      return 'offline'
  }
}

function buildSessionCreateInput(input: WorkerRuntimeMachineInput): CreateWorkerRuntimeSessionInput {
  return {
    documentId: input.documentId,
    replicaId: input.replicaId,
    persistState: input.persistState,
    initialSelection: input.initialSelection,
    ...(input.perfSession ? { perfSession: input.perfSession } : {}),
    ...(input.zero ? { zero: input.zero } : {}),
    ...(input.fetchImpl ? { fetchImpl: input.fetchImpl } : {}),
    ...(input.createWorker ? { createWorker: input.createWorker } : {}),
  }
}

function createPersistableSessionInput(input: WorkerRuntimeMachineInput): WorkerRuntimeMachineInput {
  const sessionInput = { ...input }
  NON_PERSISTED_SESSION_INPUT_KEYS.forEach((key) => {
    if (key in input) {
      Object.defineProperty(sessionInput, key, {
        configurable: true,
        enumerable: false,
        value: input[key],
        writable: false,
      })
    }
  })
  return sessionInput
}

function runtimeResourceIdForInput(input: WorkerRuntimeMachineInput): string {
  return `${input.documentId}:${input.replicaId}`
}

export function getWorkerRuntimeController(context: WorkerRuntimeMachineContext): WorkerRuntimeSessionController | null {
  return workerRuntimeResourcesById.get(context.runtimeResourceId)?.controller ?? null
}

export function getWorkerRuntimeHandle(context: WorkerRuntimeMachineContext): WorkerHandle | null {
  return workerRuntimeResourcesById.get(context.runtimeResourceId)?.handle ?? null
}

function hasWorkerRuntimeController(context: WorkerRuntimeMachineContext): boolean {
  return getWorkerRuntimeController(context) !== null
}

function buildRuntimeSessionActorInput(context: WorkerRuntimeMachineContext): WorkerRuntimeMachineInput {
  const sessionInput = context.sessionInput
  return createPersistableSessionInput({
    documentId: sessionInput.documentId,
    replicaId: sessionInput.replicaId,
    persistState: context.persistState,
    initialSelection: context.selection,
    connectionStateName: context.connectionStateName,
    ...(sessionInput.createSession ? { createSession: sessionInput.createSession } : {}),
    ...(sessionInput.createWorker ? { createWorker: sessionInput.createWorker } : {}),
    ...(sessionInput.fetchImpl ? { fetchImpl: sessionInput.fetchImpl } : {}),
    ...(sessionInput.perfSession ? { perfSession: sessionInput.perfSession } : {}),
    ...(sessionInput.zero ? { zero: sessionInput.zero } : {}),
  })
}

function createInitialContext(input: WorkerRuntimeMachineInput): WorkerRuntimeMachineContext {
  return {
    sessionInput: createPersistableSessionInput(input),
    persistState: input.persistState,
    runtimeResourceId: runtimeResourceIdForInput(input),
    runtimeResourceVersion: 0,
    runtimeState: null,
    selection: input.initialSelection,
    connectionStateName: initialConnectionStateName(input),
    error: null,
  }
}

const storeRuntimeResourcesAction = ({
  context,
  event,
}: {
  context: WorkerRuntimeMachineContext
  event: WorkerRuntimeMachineEvent
}): void => {
  if (event.type !== 'session.ready') {
    return
  }
  workerRuntimeResourcesById.set(context.runtimeResourceId, {
    controller: event.controller,
    handle: event.controller.handle,
  })
}

const clearRuntimeResourcesAction = ({ context }: { context: WorkerRuntimeMachineContext }): void => {
  workerRuntimeResourcesById.delete(context.runtimeResourceId)
}

export function createWorkerRuntimeMachine() {
  const runtimeSessionActor = fromCallback(
    ({
      sendBack,
      receive,
      input,
    }: {
      sendBack: (event: WorkerRuntimeMachineEvent) => void
      receive: (listener: (event: WorkerRuntimeMachineEvent) => void) => void
      input: WorkerRuntimeMachineInput
    }) => {
      const createSession = input.createSession ?? createWorkerRuntimeSessionController
      let controller: WorkerRuntimeSessionController | null = null
      let disposed = false
      let pendingSelection = input.initialSelection
      let pendingConnectionStateName = initialConnectionStateName(input)

      const applyExternalSyncState = async (): Promise<void> => {
        if (!controller) {
          return
        }
        const value = await controller.invoke(
          'setExternalSyncState',
          mapConnectionStateToRuntimeSyncState(pendingConnectionStateName, Boolean(input.zero)),
        )
        if (!disposed && isWorkbookWorkerStateSnapshotValue(value)) {
          sendBack({ type: 'session.runtime', runtimeState: value })
        }
      }

      receive((event) => {
        if (event.type === 'selection.changed') {
          pendingSelection = event.selection
          if (!controller) {
            return
          }
          void (async () => {
            try {
              await controller.setSelection(event.selection)
            } catch (error) {
              if (!disposed) {
                sendBack({
                  type: 'session.error',
                  message: error instanceof Error ? error.message : String(error),
                })
              }
            }
          })()
          return
        }
        if (event.type === 'connection.changed') {
          pendingConnectionStateName = event.connectionStateName
          if (!controller) {
            return
          }
          void (async () => {
            try {
              await applyExternalSyncState()
            } catch (error) {
              if (!disposed) {
                sendBack({
                  type: 'session.error',
                  message: error instanceof Error ? error.message : String(error),
                })
              }
            }
          })()
        }
      })

      void (async () => {
        try {
          const createdController = await createSession(
            buildSessionCreateInput({
              ...input,
              initialSelection: pendingSelection,
            }),
            {
              onRuntimeState(runtimeState) {
                sendBack({ type: 'session.runtime', runtimeState })
              },
              onSelection(selection) {
                if (!sameWorkerRuntimeSelection(selection, pendingSelection)) {
                  return
                }
                pendingSelection = selection
                sendBack({ type: 'session.selection', selection })
              },
              onPhase(phase) {
                sendBack({ type: 'session.phase', phase })
              },
              onError(message) {
                sendBack({ type: 'session.error', message })
              },
            },
          )
          if (disposed) {
            createdController.dispose()
            return
          }
          controller = createdController
          sendBack({ type: 'session.ready', controller: createdController })
          try {
            await applyExternalSyncState()
          } catch (error) {
            if (!disposed) {
              sendBack({
                type: 'session.error',
                message: error instanceof Error ? error.message : String(error),
              })
            }
          }
          if (
            pendingSelection.sheetName !== createdController.selection.sheetName ||
            pendingSelection.address !== createdController.selection.address
          ) {
            try {
              await createdController.setSelection(pendingSelection)
            } catch (error) {
              if (!disposed) {
                sendBack({
                  type: 'session.error',
                  message: error instanceof Error ? error.message : String(error),
                })
              }
            }
          }
        } catch (error) {
          if (!disposed) {
            sendBack({
              type: 'session.failed',
              message: error instanceof Error ? error.message : String(error),
            })
          }
        }
      })()

      return () => {
        disposed = true
        controller?.dispose()
      }
    },
  )

  return setup<
    WorkerRuntimeMachineContext,
    WorkerRuntimeMachineEvent,
    {
      readonly runtimeSession: typeof runtimeSessionActor
    },
    {},
    {},
    {},
    never,
    string,
    WorkerRuntimeMachineInput
  >({
    actors: {
      runtimeSession: runtimeSessionActor,
    },
  }).createMachine({
    id: 'workerRuntime',
    initial: 'active',
    context: ({ input }) => createInitialContext(input),
    states: {
      active: {
        invoke: {
          id: 'runtimeSession',
          src: 'runtimeSession',
          input: ({ context }) => buildRuntimeSessionActorInput(context),
        },
        on: {
          retry: {
            target: '#workerRuntime.active',
            reenter: true,
            actions: [
              clearRuntimeResourcesAction,
              assign({
                persistState: ({ context, event }) => event['persistState'] ?? context.persistState,
                runtimeResourceVersion: ({ context }) => context.runtimeResourceVersion + 1,
                runtimeState: () => null,
                error: () => null,
              }),
            ],
          },
          'error.clear': {
            actions: assign({
              error: () => null,
            }),
          },
          'selection.changed': {
            actions: [
              assign({
                selection: ({ event }) => event['selection'],
              }),
              sendTo('runtimeSession', ({ event }) => event),
            ],
          },
          'connection.changed': {
            actions: [
              assign({
                connectionStateName: ({ event }) => event['connectionStateName'],
              }),
              sendTo('runtimeSession', ({ event }) => event),
            ],
          },
          'session.runtime': {
            actions: assign({
              runtimeState: ({ event }) => event['runtimeState'],
            }),
          },
          'session.selection': {
            actions: assign({
              selection: ({ event }) => event['selection'],
            }),
          },
          'session.error': {
            actions: assign({
              error: ({ event }) => event['message'],
            }),
          },
          'session.failed': {
            target: 'failed',
            actions: [
              clearRuntimeResourcesAction,
              assign({
                error: ({ event }) => event['message'],
                runtimeResourceVersion: ({ context }) => context.runtimeResourceVersion + 1,
                runtimeState: () => null,
              }),
            ],
          },
          'session.ready': [
            {
              guard: ({ context }) =>
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === 'live',
              target: '.live',
              actions: [
                storeRuntimeResourcesAction,
                assign({
                  runtimeResourceVersion: ({ context }) => context.runtimeResourceVersion + 1,
                  runtimeState: ({ event }) => event['controller'].runtimeState,
                  selection: ({ event }) => event['controller'].selection,
                  error: () => null,
                }),
              ],
            },
            {
              guard: ({ context }) =>
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === 'syncing',
              target: '.syncing',
              actions: [
                storeRuntimeResourcesAction,
                assign({
                  runtimeResourceVersion: ({ context }) => context.runtimeResourceVersion + 1,
                  runtimeState: ({ event }) => event['controller'].runtimeState,
                  selection: ({ event }) => event['controller'].selection,
                  error: () => null,
                }),
              ],
            },
            {
              guard: ({ context }) =>
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === 'offline',
              target: '.offline',
              actions: [
                storeRuntimeResourcesAction,
                assign({
                  runtimeResourceVersion: ({ context }) => context.runtimeResourceVersion + 1,
                  runtimeState: ({ event }) => event['controller'].runtimeState,
                  selection: ({ event }) => event['controller'].selection,
                  error: () => null,
                }),
              ],
            },
            {
              target: '.localReady',
              actions: [
                storeRuntimeResourcesAction,
                assign({
                  runtimeResourceVersion: ({ context }) => context.runtimeResourceVersion + 1,
                  runtimeState: ({ event }) => event['controller'].runtimeState,
                  selection: ({ event }) => event['controller'].selection,
                  error: () => null,
                }),
              ],
            },
          ],
          'session.phase': [
            {
              guard: ({ event }) => event['phase'] === 'hydratingLocal',
              target: '.hydratingLocal',
            },
            {
              guard: ({ event }) => event['phase'] === 'syncing',
              target: '.syncing',
            },
            {
              guard: ({ event }) => event['phase'] === 'reconciling',
              target: '.reconciling',
            },
            {
              guard: ({ event }) => event['phase'] === 'recovering',
              target: '.recovering',
            },
            {
              guard: ({ context, event }) =>
                event['phase'] === 'steady' &&
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === 'live',
              target: '.live',
            },
            {
              guard: ({ context, event }) =>
                event['phase'] === 'steady' &&
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === 'syncing',
              target: '.syncing',
            },
            {
              guard: ({ context, event }) =>
                event['phase'] === 'steady' &&
                resolveSteadySubstate({
                  hasZero: Boolean(context.sessionInput.zero),
                  connectionStateName: context.connectionStateName,
                }) === 'offline',
              target: '.offline',
            },
            {
              guard: ({ event }) => event['phase'] === 'steady',
              target: '.localReady',
            },
          ],
        },
        initial: 'booting',
        states: {
          booting: {},
          hydratingLocal: {},
          syncing: {
            on: {
              'connection.changed': [
                {
                  guard: ({ context, event }) =>
                    hasWorkerRuntimeController(context) &&
                    resolveSteadySubstate({
                      hasZero: Boolean(context.sessionInput.zero),
                      connectionStateName: event['connectionStateName'],
                    }) === 'live',
                  target: '#workerRuntime.active.live',
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
                {
                  guard: ({ context, event }) =>
                    hasWorkerRuntimeController(context) &&
                    resolveSteadySubstate({
                      hasZero: Boolean(context.sessionInput.zero),
                      connectionStateName: event['connectionStateName'],
                    }) === 'offline',
                  target: '#workerRuntime.active.offline',
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
                {
                  guard: ({ context, event }) =>
                    hasWorkerRuntimeController(context) &&
                    resolveSteadySubstate({
                      hasZero: Boolean(context.sessionInput.zero),
                      connectionStateName: event['connectionStateName'],
                    }) === 'localReady',
                  target: '#workerRuntime.active.localReady',
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
                {
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
              ],
            },
          },
          localReady: {
            on: {
              'connection.changed': [
                {
                  guard: ({ event }) => event['connectionStateName'] === 'connected',
                  target: '#workerRuntime.active.live',
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
                {
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
              ],
            },
          },
          live: {
            on: {
              'connection.changed': [
                {
                  guard: ({ event }) => event['connectionStateName'] === 'connected',
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
                {
                  guard: ({ event }) => event['connectionStateName'] === 'connecting',
                  target: '#workerRuntime.active.syncing',
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
                {
                  target: '#workerRuntime.active.offline',
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
              ],
            },
          },
          offline: {
            on: {
              'connection.changed': [
                {
                  guard: ({ event }) => event['connectionStateName'] === 'connected',
                  target: '#workerRuntime.active.live',
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
                {
                  guard: ({ event }) => event['connectionStateName'] === 'connecting',
                  target: '#workerRuntime.active.syncing',
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
                {
                  actions: [
                    assign({
                      connectionStateName: ({ event }) => event['connectionStateName'],
                    }),
                    sendTo('runtimeSession', ({ event }) => event),
                  ],
                },
              ],
            },
          },
          reconciling: {},
          recovering: {},
        },
      },
      failed: {
        on: {
          'error.clear': {
            actions: assign({
              error: () => null,
            }),
          },
          retry: {
            target: 'active',
            actions: [
              clearRuntimeResourcesAction,
              assign({
                persistState: ({ context, event }) => event['persistState'] ?? context.persistState,
                runtimeResourceVersion: ({ context }) => context.runtimeResourceVersion + 1,
                runtimeState: () => null,
                error: () => null,
              }),
            ],
          },
        },
      },
    },
  })
}
