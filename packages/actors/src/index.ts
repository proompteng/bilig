import { assign, fromPromise, setup } from 'xstate'

interface BootstrapContext<Config, Session> {
  readonly config: Config | null
  readonly session: Session | null
  readonly error: string | null
  readonly shouldLoadSession: boolean
  readonly loadConfig: () => Promise<Config>
  readonly loadSession: (config: Config) => Promise<Session>
  readonly evaluateShouldLoadSession: (config: Config) => boolean
}

export interface BootstrapMachineInput<Config, Session> {
  readonly loadConfig: () => Promise<Config>
  readonly loadSession: (config: Config) => Promise<Session>
  readonly shouldLoadSession?: (config: Config) => boolean
}

interface LoadConfigActorInput<Config> {
  readonly loader: () => Promise<Config>
  readonly shouldLoadSession: (config: Config) => boolean
}

interface LoadSessionActorInput<Config, Session> {
  readonly config: Config
  readonly loader: (config: Config) => Promise<Session>
}

export function createBootstrapMachine<Config, Session>() {
  const loadConfigActor = fromPromise(async ({ input }: { input: LoadConfigActorInput<Config> }) => ({
    config: await input.loader(),
    shouldLoadSession: input.shouldLoadSession,
  }))
  const loadSessionActor = fromPromise(async ({ input }: { input: LoadSessionActorInput<Config, Session> }) => input.loader(input.config))

  return setup<
    BootstrapContext<Config, Session>,
    { type: 'retry' },
    {
      readonly loadConfig: typeof loadConfigActor
      readonly loadSession: typeof loadSessionActor
    },
    {},
    {},
    {},
    never,
    string,
    BootstrapMachineInput<Config, Session>
  >({
    actors: {
      loadConfig: loadConfigActor,
      loadSession: loadSessionActor,
    },
  }).createMachine({
    id: 'bootstrap',
    initial: 'loadingConfig',
    context: ({ input }) => ({
      config: null,
      session: null,
      error: null,
      shouldLoadSession: true,
      loadConfig: input.loadConfig,
      loadSession: input.loadSession,
      evaluateShouldLoadSession: input.shouldLoadSession ?? (() => true),
    }),
    states: {
      loadingConfig: {
        invoke: {
          src: 'loadConfig',
          input: ({ context }) => ({
            loader: context.loadConfig,
            shouldLoadSession: context.evaluateShouldLoadSession,
          }),
          onDone: {
            target: 'decideSession',
            actions: assign({
              config: ({ event }) => event.output.config,
              shouldLoadSession: ({ event }) => event.output.shouldLoadSession(event.output.config),
              error: () => null,
            }),
          },
          onError: {
            target: 'failed',
            actions: assign({
              error: ({ event }) => (event.error instanceof Error ? event.error.message : String(event.error)),
            }),
          },
        },
      },
      decideSession: {
        always: [
          {
            guard: ({ context }) => context.shouldLoadSession,
            target: 'loadingSession',
          },
          { target: 'ready' },
        ],
      },
      loadingSession: {
        invoke: {
          src: 'loadSession',
          input: ({ context }) => {
            if (context.config === null) {
              throw new Error('Cannot load session before config')
            }
            return {
              config: context.config,
              loader: context.loadSession,
            }
          },
          onDone: {
            target: 'ready',
            actions: assign({
              session: ({ event }) => event.output,
              error: () => null,
            }),
          },
          onError: {
            target: 'failed',
            actions: assign({
              error: ({ event }) => (event.error instanceof Error ? event.error.message : String(event.error)),
            }),
          },
        },
      },
      ready: {
        type: 'final',
      },
      failed: {
        on: {
          retry: {
            target: 'loadingConfig',
            actions: assign({
              config: () => null,
              session: () => null,
              error: () => null,
              shouldLoadSession: () => true,
            }),
          },
        },
      },
    },
  })
}

interface DocumentSupervisorContext {
  readonly documentId: string
  readonly browserSubscriberCount: number
  readonly lastKnownCursor: number | null
  readonly lastSnapshotCursor: number | null
  readonly lastOperation: string | null
  readonly lastError: string | null
}

type DocumentSupervisorEvent =
  | { type: 'browser.attached' }
  | { type: 'browser.detached' }
  | { type: 'cursor.updated'; cursor: number }
  | { type: 'snapshot.updated'; cursor: number }
  | { type: 'operation.recorded'; operation: string }
  | { type: 'error.raised'; message: string }
  | { type: 'reset' }

export function createDocumentSupervisorMachine(documentId: string) {
  return setup<DocumentSupervisorContext, DocumentSupervisorEvent>({}).createMachine({
    id: `document-supervisor:${documentId}`,
    initial: 'idle',
    context: {
      documentId,
      browserSubscriberCount: 0,
      lastKnownCursor: null,
      lastSnapshotCursor: null,
      lastOperation: null,
      lastError: null,
    },
    states: {
      idle: {
        on: {
          'browser.attached': {
            target: 'active',
            actions: assign({
              browserSubscriberCount: ({ context }) => context.browserSubscriberCount + 1,
              lastError: () => null,
            }),
          },
          'cursor.updated': {
            target: 'active',
            actions: assign({
              lastKnownCursor: ({ event }) => event.cursor,
              lastError: () => null,
            }),
          },
          'snapshot.updated': {
            target: 'active',
            actions: assign({
              lastSnapshotCursor: ({ event }) => event.cursor,
              lastError: () => null,
            }),
          },
          'operation.recorded': {
            target: 'active',
            actions: assign({
              lastOperation: ({ event }) => event.operation,
              lastError: () => null,
            }),
          },
          'error.raised': {
            target: 'degraded',
            actions: assign({
              lastError: ({ event }) => event.message,
            }),
          },
        },
      },
      active: {
        on: {
          'browser.attached': {
            actions: assign({
              browserSubscriberCount: ({ context }) => context.browserSubscriberCount + 1,
              lastError: () => null,
            }),
          },
          'browser.detached': {
            actions: assign({
              browserSubscriberCount: ({ context }) => Math.max(0, context.browserSubscriberCount - 1),
            }),
          },
          'cursor.updated': {
            actions: assign({
              lastKnownCursor: ({ event }) => event.cursor,
              lastError: () => null,
            }),
          },
          'snapshot.updated': {
            actions: assign({
              lastSnapshotCursor: ({ event }) => event.cursor,
              lastError: () => null,
            }),
          },
          'operation.recorded': {
            actions: assign({
              lastOperation: ({ event }) => event.operation,
              lastError: () => null,
            }),
          },
          'error.raised': {
            target: 'degraded',
            actions: assign({
              lastError: ({ event }) => event.message,
            }),
          },
        },
      },
      degraded: {
        on: {
          reset: {
            target: 'active',
            actions: assign({
              lastError: () => null,
            }),
          },
          'browser.attached': {
            target: 'active',
            actions: assign({
              browserSubscriberCount: ({ context }) => context.browserSubscriberCount + 1,
              lastError: () => null,
            }),
          },
          'cursor.updated': {
            target: 'active',
            actions: assign({
              lastKnownCursor: ({ event }) => event.cursor,
              lastError: () => null,
            }),
          },
          'snapshot.updated': {
            target: 'active',
            actions: assign({
              lastSnapshotCursor: ({ event }) => event.cursor,
              lastError: () => null,
            }),
          },
          'operation.recorded': {
            target: 'active',
            actions: assign({
              lastOperation: ({ event }) => event.operation,
              lastError: () => null,
            }),
          },
        },
      },
    },
  })
}
