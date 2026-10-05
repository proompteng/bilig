import React, { Suspense, lazy, useMemo } from 'react'
import ReactDOM from 'react-dom/client'
import { useActorRef, useSelector } from '@xstate/react'
import { ZeroProvider } from '@rocicorp/zero/react'
import { createBootstrapMachine } from '@bilig/actors'
import { loadRuntimeConfig, mutators, schema } from '@bilig/zero-sync'
import type { RuntimeSession } from '@bilig/contracts'
import { App } from './App.js'
import {
  createLocalOnlyRuntimeConfig,
  createZeroQueryContext,
  normalizeRuntimeConfigUserId,
  resolveRemoteSyncEnabled,
  resolveRuntimeConfig,
  type RuntimeConfig,
} from './runtime-config'
import { resolveWebEntryRoute } from './root-route.js'
import { createLocalOnlyRuntimeSession, loadRuntimeSession } from './session'
import { resolveZeroCacheUrl, ZERO_CONNECT_MAX_HEADER_LENGTH } from './zero-connection'
import type { ZeroConnectionState } from './worker-workbook-app-model.js'
import type { BiligRuntimeConfig } from '@bilig/zero-sync'
import { getOrCreateReactRoot } from './react-root-lifecycle.js'

const IsolatedWorkbookPaneRendererRoute = lazy(async () => {
  const module = await import('./IsolatedWorkbookPaneRendererRoute.js')
  return {
    default: module.IsolatedWorkbookPaneRendererRoute,
  }
})

const root = getOrCreateReactRoot({
  container: document.getElementById('root')!,
  createRoot: ReactDOM.createRoot,
  hot: import.meta.hot,
})
const entryRoute = resolveWebEntryRoute(window.location.pathname)
const remoteSyncEnabled = resolveRemoteSyncEnabled(import.meta.env)
const LOCAL_ONLY_CONNECTION_STATE: ZeroConnectionState = {
  name: 'closed',
  reason: 'Remote sync disabled for this environment',
}

declare global {
  interface Window {
    __biligRuntimeBuild?: {
      readonly app: 'bilig-web'
      readonly buildKind: 'development' | 'production'
      readonly dev: boolean
      readonly entryRoute: ReturnType<typeof resolveWebEntryRoute>
      readonly mode: string
      readonly prod: boolean
      readonly remoteSyncEnabled: boolean
    }
  }
}

window.__biligRuntimeBuild = {
  app: 'bilig-web',
  buildKind: import.meta.env.PROD ? 'production' : 'development',
  dev: import.meta.env.DEV,
  entryRoute,
  mode: import.meta.env.MODE,
  prod: import.meta.env.PROD,
  remoteSyncEnabled,
}

interface BootstrapConfig {
  readonly rawConfig: BiligRuntimeConfig
  readonly runtimeConfig: RuntimeConfig
}

const bootstrapMachine = createBootstrapMachine<BootstrapConfig, RuntimeSession>()

interface RemoteSyncAppProps {
  readonly appConfig: BiligRuntimeConfig
  readonly rawConfig: BiligRuntimeConfig
  readonly session: RuntimeSession
}

function RemoteSyncApp({ appConfig, rawConfig, session }: RemoteSyncAppProps) {
  const { authToken, userId } = session
  const zeroQueryContext = useMemo(() => createZeroQueryContext({ userId }), [userId])

  return (
    <ZeroProvider
      cacheURL={resolveZeroCacheUrl(rawConfig.zeroCacheUrl)}
      auth={authToken}
      context={zeroQueryContext}
      userID={userId}
      schema={schema}
      mutators={mutators}
      maxHeaderLength={ZERO_CONNECT_MAX_HEADER_LENGTH}
    >
      <App config={appConfig} />
    </ZeroProvider>
  )
}

function BootstrapShell() {
  return (
    <div aria-hidden="true" className="min-h-screen bg-[var(--wb-app-bg)] font-sans text-transparent" data-testid="bootstrap-shell">
      <div className="border-b border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2">
        <div className="flex min-h-[40px] items-center gap-3">
          <div className="h-8 w-28 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-muted)]" />
          <div className="h-8 w-[5.5rem] rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-muted)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
          <div className="h-5 w-px bg-[var(--wb-border)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
          <div className="h-7 w-7 rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)]" />
        </div>
      </div>
      <div className="border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-2">
        <div className="flex items-center gap-2">
          <div className="h-8 w-24 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
          <div className="h-8 flex-1 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
        </div>
      </div>
      <div className="flex-1 px-0">
        <div className="h-full overflow-hidden bg-[var(--wb-surface)]">
          <div className="grid grid-cols-[46px_repeat(10,minmax(0,1fr))] border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)]">
            <div className="h-6 border-r border-[var(--wb-border)] bg-[var(--wb-surface-subtle)]" />
            {Array.from({ length: 10 }, (_, index) => (
              <div
                className="h-6 border-r border-[var(--wb-grid-border)] bg-[var(--wb-surface-subtle)] last:border-r-0"
                key={`bootstrap-col-${index}`}
              />
            ))}
          </div>
          {Array.from({ length: 18 }, (_row, index) => (
            <div
              className="grid grid-cols-[46px_repeat(10,minmax(0,1fr))] border-b border-[var(--wb-grid-border)] last:border-b-0"
              key={`bootstrap-row-${index}`}
            >
              <div className="h-[22px] border-r border-[var(--wb-border)] bg-[var(--wb-surface-subtle)]" />
              {Array.from({ length: 10 }, (_cell, cellIndex) => (
                <div
                  className="h-[22px] border-r border-[var(--wb-grid-border)] bg-[var(--wb-surface)] last:border-r-0"
                  key={`bootstrap-cell-${index}-${cellIndex}`}
                />
              ))}
            </div>
          ))}
        </div>
      </div>
      <div className="flex min-h-11 items-center justify-between gap-3 border-t border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2.5 py-1.5">
        <div className="flex items-center gap-1.5">
          <div className="h-8 w-20 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
          <div className="h-8 w-20 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
          <div className="h-8 w-8 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
        </div>
        <div className="h-7 w-28 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)]" />
      </div>
    </div>
  )
}

function BootstrapRoot() {
  const actorRef = useActorRef(bootstrapMachine, {
    input: {
      autoRetryDelayMs: 750,
      failedRetryDelayMs: remoteSyncEnabled ? 5_000 : 0,
      maxAutoRetryAttempts: remoteSyncEnabled ? 3 : 0,
      loadConfig: async () => {
        if (!remoteSyncEnabled) {
          const rawConfig = createLocalOnlyRuntimeConfig()
          return {
            rawConfig,
            runtimeConfig: resolveRuntimeConfig(rawConfig),
          } satisfies BootstrapConfig
        }
        const rawConfig = await loadRuntimeConfig()
        return {
          rawConfig,
          runtimeConfig: resolveRuntimeConfig(rawConfig),
        } satisfies BootstrapConfig
      },
      loadSession: async () => (remoteSyncEnabled ? await loadRuntimeSession() : createLocalOnlyRuntimeSession()),
    },
  })
  const snapshot = useSelector(actorRef, (value) => value)

  if (snapshot.matches('failed')) {
    return (
      <div className="error-banner" data-testid="worker-error">
        {snapshot.context.error ?? 'Failed to bootstrap the web app'}
        <button onClick={() => actorRef.send({ type: 'retry' })} type="button">
          Retry
        </button>
      </div>
    )
  }

  if (!snapshot.matches('ready')) {
    return <BootstrapShell />
  }

  const config = snapshot.context.config
  const session = snapshot.context.session
  if (!config) {
    return (
      <div className="error-banner" data-testid="worker-error">
        Missing runtime config after bootstrap
      </div>
    )
  }

  if (!session) {
    return (
      <div className="error-banner" data-testid="worker-error">
        Missing runtime session after bootstrap
      </div>
    )
  }

  const appConfig = normalizeRuntimeConfigUserId(config.rawConfig, session)

  if (!remoteSyncEnabled) {
    return <App config={appConfig} connectionState={LOCAL_ONLY_CONNECTION_STATE} />
  }

  return <RemoteSyncApp appConfig={appConfig} rawConfig={config.rawConfig} session={session} />
}

root.render(
  <React.StrictMode>
    {entryRoute === 'isolated-workbook-pane-renderer' ? (
      <Suspense fallback={<BootstrapShell />}>
        <IsolatedWorkbookPaneRendererRoute />
      </Suspense>
    ) : (
      <BootstrapRoot />
    )}
  </React.StrictMode>,
)
