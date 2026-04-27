#!/usr/bin/env bun

import { existsSync, readFileSync, readlinkSync } from 'node:fs'
import net from 'node:net'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import {
  canUsePort as canUsePortWithListeners,
  resolvePreferredPort,
  resolvePreferredZeroPort,
  resolveRequestedOrAvailablePort,
} from './dev-web-local-ports.js'
import { ensureWasmKernelArtifact } from './ensure-wasm-kernel.js'

const composeFiles = ['compose.yaml', 'compose.dev-local.yaml'] as const
const composeProject = process.env['BILIG_DEV_COMPOSE_PROJECT'] ?? 'bilig-dev-local'
const postgresService = 'postgres'
const zeroCacheService = 'zero-cache-local'
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const webAppDir = resolve(repoRoot, 'apps/web')
const preferredAppPort = Number.parseInt(process.env['PORT'] ?? process.env['BILIG_SYNC_SERVER_PORT'] ?? '4321', 10)
const preferredPostgresPort = resolvePreferredPort(process.env['BILIG_DEV_POSTGRES_PORT'], 55432)
const preferredWebPort = Number.parseInt(process.env['BILIG_WEB_DEV_PORT'] ?? '5173', 10)
const configuredZeroProxyUpstream = process.env['BILIG_ZERO_PROXY_UPSTREAM']
const disableCompose = process.env['BILIG_DEV_DISABLE_COMPOSE'] === '1'
const webServerMode = process.env['BILIG_DEV_WEB_SERVER_MODE'] === 'preview' ? 'preview' : 'dev'
const appServerMode = process.env['BILIG_DEV_APP_SERVER_MODE'] === 'run' ? 'run' : 'watch'
const skipPreviewBuild = process.env['BILIG_DEV_WEB_PREVIEW_BUILD'] === '0'
const preferredZeroPort = resolvePreferredZeroPort(process.env['BILIG_DEV_ZERO_PORT'], configuredZeroProxyUpstream, 4848)
const composePublishedHost = resolveComposePublishedHost()
const cleanupCompose = process.env['BILIG_DEV_CLEANUP_COMPOSE'] === 'true'
let resolvedAppPort = String(preferredAppPort)
let resolvedPostgresPort = String(preferredPostgresPort)
let resolvedZeroPort = String(preferredZeroPort)

interface DevChildProcess {
  readonly exited: Promise<number | null>
  kill(signal?: 'SIGINT' | 'SIGTERM' | number): void
}

function childStdinMode(): 'ignore' | 'inherit' {
  return process.stdin.isTTY ? 'inherit' : 'ignore'
}

function commandExists(command: string): boolean {
  return Bun.which(command) !== null
}

function composeEnv(): NodeJS.ProcessEnv {
  return {
    ...process.env,
    BILIG_E2E_POSTGRES_PORT: resolvedPostgresPort,
    BILIG_DEV_POSTGRES_PORT: resolvedPostgresPort,
    BILIG_DEV_APP_PORT: resolvedAppPort,
    BILIG_DEV_ZERO_PORT: resolvedZeroPort,
  }
}

function parseProcRouteGateway(hexValue: string): string | null {
  if (!/^[0-9a-fA-F]{8}$/.test(hexValue)) {
    return null
  }
  const octets = hexValue
    .match(/../g)
    ?.map((chunk) => Number.parseInt(chunk, 16))
    .toReversed()
  if (!octets || octets.length !== 4 || octets.some((octet) => !Number.isInteger(octet))) {
    return null
  }
  return octets.join('.')
}

function resolveComposePublishedHost(): string {
  const explicitHost = process.env['BILIG_DOCKER_PUBLISHED_HOST']?.trim()
  if (explicitHost) {
    return explicitHost
  }

  if (process.platform !== 'linux' || !existsSync('/.dockerenv')) {
    return '127.0.0.1'
  }

  try {
    const routes = readFileSync('/proc/net/route', 'utf8').trim().split('\n').slice(1)
    for (const route of routes) {
      const fields = route.trim().split(/\s+/)
      const destination = fields[1] ?? ''
      const gateway = fields[2] ?? ''
      if (destination !== '00000000') {
        continue
      }
      return parseProcRouteGateway(gateway) ?? '127.0.0.1'
    }
  } catch {}

  return '127.0.0.1'
}

function containerRuntimeReady(): boolean {
  const runtimeCandidates = [{ command: ['docker', 'ps'] }, { command: ['podman', 'ps'] }]

  for (const candidate of runtimeCandidates) {
    if (!commandExists(candidate.command[0])) {
      continue
    }
    const result = Bun.spawnSync(candidate.command, {
      stdin: 'ignore',
      stdout: 'ignore',
      stderr: 'ignore',
      env: composeEnv(),
    })
    if (result.exitCode === 0) {
      return true
    }
  }

  return false
}

function resolveComposeInvocation(): {
  readonly command: readonly string[]
  readonly label: string
} | null {
  if (commandExists('docker')) {
    const dockerCompose = Bun.spawnSync(['docker', 'compose', 'version'], {
      stdin: 'ignore',
      stdout: 'ignore',
      stderr: 'ignore',
      env: composeEnv(),
    })
    if (dockerCompose.exitCode === 0) {
      return { command: ['docker', 'compose'], label: 'docker compose' }
    }
  }

  if (commandExists('podman')) {
    const podmanCompose = Bun.spawnSync(['podman', 'compose', 'version'], {
      stdin: 'ignore',
      stdout: 'ignore',
      stderr: 'ignore',
      env: composeEnv(),
    })
    if (podmanCompose.exitCode === 0) {
      return { command: ['podman', 'compose'], label: 'podman compose' }
    }
  }

  if (commandExists('docker-compose')) {
    const dockerComposeStandalone = Bun.spawnSync(['docker-compose', 'version'], {
      stdin: 'ignore',
      stdout: 'ignore',
      stderr: 'ignore',
      env: composeEnv(),
    })
    if (dockerComposeStandalone.exitCode === 0) {
      return { command: ['docker-compose'], label: 'docker-compose' }
    }
  }

  if (commandExists('podman-compose')) {
    const podmanComposeStandalone = Bun.spawnSync(['podman-compose', 'version'], {
      stdin: 'ignore',
      stdout: 'ignore',
      stderr: 'ignore',
      env: composeEnv(),
    })
    if (podmanComposeStandalone.exitCode === 0) {
      return { command: ['podman-compose'], label: 'podman-compose' }
    }
  }

  return null
}

function composeArgs(args: readonly string[]): readonly string[] {
  const invocation = resolveComposeInvocation()
  if (!invocation) {
    throw new Error('A supported compose command is required for pnpm dev:web-local.')
  }
  return [...invocation.command, ...composeFiles.flatMap((file) => ['-f', file]), '-p', composeProject, ...args]
}

function composeSupportsWait(): boolean {
  const invocation = resolveComposeInvocation()
  if (!invocation) {
    return false
  }
  return invocation.command[0] === 'docker' && invocation.command[1] === 'compose'
}

function runComposeSync(args: readonly string[], options?: { readonly allowFailure?: boolean }): void {
  const invocation = resolveComposeInvocation()
  if (!invocation) {
    throw new Error('A supported compose command is required for pnpm dev:web-local.')
  }

  const result = Bun.spawnSync(composeArgs(args), {
    stdin: 'inherit',
    stdout: 'inherit',
    stderr: 'inherit',
    env: composeEnv(),
  })

  if (!options?.allowFailure && result.exitCode !== 0) {
    throw new Error(`${invocation.label} ${args.join(' ')} failed with exit code ${result.exitCode ?? 1}`)
  }
}

async function waitForHttp(url: string, timeoutMs = 120_000): Promise<void> {
  const deadline = Date.now() + timeoutMs
  const poll = async (lastError: string): Promise<void> => {
    if (Date.now() > deadline) {
      throw new Error(`Timed out waiting for ${url} (${lastError})`)
    }
    try {
      const response = await fetch(url)
      if (response.ok) {
        return
      }
      await Bun.sleep(250)
      return poll(`HTTP ${response.status}`)
    } catch (error) {
      await Bun.sleep(250)
      return poll(error instanceof Error ? error.message : String(error))
    }
  }
  return poll('not started')
}

async function waitForTcp(host: string, port: number, timeoutMs = 120_000): Promise<void> {
  const deadline = Date.now() + timeoutMs
  const poll = async (lastError: string): Promise<void> => {
    if (Date.now() > deadline) {
      throw new Error(`Timed out waiting for tcp://${host}:${String(port)} (${lastError})`)
    }

    const connected = await new Promise<boolean>((resolveTcp) => {
      const socket = net.createConnection({ host, port })
      socket.once('connect', () => {
        socket.end()
        resolveTcp(true)
      })
      socket.once('error', () => {
        socket.destroy()
        resolveTcp(false)
      })
    })

    if (connected) {
      return
    }
    await Bun.sleep(250)
    return poll('connection refused')
  }

  return poll('not started')
}

async function ensureComposeRuntime(): Promise<void> {
  if (resolveComposeInvocation()) {
    if (containerRuntimeReady()) {
      return
    }

    if (process.platform === 'darwin' && Bun.which('open') && commandExists('docker')) {
      Bun.spawnSync(['open', '-a', 'Docker'], {
        stdin: 'ignore',
        stdout: 'ignore',
        stderr: 'ignore',
      })
    }

    const deadline = Date.now() + 120_000
    const waitForContainerRuntime = async (): Promise<void> => {
      if (containerRuntimeReady()) {
        return
      }
      if (Date.now() > deadline) {
        throw new Error('A running container runtime is required for pnpm dev:web-local.')
      }
      await Bun.sleep(2_000)
      return waitForContainerRuntime()
    }

    await waitForContainerRuntime()
    return
  }
  throw new Error('A supported compose command is required for pnpm dev:web-local.')
}

function listListeningPids(port: number): string[] {
  if (commandExists('lsof')) {
    const result = Bun.spawnSync(['lsof', '-tiTCP:' + String(port), '-sTCP:LISTEN'], {
      stdin: 'ignore',
      stdout: 'pipe',
      stderr: 'ignore',
    })
    if (result.exitCode !== 0) {
      return []
    }
    return new TextDecoder()
      .decode(result.stdout)
      .split('\n')
      .map((value) => value.trim())
      .filter((value) => value.length > 0)
  }

  if (!commandExists('ss')) {
    return []
  }

  const result = Bun.spawnSync(['ss', '-ltnp', `sport = :${String(port)}`], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'ignore',
  })
  if (result.exitCode !== 0) {
    return []
  }

  return [
    ...new Set(
      Array.from(new TextDecoder().decode(result.stdout).matchAll(/pid=(\d+)/g), (match) => match[1]?.trim()).filter(
        (value): value is string => Boolean(value && value.length > 0),
      ),
    ),
  ]
}

function commandForPid(pid: string): string {
  const result = Bun.spawnSync(['ps', '-p', pid, '-o', 'command='], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'ignore',
  })
  if (result.exitCode !== 0) {
    return ''
  }
  return new TextDecoder().decode(result.stdout).trim()
}

function cwdForPid(pid: string): string {
  if (commandExists('lsof')) {
    const result = Bun.spawnSync(['lsof', '-a', '-p', pid, '-d', 'cwd', '-Fn'], {
      stdin: 'ignore',
      stdout: 'pipe',
      stderr: 'ignore',
    })
    if (result.exitCode !== 0) {
      return ''
    }
    return (
      new TextDecoder()
        .decode(result.stdout)
        .split('\n')
        .find((line) => line.startsWith('n'))
        ?.slice(1)
        .trim() ?? ''
    )
  }

  if (process.platform !== 'linux') {
    return ''
  }

  try {
    return readlinkSync(`/proc/${pid}/cwd`)
  } catch {
    return ''
  }
}

function isRepoOwnedListener(pid: string): boolean {
  const command = commandForPid(pid)
  if (command.includes(repoRoot)) {
    return true
  }
  const cwd = cwdForPid(pid)
  return cwd === repoRoot || cwd.startsWith(`${repoRoot}/`)
}

async function reapStaleRepoListeners(ports: readonly number[]): Promise<void> {
  const repoOwnedPids = [...new Set(ports.flatMap((port) => listListeningPids(port)).filter(isRepoOwnedListener))]
  for (const pid of repoOwnedPids) {
    try {
      process.kill(Number.parseInt(pid, 10), 'SIGTERM')
    } catch {}
  }

  const deadline = Date.now() + 5_000
  const waitUntilCleared = async (): Promise<void> => {
    if (Date.now() > deadline) {
      return
    }
    const occupied = ports.some((port) => listListeningPids(port).some(isRepoOwnedListener))
    if (!occupied) {
      return
    }
    await Bun.sleep(100)
    return waitUntilCleared()
  }
  await waitUntilCleared()
}

async function resolveWebPort(preferredPort: number): Promise<number> {
  return resolveRequestedOrAvailablePort({
    preferredPort,
    explicitPort: process.env['BILIG_WEB_DEV_PORT'],
    label: explicitPortLabel('Web port', preferredPort),
    canUseRequestedPort: isPortAvailable,
  })
}

async function resolveAppPort(preferredPort: number): Promise<number> {
  return resolveRequestedOrAvailablePort({
    preferredPort,
    explicitPort: process.env['PORT'] ?? process.env['BILIG_SYNC_SERVER_PORT'],
    label: explicitPortLabel('App port', preferredPort),
    canUseRequestedPort: isPortAvailable,
  })
}

async function resolveComposePort(preferredPort: number, explicitPort: string | undefined, name: string): Promise<number> {
  return resolveRequestedOrAvailablePort({
    preferredPort,
    explicitPort,
    label: explicitPortLabel(name, preferredPort),
    canUseRequestedPort: isPortAvailable,
  })
}

function explicitPortLabel(name: string, preferredPort: number): string {
  return `${name} ${preferredPort}`
}

async function isPortAvailable(port: number): Promise<boolean> {
  return canUsePortWithListeners({
    port,
    listListeningPids,
    bindProbe: (candidatePort) =>
      new Promise((resolvePortAvailability) => {
        const server = net.createServer((socket) => {
          socket.destroy()
        })

        server.once('error', () => resolvePortAvailability(false))
        server.once('listening', () => {
          server.close(() => resolvePortAvailability(true))
        })

        server.listen(candidatePort, '127.0.0.1')
      }),
  })
}

function spawnAppDev(
  appPort: string,
  publicServerUrl: string,
  webAppBaseUrl: string,
  options?: {
    readonly postgresUrl?: string
    readonly zeroProxyUpstream?: string
  },
): DevChildProcess {
  const env: NodeJS.ProcessEnv = {
    ...process.env,
    HOST: process.env['HOST'] ?? '0.0.0.0',
    PORT: appPort,
    BILIG_PUBLIC_SERVER_URL: publicServerUrl,
    BILIG_WEB_APP_BASE_URL: webAppBaseUrl,
    BILIG_CORS_ORIGIN: webAppBaseUrl,
    BILIG_RUN_DATA_MIGRATIONS_ON_BOOT: process.env['BILIG_RUN_DATA_MIGRATIONS_ON_BOOT'] ?? 'true',
  }
  if (options?.postgresUrl) {
    env['DATABASE_URL'] = options.postgresUrl
  }
  if (options?.zeroProxyUpstream) {
    env['BILIG_ZERO_PROXY_UPSTREAM'] = options.zeroProxyUpstream
    env['BILIG_ZERO_CACHE_URL'] = '/zero'
  }
  const command =
    appServerMode === 'run'
      ? ['pnpm', '--filter', '@bilig/app', 'exec', 'tsx', 'src/index.ts']
      : ['pnpm', '--filter', '@bilig/app', 'run', 'dev']
  return Bun.spawn(command, {
    stdin: childStdinMode(),
    stdout: 'inherit',
    stderr: 'inherit',
    env,
  })
}

function spawnWebDev(webPort: number, publicServerUrl: string): DevChildProcess {
  return Bun.spawn(['node', '../../node_modules/vite/bin/vite.js', '--host', '0.0.0.0', '--port', String(webPort), '--strictPort'], {
    cwd: webAppDir,
    stdin: childStdinMode(),
    stdout: 'inherit',
    stderr: 'inherit',
    env: {
      ...process.env,
      BILIG_SYNC_SERVER_PORT: new URL(publicServerUrl).port,
      BILIG_SYNC_SERVER_TARGET: publicServerUrl,
      VITE_BILIG_REMOTE_SYNC: process.env['BILIG_E2E_REMOTE_SYNC'] ?? '1',
    },
  })
}

function buildWebPreview(publicServerUrl: string): void {
  ensureWasmKernelArtifact()
  const env: NodeJS.ProcessEnv = {
    ...process.env,
    BILIG_SYNC_SERVER_PORT: new URL(publicServerUrl).port,
    BILIG_SYNC_SERVER_TARGET: publicServerUrl,
    VITE_BILIG_REMOTE_SYNC: process.env['BILIG_E2E_REMOTE_SYNC'] ?? '1',
  }
  const result = Bun.spawnSync(['pnpm', '--filter', '@bilig/web', 'build'], {
    cwd: repoRoot,
    stdin: childStdinMode(),
    stdout: 'inherit',
    stderr: 'inherit',
    env,
  })
  if (result.exitCode !== 0) {
    throw new Error(`@bilig/web build failed with exit code ${result.exitCode ?? 1}`)
  }
}

function spawnWebPreview(webPort: number, publicServerUrl: string): DevChildProcess {
  return Bun.spawn(
    ['node', '../../node_modules/vite/bin/vite.js', 'preview', '--host', '0.0.0.0', '--port', String(webPort), '--strictPort'],
    {
      cwd: webAppDir,
      stdin: childStdinMode(),
      stdout: 'inherit',
      stderr: 'inherit',
      env: {
        ...process.env,
        BILIG_SYNC_SERVER_PORT: new URL(publicServerUrl).port,
        BILIG_SYNC_SERVER_TARGET: publicServerUrl,
        VITE_BILIG_REMOTE_SYNC: process.env['BILIG_E2E_REMOTE_SYNC'] ?? '1',
      },
    },
  )
}

function killIfRunning(process: DevChildProcess | null | undefined, signal: NodeJS.Signals): void {
  if (!process) {
    return
  }
  try {
    process.kill(signal)
  } catch {}
}

function cleanupComposeStack(): void {
  if (disableCompose || !cleanupCompose) {
    return
  }
  runComposeSync(['down', '-v', '--remove-orphans'], { allowFailure: true })
}

console.log(disableCompose ? 'Starting local dev stack without compose dependencies...' : 'Starting local compose dependencies...')
if (!disableCompose) {
  await ensureComposeRuntime()
  cleanupComposeStack()
}
await reapStaleRepoListeners(Array.from({ length: 10 }, (_, index) => preferredWebPort + index).concat(preferredAppPort))
const appPort = await resolveAppPort(preferredAppPort)
resolvedAppPort = String(appPort)
if (!disableCompose) {
  resolvedPostgresPort = String(await resolveComposePort(preferredPostgresPort, process.env['BILIG_DEV_POSTGRES_PORT'], 'Postgres port'))
  const explicitZeroPort =
    process.env['BILIG_DEV_ZERO_PORT'] ?? (configuredZeroProxyUpstream ? new URL(configuredZeroProxyUpstream).port : undefined)
  resolvedZeroPort = String(await resolveComposePort(preferredZeroPort, explicitZeroPort, 'Zero port'))
}
const zeroProxyUpstream = configuredZeroProxyUpstream ?? `http://${composePublishedHost}:${resolvedZeroPort}`
const zeroHealthUrl = `${zeroProxyUpstream}/keepalive`
const postgresUrl = disableCompose
  ? undefined
  : (process.env['DATABASE_URL'] ?? `postgresql://bilig:bilig@${composePublishedHost}:${resolvedPostgresPort}/bilig`)
const publicServerUrl = process.env['BILIG_PUBLIC_SERVER_URL'] ?? `http://127.0.0.1:${appPort}`
const appHealthUrl = `${publicServerUrl}/healthz`
if (!disableCompose && composeSupportsWait()) {
  runComposeSync(['up', '-d', postgresService])
} else if (!disableCompose) {
  runComposeSync(['up', '-d', postgresService])
}
if (!disableCompose) {
  await waitForTcp(composePublishedHost, Number(resolvedPostgresPort))
}
const webPort = await resolveWebPort(preferredWebPort)
const webAppBaseUrl = process.env['BILIG_WEB_APP_BASE_URL'] ?? `http://localhost:${webPort}`

console.log(`Starting local app dev server (app=${publicServerUrl})...`)
const appChild = spawnAppDev(String(appPort), publicServerUrl, webAppBaseUrl, {
  ...(postgresUrl ? { postgresUrl } : {}),
  ...(disableCompose ? {} : { zeroProxyUpstream }),
})
let webChild: DevChildProcess | null = null

let shuttingDown = false

function forwardSignal(signal: NodeJS.Signals): void {
  if (shuttingDown) {
    return
  }
  shuttingDown = true
  killIfRunning(appChild, signal)
  killIfRunning(webChild, signal)
  cleanupComposeStack()
}

process.on('SIGINT', () => forwardSignal('SIGINT'))
process.on('SIGTERM', () => forwardSignal('SIGTERM'))

try {
  await waitForHttp(appHealthUrl)
  if (!disableCompose) {
    runComposeSync(['up', '-d', zeroCacheService])
    await waitForHttp(zeroHealthUrl)
  }
  if (webServerMode === 'preview') {
    if (skipPreviewBuild) {
      console.log(`Reusing existing preview web bundle for browser stack (web=${webAppBaseUrl})...`)
    } else {
      console.log(`Building preview web bundle for browser stack (web=${webAppBaseUrl})...`)
      buildWebPreview(publicServerUrl)
    }
    console.log(`Starting local web preview server (web=${webAppBaseUrl})...`)
    webChild = spawnWebPreview(webPort, publicServerUrl)
  } else {
    console.log(`Starting local web dev server (web=${webAppBaseUrl})...`)
    webChild = spawnWebDev(webPort, publicServerUrl)
  }
  console.log('App is healthy.')
  await waitForHttp(webAppBaseUrl)
  console.log(
    disableCompose
      ? `Local dev stack ready: web=${webAppBaseUrl} app=${publicServerUrl} sync=local-only`
      : `Local dev stack ready: web=${webAppBaseUrl} app=${publicServerUrl} zero=${zeroProxyUpstream}`,
  )
} catch (error) {
  forwardSignal('SIGTERM')
  throw error
}

const exitCode = await Promise.race([
  appChild.exited.then((code) => ({ code, source: '@bilig/app' })),
  (webChild?.exited ?? new Promise<never>(() => undefined)).then((code) => ({
    code,
    source: '@bilig/web',
  })),
])

forwardSignal('SIGTERM')

if (exitCode.code !== 0) {
  console.error(`${exitCode.source} exited with code ${exitCode.code ?? 1}`)
}

process.exit(exitCode.code ?? 0)
