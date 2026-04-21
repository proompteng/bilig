#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import net from 'node:net'
import { resolveBrowserLocalWebMode } from './browser-stack-config.js'

const textDecoder = new TextDecoder()
const playwrightArgs = process.argv.slice(2)
const CLIPBOARD_GLOBAL_GREP = '@clipboard-global'
const BROWSER_PERF_GREP = '@browser-perf'
const BROWSER_SERIAL_GREP = '@browser-serial'
const requestedBrowserStack = process.env['BILIG_BROWSER_STACK'] ?? 'auto'
const normalizedBrowserStack =
  requestedBrowserStack === 'compose' || requestedBrowserStack === 'local' || requestedBrowserStack === 'auto'
    ? requestedBrowserStack
    : 'local'
const isCi = process.env['CI'] === '1' || process.env['CI'] === 'true'
if (requestedBrowserStack !== normalizedBrowserStack) {
  console.warn(`Unknown BILIG_BROWSER_STACK "${requestedBrowserStack}", defaulting to "local" stack.`)
}
type ComposeInvocation = {
  label: string
  command: string[]
  version: string
}
interface BrowserStackProcess {
  readonly exited: Promise<number | null>
  kill(signal?: 'SIGINT' | 'SIGTERM' | 'SIGKILL' | number): void
}

let composeInvocation: ComposeInvocation | null = null
let composeInvocationProbed = false
let composeInvocationLogged = false

function commandExists(command: string): boolean {
  return Bun.which(command) !== null
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

function resolvePublishedServiceHosts(): string[] {
  const explicitHost = process.env['BILIG_E2E_HOST']?.trim()
  if (explicitHost) {
    return [explicitHost]
  }

  const hosts = ['127.0.0.1', 'host.docker.internal', 'host.containers.internal']

  if (process.platform !== 'linux' || !existsSync('/proc/net/route')) {
    return hosts
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
      const parsedGateway = parseProcRouteGateway(gateway)
      if (parsedGateway) {
        hosts.push(parsedGateway)
      }
      break
    }
  } catch {}

  return Array.from(new Set(hosts))
}

function probeComposeInvocation(): ComposeInvocation | null {
  const composeCandidates = [
    {
      label: 'docker compose',
      command: ['docker', 'compose'],
      runtimeProbe: ['docker', 'ps'],
    },
    {
      label: 'podman compose',
      command: ['podman', 'compose'],
      runtimeProbe: ['podman', 'ps'],
    },
    {
      label: 'docker-compose',
      command: ['docker-compose'],
      runtimeProbe: ['docker', 'ps'],
    },
    {
      label: 'podman-compose',
      command: ['podman-compose'],
      runtimeProbe: ['podman', 'ps'],
    },
  ]

  for (const candidate of composeCandidates) {
    if (!commandExists(candidate.command[0])) {
      continue
    }

    const result = Bun.spawnSync([...candidate.command, 'version'], {
      stdin: 'ignore',
      stdout: 'pipe',
      stderr: 'pipe',
    })
    if (result.exitCode !== 0) {
      continue
    }

    const runtimeProbe = Bun.spawnSync(candidate.runtimeProbe, {
      stdin: 'ignore',
      stdout: 'pipe',
      stderr: 'pipe',
    })
    if (runtimeProbe.exitCode !== 0) {
      continue
    }

    const version = [textDecoder.decode(result.stdout), textDecoder.decode(result.stderr)].join('').trim()
    return {
      label: candidate.label,
      command: candidate.command,
      version,
    }
  }

  return null
}

function resolveComposeInvocation(): ComposeInvocation | null {
  if (!composeInvocationProbed) {
    composeInvocation = probeComposeInvocation()
    composeInvocationProbed = true
  }

  return composeInvocation
}

function requireComposeInvocation(required: boolean): ComposeInvocation | null {
  const invocation = resolveComposeInvocation()

  if (!invocation && required) {
    throw new Error(
      'container compose is required for BILIG_BROWSER_STACK=compose, but none of `docker compose`, `podman compose`, `docker-compose`, or `podman-compose` is available.',
    )
  }

  if (invocation && !composeInvocationLogged) {
    const version = invocation.version ? ` (${invocation.version})` : ''
    console.log(`compose is available via "${invocation.label}"${version}.`)
    composeInvocationLogged = true
  }

  return invocation
}

const compose = resolveComposeInvocation()
const composeLabel = compose ? compose.label : 'unavailable'
const browserStack = normalizedBrowserStack === 'compose' && compose ? 'compose' : 'local'

if (normalizedBrowserStack === 'compose' && compose) {
  console.log(`BILIG_BROWSER_STACK=compose requested; using compose command "${composeLabel}"`)
}

if (normalizedBrowserStack === 'compose' && !compose && isCi) {
  throw new Error('BILIG_BROWSER_STACK=compose is required in CI, but no supported compose command is available.')
}

if (normalizedBrowserStack === 'compose' && !compose && !isCi) {
  const fallbackCommand = '`docker compose`, `podman compose`, `docker-compose`, or `podman-compose`'
  console.warn(
    `compose unavailable in this environment, falling back to local Playwright server for browser tests (requested compose command: ${fallbackCommand})`,
  )
}
const composeFile = process.env['BILIG_E2E_COMPOSE_FILE'] ?? 'compose.yaml'
const composeProject = process.env['BILIG_E2E_COMPOSE_PROJECT'] ?? `bilig-e2e-${Date.now()}`
const composeStartupTimeoutMs = resolveTimeoutMs(process.env['BILIG_E2E_STARTUP_TIMEOUT_MS'], isCi ? 300_000 : 120_000)
const LOCAL_STACK_STARTUP_ATTEMPTS = 3
const LOCAL_STACK_STABILITY_GRACE_MS = 1_000
const LOCAL_PLAYWRIGHT_PHASE_ATTEMPTS = 2

const DEFAULT_PREVIEW_PORTS = [4179, 4180] as const

function resolveTimeoutMs(value: string | undefined, fallbackMs: number): number {
  if (!value) {
    return fallbackMs
  }

  const parsed = Number.parseInt(value, 10)
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return fallbackMs
  }

  return parsed
}

function parsePidList(output: string): number[] {
  if (!output) {
    return []
  }
  return output
    .split(/\s+/)
    .map((value) => Number.parseInt(value, 10))
    .filter((value) => Number.isInteger(value))
}

function parseSsPids(output: string): number[] {
  const matches = output.matchAll(/pid=(\d+)/g)
  const pids: number[] = []
  for (const match of matches) {
    const pid = Number.parseInt(match[1] ?? '', 10)
    if (Number.isInteger(pid)) {
      pids.push(pid)
    }
  }
  return pids
}

function getListeningPids(port: number): number[] {
  if (commandExists('lsof')) {
    const result = Bun.spawnSync(['lsof', '-tiTCP:' + String(port), '-sTCP:LISTEN'], {
      stdin: 'ignore',
      stdout: 'pipe',
      stderr: 'ignore',
    })
    if (result.exitCode !== 0) {
      return []
    }
    return parsePidList(textDecoder.decode(result.stdout).trim())
  }

  if (commandExists('ss')) {
    const result = Bun.spawnSync(['ss', '-ltnp', 'sport = :' + String(port)], {
      stdin: 'ignore',
      stdout: 'pipe',
      stderr: 'ignore',
    })
    if (result.exitCode !== 0) {
      return []
    }
    return parseSsPids(textDecoder.decode(result.stdout))
  }

  if (commandExists('netstat')) {
    const result = Bun.spawnSync(['netstat', '-ltnp'], {
      stdin: 'ignore',
      stdout: 'pipe',
      stderr: 'ignore',
    })
    if (result.exitCode !== 0) {
      return []
    }
    const lines = textDecoder
      .decode(result.stdout)
      .split('\n')
      .filter((line) => line.includes(':' + String(port)) && line.includes('LISTEN'))
    const pids: number[] = []
    for (const line of lines) {
      const fields = line.trim().split(/\s+/)
      const program = fields.at(-1) ?? ''
      const pid = Number.parseInt(program.split('/', 1)[0] ?? '', 10)
      if (Number.isInteger(pid)) {
        pids.push(pid)
      }
    }
    return pids
  }

  if (commandExists('fuser')) {
    const result = Bun.spawnSync(['fuser', '-n', 'tcp', String(port)], {
      stdin: 'ignore',
      stdout: 'pipe',
      stderr: 'ignore',
    })
    if (result.exitCode !== 0) {
      return []
    }
    return parsePidList(textDecoder.decode(result.stdout).trim())
  }

  return []
}

function sleep(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms)
}

async function reservePort(port: number): Promise<number | null> {
  return await new Promise<number | null>((resolve) => {
    const server = net.createServer()

    server.unref()
    server.once('error', () => {
      resolve(null)
    })
    server.listen(port, () => {
      const address = server.address()
      server.close(() => {
        if (address && typeof address === 'object') {
          resolve(address.port)
          return
        }
        resolve(null)
      })
    })
  })
}

async function resolvePort(envValue: string | undefined, preferredPort: number): Promise<string> {
  if (envValue) {
    return envValue
  }

  const preferred = await reservePort(preferredPort)
  if (preferred !== null) {
    return String(preferred)
  }

  const ephemeral = await reservePort(0)
  if (ephemeral !== null) {
    return String(ephemeral)
  }

  throw new Error(`Unable to allocate a free TCP port for browser tests near ${preferredPort}.`)
}

const e2eWebPort = await resolvePort(process.env['BILIG_E2E_WEB_PORT'], 4180)
const e2eSyncServerPort = await resolvePort(process.env['BILIG_E2E_SYNC_SERVER_PORT'], 54422)
const e2eZeroPort = await resolvePort(process.env['BILIG_E2E_ZERO_PORT'], 54849)
const e2ePostgresPort = await resolvePort(process.env['BILIG_E2E_POSTGRES_PORT'], 55433)
const e2eHostCandidates = resolvePublishedServiceHosts()
const configuredE2eBaseUrl = process.env['BILIG_E2E_BASE_URL']
const configuredE2eSyncServerUrl = process.env['BILIG_E2E_SYNC_SERVER_URL']
const configuredE2eZeroKeepaliveUrl = process.env['BILIG_E2E_ZERO_KEEPALIVE_URL']
let e2eHost = e2eHostCandidates[0] ?? '127.0.0.1'

function getE2eBaseUrl(): string {
  return configuredE2eBaseUrl ?? `http://${e2eHost}:${e2eWebPort}`
}

function getE2eSyncServerUrl(): string {
  return configuredE2eSyncServerUrl ?? `http://${e2eHost}:${e2eSyncServerPort}`
}

function getE2eZeroKeepaliveUrl(): string {
  return configuredE2eZeroKeepaliveUrl ?? `${getE2eBaseUrl()}/zero/keepalive`
}

function terminatePreviewServers(): void {
  const ports = [...DEFAULT_PREVIEW_PORTS, Number.parseInt(e2eWebPort, 10), Number.parseInt(e2eSyncServerPort, 10)].filter(
    (port) => Number.isInteger(port) && port > 0,
  )
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const pids = Array.from(new Set(ports.flatMap((port) => getListeningPids(port))))
    if (pids.length === 0) {
      return
    }

    for (const pid of pids) {
      try {
        process.kill(pid, 'SIGTERM')
      } catch {}
    }

    sleep(300)

    for (const pid of pids) {
      try {
        process.kill(pid, 0)
        process.kill(pid, 'SIGKILL')
      } catch {}
    }

    sleep(100)
  }
}

function runPlaywright(args: string[]): void {
  const result = Bun.spawnSync(['pnpm', 'exec', 'playwright', 'test', ...args], {
    stdin: 'inherit',
    stdout: 'inherit',
    stderr: 'inherit',
    env: {
      ...process.env,
      BILIG_BROWSER_STACK: browserStack,
      BILIG_DEV_DISABLE_COMPOSE: browserStack === 'local' ? '1' : (process.env['BILIG_DEV_DISABLE_COMPOSE'] ?? '0'),
      BILIG_E2E_REMOTE_SYNC: browserStack === 'local' ? '0' : (process.env['BILIG_E2E_REMOTE_SYNC'] ?? '1'),
      BILIG_E2E_WEB_PORT: e2eWebPort,
      BILIG_E2E_SYNC_SERVER_PORT: e2eSyncServerPort,
      BILIG_E2E_ZERO_PORT: e2eZeroPort,
      BILIG_E2E_POSTGRES_PORT: e2ePostgresPort,
      BILIG_E2E_HOST: e2eHost,
      BILIG_E2E_BASE_URL: getE2eBaseUrl(),
      BILIG_E2E_SYNC_SERVER_URL: getE2eSyncServerUrl(),
      BILIG_E2E_ZERO_KEEPALIVE_URL: getE2eZeroKeepaliveUrl(),
      ...(browserStack === 'local' ? { BILIG_E2E_MANAGED_STACK: '1' } : {}),
    },
  })
  if (result.exitCode !== 0) {
    throw new Error(`playwright test ${args.join(' ')} failed with exit code ${result.exitCode ?? 1}`)
  }
}

function configuredPlaywrightArgSets(): string[][] {
  if (playwrightArgs.length > 0) {
    return [playwrightArgs]
  }

  return [
    ['--grep-invert', `${CLIPBOARD_GLOBAL_GREP}|${BROWSER_PERF_GREP}|${BROWSER_SERIAL_GREP}`],
    ['--workers=1', '--grep', BROWSER_PERF_GREP],
    ['--workers=1', '--grep', BROWSER_SERIAL_GREP],
    ['--workers=1', '--grep', CLIPBOARD_GLOBAL_GREP],
  ]
}

function runConfiguredPlaywrightSuites(): void {
  for (const args of configuredPlaywrightArgSets()) {
    runPlaywright(args)
  }
}

async function pollHttp(url: string, deadline: number, lastError = 'unknown error'): Promise<void> {
  if (Date.now() >= deadline) {
    throw new Error(`Timed out waiting for ${url}: ${lastError}`)
  }
  try {
    const response = await fetch(url)
    if (response.ok) {
      return
    }
    lastError = `HTTP ${response.status}`
  } catch (error) {
    lastError = error instanceof Error ? error.message : String(error)
  }
  await Bun.sleep(250)
  await pollHttp(url, deadline, lastError)
}

async function waitForHttp(url: string, timeoutMs = 120_000): Promise<void> {
  await pollHttp(url, Date.now() + timeoutMs)
}

async function resolveReachableHttpHost(hosts: readonly string[], port: string, pathname: string, timeoutMs: number): Promise<string> {
  return pollReachableHttpHost(hosts, port, pathname, Date.now() + timeoutMs)
}

async function pollReachableHttpHost(
  hosts: readonly string[],
  port: string,
  pathname: string,
  deadline: number,
  lastError = 'unknown error',
): Promise<string> {
  if (Date.now() >= deadline) {
    throw new Error(`Timed out waiting for a reachable host on port ${port} (${hosts.join(', ')}): ${lastError}`)
  }

  const requestTimeoutMs = Math.max(250, Math.min(2_000, deadline - Date.now()))

  try {
    return await Promise.any(
      hosts.map(async (host) => {
        const url = `http://${host}:${port}${pathname}`
        let response: Response
        try {
          response = await fetch(url, { signal: AbortSignal.timeout(requestTimeoutMs) })
        } catch (error) {
          throw new Error(`${url}: ${error instanceof Error ? error.message : String(error)}`, {
            cause: error,
          })
        }
        if (!response.ok) {
          throw new Error(`${url}: HTTP ${response.status}`)
        }
        return host
      }),
    )
  } catch (error) {
    const nextLastError =
      error instanceof AggregateError && error.errors.length > 0
        ? error.errors.map((entry) => (entry instanceof Error ? entry.message : String(entry))).join('; ')
        : error instanceof Error
          ? error.message
          : String(error)
    await Bun.sleep(250)
    return pollReachableHttpHost(hosts, port, pathname, deadline, nextLastError)
  }
}

async function pollTcp(host: string, port: number, deadline: number, lastError = 'unknown error'): Promise<void> {
  if (Date.now() >= deadline) {
    throw new Error(`Timed out waiting for ${host}:${port}: ${lastError}`)
  }

  try {
    await new Promise<void>((resolve, reject) => {
      const socket = net.createConnection({ host, port })
      const cleanup = () => {
        socket.removeAllListeners()
        socket.destroy()
      }

      socket.once('connect', () => {
        cleanup()
        resolve()
      })
      socket.once('error', (error) => {
        cleanup()
        reject(error)
      })
    })
    return
  } catch (error) {
    lastError = error instanceof Error ? error.message : String(error)
  }

  await Bun.sleep(250)
  await pollTcp(host, port, deadline, lastError)
}

async function waitForTcp(host: string, port: number, timeoutMs = 120_000): Promise<void> {
  await pollTcp(host, port, Date.now() + timeoutMs)
}

function runDockerCompose(args: string[], env = process.env): void {
  const invocation = requireComposeInvocation(true)
  if (!invocation) {
    throw new Error('compose command is unavailable; cannot run compose stack.')
  }

  const result = Bun.spawnSync([...invocation.command, '-f', composeFile, '-p', composeProject, ...args], {
    stdin: 'inherit',
    stdout: 'inherit',
    stderr: 'inherit',
    env: {
      ...env,
      BILIG_E2E_WEB_PORT: e2eWebPort,
      BILIG_E2E_SYNC_SERVER_PORT: e2eSyncServerPort,
      BILIG_E2E_ZERO_PORT: e2eZeroPort,
      BILIG_E2E_POSTGRES_PORT: e2ePostgresPort,
    },
  })
  if (result.exitCode !== 0) {
    throw new Error(`${invocation.label} ${args.join(' ')} failed with exit code ${result.exitCode ?? 1}`)
  }
}

function collectComposeLogs(): string {
  const invocation = requireComposeInvocation(false)
  if (!invocation) {
    return 'compose command is unavailable; compose logs were not collected.'
  }

  const result = Bun.spawnSync([...invocation.command, '-f', composeFile, '-p', composeProject, 'logs', '--no-color'], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  return [textDecoder.decode(result.stdout), textDecoder.decode(result.stderr)].join('').trim()
}

async function runComposePlaywright(): Promise<void> {
  requireComposeInvocation(true)

  terminatePreviewServers()
  runDockerCompose(['up', '-d', '--build', 'postgres', 'bilig-app', 'zero-cache'])
  try {
    console.log(
      `compose browser stack starting with hostCandidates=${e2eHostCandidates.join(',')}, web=${e2eWebPort}, sync=${e2eSyncServerPort}, zero=${e2eZeroPort}, postgres=${e2ePostgresPort}, startupTimeoutMs=${String(composeStartupTimeoutMs)}`,
    )
    if (!process.env['BILIG_E2E_HOST'] && !configuredE2eBaseUrl && !configuredE2eSyncServerUrl && !configuredE2eZeroKeepaliveUrl) {
      e2eHost = await resolveReachableHttpHost(e2eHostCandidates, e2eWebPort, '/healthz', composeStartupTimeoutMs)
      console.log(`compose browser stack resolved host=${e2eHost}`)
    }
    await waitForHttp(`${getE2eBaseUrl()}/healthz`, composeStartupTimeoutMs)
    await waitForHttp(`${getE2eSyncServerUrl()}/healthz`, composeStartupTimeoutMs)
    await waitForTcp(e2eHost, Number.parseInt(e2eZeroPort, 10), composeStartupTimeoutMs)
    await waitForHttp(getE2eZeroKeepaliveUrl(), composeStartupTimeoutMs)
    runConfiguredPlaywrightSuites()
  } catch (error) {
    const logs = collectComposeLogs()
    throw new Error(`${error instanceof Error ? error.message : String(error)}\n${logs}`, {
      cause: error,
    })
  } finally {
    runDockerCompose(['down', '-v', '--remove-orphans'])
  }
}

async function stopLocalPlaywrightStack(child: BrowserStackProcess): Promise<void> {
  try {
    child.kill('SIGTERM')
  } catch {}

  const exited = await Promise.race([child.exited, Bun.sleep(5_000).then(() => null)])
  if (exited !== null) {
    return
  }

  try {
    child.kill('SIGKILL')
  } catch {}
  await child.exited.catch(() => undefined)
}

async function waitForLocalPlaywrightStack(child: BrowserStackProcess): Promise<void> {
  await Promise.race([
    (async () => {
      await waitForHttp(`${getE2eSyncServerUrl()}/runtime-config.json`, composeStartupTimeoutMs)
      await waitForHttp(getE2eBaseUrl(), composeStartupTimeoutMs)
    })(),
    child.exited.then((code) => {
      throw new Error(`local browser stack exited before ready with code ${code ?? 1}`)
    }),
  ])
}

async function readLocalPlaywrightStackExitCode(child: BrowserStackProcess): Promise<number | null | undefined> {
  return await Promise.race([child.exited, Bun.sleep(0).then(() => undefined)])
}

async function isLocalPlaywrightStackReady(child: BrowserStackProcess): Promise<boolean> {
  const exitCode = await readLocalPlaywrightStackExitCode(child)
  if (exitCode !== undefined) {
    return false
  }

  try {
    await waitForHttp(`${getE2eSyncServerUrl()}/runtime-config.json`, 2_000)
    await waitForHttp(getE2eBaseUrl(), 2_000)
    return true
  } catch {
    return false
  }
}

async function waitForLocalPlaywrightStackStable(child: BrowserStackProcess): Promise<void> {
  await waitForLocalPlaywrightStack(child)
  await Bun.sleep(LOCAL_STACK_STABILITY_GRACE_MS)
  if (!(await isLocalPlaywrightStackReady(child))) {
    throw new Error('local browser stack exited immediately after reporting ready')
  }
}

async function startLocalPlaywrightStack(attempt = 1): Promise<BrowserStackProcess> {
  terminatePreviewServers()
  const child = Bun.spawn(['bun', 'scripts/run-dev-web-local.ts'], {
    stdin: 'ignore',
    stdout: 'inherit',
    stderr: 'inherit',
    env: {
      ...process.env,
      BILIG_WEB_DEV_PORT: e2eWebPort,
      PORT: e2eSyncServerPort,
      BILIG_DEV_POSTGRES_PORT: e2ePostgresPort,
      BILIG_DEV_ZERO_PORT: e2eZeroPort,
      BILIG_DEV_WEB_SERVER_MODE: resolveBrowserLocalWebMode(process.env),
      BILIG_DEV_APP_SERVER_MODE: 'run',
      BILIG_DEV_COMPOSE_PROJECT: 'bilig-playwright-local',
      BILIG_DEV_CLEANUP_COMPOSE: 'true',
      BILIG_DEV_DISABLE_COMPOSE: '1',
      BILIG_E2E_REMOTE_SYNC: '0',
    },
  })
  try {
    await waitForLocalPlaywrightStackStable(child)
    return child
  } catch (error) {
    await stopAndReapLocalPlaywrightStack(child)
    if (attempt >= LOCAL_STACK_STARTUP_ATTEMPTS) {
      throw error
    }
    console.warn(
      `local browser stack was not stable after startup; restarting (${String(attempt + 1)}/${String(LOCAL_STACK_STARTUP_ATTEMPTS)})`,
    )
    return startLocalPlaywrightStack(attempt + 1)
  }
}

async function stopAndReapLocalPlaywrightStack(child: BrowserStackProcess | null): Promise<void> {
  if (child) {
    await stopLocalPlaywrightStack(child)
  }
  terminatePreviewServers()
  await Bun.sleep(1_000)
  terminatePreviewServers()
}

async function ensureLocalPlaywrightStack(child: BrowserStackProcess | null): Promise<BrowserStackProcess> {
  if (child && (await isLocalPlaywrightStackReady(child))) {
    return child
  }

  await stopAndReapLocalPlaywrightStack(child)
  return startLocalPlaywrightStack()
}

async function runLocalPlaywrightPhase(args: string[], child: BrowserStackProcess | null, attempt = 1): Promise<BrowserStackProcess> {
  const currentChild = await ensureLocalPlaywrightStack(child)
  try {
    runPlaywright(args)
    return currentChild
  } catch (error) {
    if ((await isLocalPlaywrightStackReady(currentChild)) || attempt >= LOCAL_PLAYWRIGHT_PHASE_ATTEMPTS) {
      throw error
    }
    console.warn(
      `local browser stack exited during Playwright phase "${args.join(' ')}"; restarting (${String(attempt + 1)}/${String(LOCAL_PLAYWRIGHT_PHASE_ATTEMPTS)})`,
    )
    await stopAndReapLocalPlaywrightStack(currentChild)
    return runLocalPlaywrightPhase(args, null, attempt + 1)
  }
}

async function runLocalPlaywright(): Promise<void> {
  let child: BrowserStackProcess | null = null
  try {
    for (const args of configuredPlaywrightArgSets()) {
      // oxlint-disable-next-line eslint(no-await-in-loop)
      child = await runLocalPlaywrightPhase(args, child)
    }
  } finally {
    await stopAndReapLocalPlaywrightStack(child)
  }
}

if (browserStack === 'compose') {
  await runComposePlaywright()
} else {
  await runLocalPlaywright()
}
