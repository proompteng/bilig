import { spawn } from 'node:child_process'

interface WatchableChildProcess {
  readonly pid?: number
  readonly kill: (signal?: 'SIGTERM' | 'SIGKILL') => boolean | void
}

export interface ChildRssWatchdogOptions {
  readonly maxRssBytes: number
  readonly intervalMs?: number
  readonly onSample?: (rssBytes: number) => void
  readonly onLimitExceeded: (rssBytes: number) => void
}

const defaultRssCheckIntervalMs = 1_000
export const defaultSelfRssCheckIntervalMs = 500

export function terminateChildProcess(
  child: WatchableChildProcess,
  signal: 'SIGTERM' | 'SIGKILL',
  options: { readonly processGroup?: boolean } = {},
): void {
  const pid = child.pid ?? 0
  if (options.processGroup === true && process.platform !== 'win32' && Number.isInteger(pid) && pid > 0) {
    try {
      process.kill(-pid, signal)
      return
    } catch {
      // Fall back to the direct child below. The close/error handlers own final classification.
    }
  }
  try {
    child.kill(signal)
  } catch {
    // The close/error handlers own the final failure classification.
  }
}

export function startChildRssWatchdog(child: WatchableChildProcess, options: ChildRssWatchdogOptions): () => void {
  const pid = child.pid
  if (!Number.isInteger(pid) || pid <= 0 || !Number.isFinite(options.maxRssBytes) || options.maxRssBytes <= 0) {
    return () => undefined
  }

  const intervalMs = Math.max(10, Math.trunc(options.intervalMs ?? defaultRssCheckIntervalMs))
  let stopped = false
  let timer: ReturnType<typeof setTimeout> | undefined
  let checking = false
  let peakRssBytes = 0

  const schedule = (): void => {
    if (stopped) {
      return
    }
    timer = setTimeout(() => {
      void check()
    }, intervalMs)
    timer.unref()
  }

  const check = async (): Promise<void> => {
    if (stopped || checking) {
      schedule()
      return
    }
    checking = true
    try {
      const rssBytes = await readProcessRssBytes(pid)
      if (rssBytes !== null) {
        peakRssBytes = Math.max(peakRssBytes, rssBytes)
        options.onSample?.(peakRssBytes)
      }
      if (rssBytes !== null && rssBytes > options.maxRssBytes) {
        options.onLimitExceeded(rssBytes)
        return
      }
    } finally {
      checking = false
    }
    schedule()
  }

  schedule()

  return () => {
    stopped = true
    if (timer) {
      clearTimeout(timer)
    }
  }
}

export function formatByteSize(bytes: number): string {
  const gib = bytes / 1024 / 1024 / 1024
  if (gib >= 1) {
    return `${gib.toFixed(2)} GiB`
  }
  const mib = bytes / 1024 / 1024
  return `${mib.toFixed(1)} MiB`
}

export function startSelfRssGuard(maxRssBytes: number, label: string, intervalMs = defaultSelfRssCheckIntervalMs): () => void {
  const normalizedMaxRssBytes = Math.max(1, Math.trunc(maxRssBytes))
  const timer = setInterval(() => {
    const rssBytes = process.memoryUsage().rss
    if (rssBytes <= normalizedMaxRssBytes) {
      return
    }
    console.error(`${label} exceeded RSS limit: ${formatByteSize(rssBytes)} > ${formatByteSize(normalizedMaxRssBytes)}`)
    process.exit(70)
  }, intervalMs)
  timer.unref()
  return () => clearInterval(timer)
}

function readProcessRssBytes(pid: number): Promise<number | null> {
  return new Promise((resolve) => {
    const child = spawn('/bin/ps', ['-o', 'rss=', '-p', String(pid)], {
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let settled = false
    const finish = (value: number | null): void => {
      if (settled) {
        return
      }
      settled = true
      clearTimeout(timer)
      // oxlint-disable-next-line eslint-plugin-promise(no-multiple-resolved) -- `settled` gates all event races before resolving.
      resolve(value)
    }
    const timer = setTimeout(() => {
      terminateChildProcess(child, 'SIGKILL')
      finish(null)
    }, 1_000)
    timer.unref()

    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.on('error', () => {
      finish(null)
    })
    child.on('close', (code) => {
      if (code !== 0) {
        finish(null)
        return
      }
      const rssKb = Number(stdout.trim())
      finish(Number.isFinite(rssKb) && rssKb >= 0 ? rssKb * 1024 : null)
    })
  })
}
