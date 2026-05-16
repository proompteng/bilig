export const CLIPBOARD_GLOBAL_GREP = '@clipboard-global'
export const BROWSER_CI_GREP = '@browser-ci'
export const BROWSER_PERF_GREP = '@browser-perf'
export const BROWSER_DEEP_GREP = '@browser-deep'
export const BROWSER_SERIAL_GREP = '@browser-serial'
export const BROWSER_FUZZ_GREP = '@fuzz-browser'
export const BROWSER_WEBGPU_GREP = '@browser-webgpu'
const WEBGPU_BROWSER_ENV = { BILIG_BROWSER_WEBGPU: '1' } as const
const WEBGPU_PERF_GREP = `${BROWSER_WEBGPU_GREP}.*${BROWSER_PERF_GREP}|${BROWSER_PERF_GREP}.*${BROWSER_WEBGPU_GREP}`
const WEBGPU_DEEP_GREP = `${BROWSER_WEBGPU_GREP}.*${BROWSER_DEEP_GREP}|${BROWSER_DEEP_GREP}.*${BROWSER_WEBGPU_GREP}`
const DEFAULT_PARALLEL_BROWSER_WORKERS = 2

export interface BrowserTestPhase {
  readonly label: string
  readonly args: readonly string[]
  readonly env?: Readonly<Record<string, string>>
}

export interface BrowserTestPhaseEnv {
  readonly BILIG_BROWSER_CI_SMOKE?: string | undefined
  readonly BILIG_BROWSER_INCLUDE_PERF?: string | undefined
  readonly BILIG_BROWSER_INCLUDE_DEEP?: string | undefined
  readonly BILIG_BROWSER_INCLUDE_FUZZ?: string | undefined
  readonly BILIG_BROWSER_PARALLEL_WORKERS?: string | undefined
  readonly BILIG_FUZZ_PROFILE?: string | undefined
  readonly BILIG_FUZZ_CAPTURE?: string | undefined
}

function envFlagEnabled(value: string | undefined, name: string): boolean {
  if (value === undefined || value.length === 0) {
    return false
  }
  const normalized = value.trim().toLowerCase()
  if (normalized === '1' || normalized === 'true') {
    return true
  }
  if (normalized === '0' || normalized === 'false') {
    return false
  }
  throw new Error(`${name} must be a boolean value, got ${value}`)
}

function resolveParallelBrowserWorkers(value: string | undefined): number {
  if (value === undefined || value.length === 0) {
    return DEFAULT_PARALLEL_BROWSER_WORKERS
  }
  if (!/^(?:[1-9]\d*)$/u.test(value)) {
    throw new Error(`BILIG_BROWSER_PARALLEL_WORKERS must be a positive integer, got ${value}`)
  }
  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed)) {
    throw new Error(`BILIG_BROWSER_PARALLEL_WORKERS must be a safe integer, got ${value}`)
  }
  return parsed
}

export function resolveBrowserTestPhases(input: {
  readonly playwrightArgs: readonly string[]
  readonly env: BrowserTestPhaseEnv
}): BrowserTestPhase[] {
  if (input.playwrightArgs.length > 0) {
    return [
      {
        label: 'explicit browser tests',
        args: input.playwrightArgs,
      },
    ]
  }

  const includePerf = envFlagEnabled(input.env.BILIG_BROWSER_INCLUDE_PERF, 'BILIG_BROWSER_INCLUDE_PERF')
  const includeDeep = envFlagEnabled(input.env.BILIG_BROWSER_INCLUDE_DEEP, 'BILIG_BROWSER_INCLUDE_DEEP')
  const includeFuzz = envFlagEnabled(input.env.BILIG_BROWSER_INCLUDE_FUZZ, 'BILIG_BROWSER_INCLUDE_FUZZ')
  const ciSmoke = envFlagEnabled(input.env.BILIG_BROWSER_CI_SMOKE, 'BILIG_BROWSER_CI_SMOKE')
  if (ciSmoke) {
    return [
      {
        label: 'browser ci smoke tests',
        args: [
          '--workers=' + String(resolveParallelBrowserWorkers(input.env.BILIG_BROWSER_PARALLEL_WORKERS)),
          '--grep',
          BROWSER_CI_GREP,
          '--grep-invert',
          [BROWSER_PERF_GREP, BROWSER_DEEP_GREP, BROWSER_FUZZ_GREP, BROWSER_WEBGPU_GREP].join('|'),
        ],
      },
    ]
  }

  const defaultExcludedGreps = [
    CLIPBOARD_GLOBAL_GREP,
    BROWSER_SERIAL_GREP,
    BROWSER_FUZZ_GREP,
    BROWSER_PERF_GREP,
    BROWSER_DEEP_GREP,
    BROWSER_WEBGPU_GREP,
  ]
  const phases: BrowserTestPhase[] = [
    {
      label: 'parallel browser tests',
      args: [
        '--workers=' + String(resolveParallelBrowserWorkers(input.env.BILIG_BROWSER_PARALLEL_WORKERS)),
        '--grep-invert',
        defaultExcludedGreps.join('|'),
      ],
    },
    {
      label: 'browser webgpu tests',
      args: ['--workers=1', '--grep', BROWSER_WEBGPU_GREP, '--grep-invert', [BROWSER_PERF_GREP, BROWSER_DEEP_GREP].join('|')],
      env: WEBGPU_BROWSER_ENV,
    },
  ]

  if (includePerf) {
    phases.push(
      {
        label: 'browser perf tests',
        args: ['--workers=1', '--grep', BROWSER_PERF_GREP, '--grep-invert', BROWSER_WEBGPU_GREP],
      },
      {
        label: 'browser webgpu perf tests',
        args: ['--workers=1', '--grep', WEBGPU_PERF_GREP],
        env: WEBGPU_BROWSER_ENV,
      },
    )
  }

  if (includeDeep) {
    phases.push(
      {
        label: 'browser deep tests',
        args: ['--workers=1', '--grep', BROWSER_DEEP_GREP, '--grep-invert', BROWSER_WEBGPU_GREP, '--pass-with-no-tests'],
      },
      {
        label: 'browser webgpu deep tests',
        args: ['--workers=1', '--grep', WEBGPU_DEEP_GREP],
        env: WEBGPU_BROWSER_ENV,
      },
    )
  }

  phases.push(
    {
      label: 'browser serial tests',
      args: ['--workers=1', '--grep', BROWSER_SERIAL_GREP, '--grep-invert', BROWSER_WEBGPU_GREP],
    },
    {
      label: 'clipboard global tests',
      args: ['--workers=1', '--grep', CLIPBOARD_GLOBAL_GREP],
    },
  )

  if (includeFuzz) {
    phases.push({
      label: 'browser fuzz tests',
      args: ['--workers=1', '--grep', BROWSER_FUZZ_GREP],
      env: {
        BILIG_FUZZ_BROWSER: '1',
        BILIG_FUZZ_PROFILE: input.env.BILIG_FUZZ_PROFILE ?? 'main',
        BILIG_FUZZ_CAPTURE: input.env.BILIG_FUZZ_CAPTURE ?? '1',
      },
    })
  }

  return phases
}
