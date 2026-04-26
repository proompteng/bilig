export const CLIPBOARD_GLOBAL_GREP = '@clipboard-global'
export const BROWSER_PERF_GREP = '@browser-perf'
export const BROWSER_SERIAL_GREP = '@browser-serial'
export const BROWSER_FUZZ_GREP = '@fuzz-browser'

export interface BrowserTestPhase {
  readonly label: string
  readonly args: readonly string[]
  readonly env?: Readonly<Record<string, string>>
}

export interface BrowserTestPhaseEnv {
  readonly BILIG_BROWSER_INCLUDE_PERF?: string | undefined
  readonly BILIG_BROWSER_INCLUDE_FUZZ?: string | undefined
  readonly BILIG_FUZZ_PROFILE?: string | undefined
  readonly BILIG_FUZZ_CAPTURE?: string | undefined
}

function envFlagEnabled(value: string | undefined): boolean {
  return value === '1' || value === 'true'
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

  const includePerf = envFlagEnabled(input.env.BILIG_BROWSER_INCLUDE_PERF)
  const includeFuzz = envFlagEnabled(input.env.BILIG_BROWSER_INCLUDE_FUZZ)
  const defaultExcludedGreps = [CLIPBOARD_GLOBAL_GREP, BROWSER_SERIAL_GREP, BROWSER_FUZZ_GREP, BROWSER_PERF_GREP]
  const phases: BrowserTestPhase[] = [
    {
      label: 'parallel browser tests',
      args: ['--grep-invert', defaultExcludedGreps.join('|')],
    },
  ]

  if (includePerf) {
    phases.push({
      label: 'browser perf tests',
      args: ['--workers=1', '--grep', BROWSER_PERF_GREP],
    })
  }

  phases.push(
    {
      label: 'browser serial tests',
      args: ['--workers=1', '--grep', BROWSER_SERIAL_GREP],
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
