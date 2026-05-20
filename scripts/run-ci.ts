#!/usr/bin/env bun

import { runCoverageContracts } from './coverage-contracts.ts'
import { assertLocalCiResourceGuardAllowsRun } from './ci-local-resource-guard.ts'
import { resolveCiProfile, resolveCiSkipBrowserGates } from './run-ci-config.ts'

import { spawn, type ChildProcess } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'

interface CiTask {
  readonly label: string
  readonly command?: readonly string[]
  readonly steps?: readonly CiTask[]
  readonly env?: Readonly<Record<string, string>>
  readonly execute?: () => Promise<void>
}

interface CompletedTask {
  readonly label: string
  readonly elapsedMs: number
}

const startedAt = performance.now()
const rootDir = fileURLToPath(new URL('..', import.meta.url))
const ciProfile = resolveCiProfileOrExit()
const runFullGates = ciProfile === 'full'
const runDeepGates = runFullGates
const skipBrowserGates = resolveCiSkipBrowserGatesOrExit()
const coverageReportsDirectory = process.env['BILIG_COVERAGE_DIR'] ?? `coverage/ci-${process.pid}`
const packageScripts = readPackageScripts()

process.env['BILIG_COVERAGE_DIR'] = coverageReportsDirectory

function resolveCiProfileOrExit(): ReturnType<typeof resolveCiProfile> {
  try {
    return resolveCiProfile(process.env)
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exit(1)
  }
}

function resolveCiSkipBrowserGatesOrExit(): boolean {
  try {
    return resolveCiSkipBrowserGates(process.env)
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exit(1)
  }
}

function formatSeconds(ms: number): string {
  return `${(ms / 1000).toFixed(1)}s`
}

function log(message: string): void {
  console.log(`[ci] ${message}`)
}

function pnpm(label: string, ...args: string[]): CiTask {
  return { label, command: ['pnpm', ...args] }
}

function git(label: string, ...args: string[]): CiTask {
  return { label, command: ['git', ...args] }
}

function direct(label: string, ...args: string[]): CiTask {
  return { label, command: args }
}

function directPackageScript(label: string, scriptName: string): CiTask {
  return direct(label, ...packageScriptCommand(scriptName))
}

function bunScript(label: string, script: string, ...args: string[]): CiTask {
  return direct(label, 'bun', script, ...args)
}

function tsxScript(label: string, script: string, ...args: string[]): CiTask {
  return direct(label, workspaceBin('tsx'), script, ...args)
}

function workspaceBin(name: string): string {
  return process.platform === 'win32' ? `node_modules\\.bin\\${name}.cmd` : `node_modules/.bin/${name}`
}

function readPackageScripts(): Readonly<Record<string, string>> {
  const packageJson: unknown = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'))
  if (!isStringRecordContainer(packageJson, 'scripts')) {
    return {}
  }
  return packageJson.scripts
}

function packageScriptCommand(scriptName: string): readonly string[] {
  const command = packageScripts[scriptName]
  if (!command) {
    throw new Error(`missing package script: ${scriptName}`)
  }
  const unsupportedShellTokens = new Set(['&&', '||', ';', '|', '>', '<'])
  const tokens = command
    .trim()
    .split(/\s+/u)
    .filter((token) => token.length > 0)
  const unsupportedToken = tokens.find((token) => unsupportedShellTokens.has(token))
  if (unsupportedToken) {
    throw new Error(`package script ${scriptName} uses unsupported shell token ${unsupportedToken}`)
  }
  return tokens
}

function isStringRecordContainer(value: unknown, key: string): value is Record<string, Record<string, string>> {
  if (!isRecord(value) || !isRecord(value[key])) {
    return false
  }
  return Object.values(value[key]).every((entry) => typeof entry === 'string')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function withEnv(task: CiTask, env: Readonly<Record<string, string>>): CiTask {
  return {
    ...task,
    env: {
      ...task.env,
      ...env,
    },
  }
}

async function runCoverageTask(task: Omit<CiTask, 'command' | 'steps'>): Promise<CompletedTask> {
  const taskStartedAt = performance.now()
  log(`start ${task.label}`)
  await task.execute?.()
  const elapsedMs = performance.now() - taskStartedAt
  log(`done ${task.label} in ${formatSeconds(elapsedMs)}`)
  return { label: task.label, elapsedMs }
}

async function runTask(task: CiTask, runningChildren: Set<ChildProcess>): Promise<CompletedTask> {
  if (task.steps) {
    const taskStartedAt = performance.now()
    log(`start ${task.label}`)
    await task.steps.reduce(async (previous, step) => {
      await previous
      await runTask(withEnv(step, task.env ?? {}), runningChildren)
    }, Promise.resolve())
    const elapsedMs = performance.now() - taskStartedAt
    log(`done ${task.label} in ${formatSeconds(elapsedMs)}`)
    return { label: task.label, elapsedMs }
  }

  if (task.execute) {
    return runCoverageTask(task)
  }

  if (!task.command) {
    throw new Error(`${task.label} has no command or steps`)
  }

  return new Promise((resolve, reject) => {
    const taskStartedAt = performance.now()
    log(`start ${task.label}: ${task.command.join(' ')}`)
    let settled = false
    let closeFallback: ReturnType<typeof setTimeout> | null = null
    function finish(code: number | null, signal: string | null): void {
      if (settled) {
        return
      }
      settled = true
      if (closeFallback) {
        clearTimeout(closeFallback)
        closeFallback = null
      }
      runningChildren.delete(child)
      const elapsedMs = performance.now() - taskStartedAt
      if (code === 0) {
        log(`done ${task.label} in ${formatSeconds(elapsedMs)}`)
        resolve({ label: task.label, elapsedMs })
        return
      }
      reject(new Error(`${task.label} failed after ${formatSeconds(elapsedMs)} (${signal ? `signal ${signal}` : `exit ${String(code)}`})`))
    }
    const child = spawn(task.command[0] ?? '', task.command.slice(1), {
      cwd: process.cwd(),
      env: {
        ...process.env,
        ...task.env,
      },
      stdio: 'inherit',
    })
    runningChildren.add(child)

    child.once('error', (error) => {
      if (closeFallback) {
        clearTimeout(closeFallback)
        closeFallback = null
      }
      settled = true
      runningChildren.delete(child)
      reject(error)
    })

    child.once('exit', (code, signal) => {
      closeFallback = setTimeout(() => {
        finish(code, signal)
      }, 1_000)
    })

    child.once('close', (code, signal) => {
      finish(code, signal)
    })
  })
}

function stopRunningChildren(runningChildren: Set<ChildProcess>): void {
  for (const child of runningChildren) {
    if (child.exitCode === null && child.signalCode === null) {
      child.kill('SIGTERM')
    }
  }
}

async function runStage(label: string, tasks: readonly CiTask[]): Promise<CompletedTask[]> {
  const stageStartedAt = performance.now()
  const runningChildren = new Set<ChildProcess>()
  log(`stage ${label} (${String(tasks.length)} task${tasks.length === 1 ? '' : 's'})`)
  try {
    const completed = await Promise.all(tasks.map((task) => runTask(task, runningChildren)))
    log(`stage ${label} done in ${formatSeconds(performance.now() - stageStartedAt)}`)
    return completed
  } catch (error) {
    stopRunningChildren(runningChildren)
    throw error
  }
}

async function runSequential(label: string, tasks: readonly CiTask[]): Promise<CompletedTask[]> {
  const stageStartedAt = performance.now()
  const completed: CompletedTask[] = []
  log(`stage ${label} (${String(tasks.length)} sequential task${tasks.length === 1 ? '' : 's'})`)
  await tasks.reduce(async (previous, task) => {
    await previous
    const runningChildren = new Set<ChildProcess>()
    completed.push(await runTask(task, runningChildren))
  }, Promise.resolve())
  log(`stage ${label} done in ${formatSeconds(performance.now() - stageStartedAt)}`)
  return completed
}

const coverageLane: CiTask = {
  label: 'coverage + contracts',
  steps: [
    pnpm('coverage', 'coverage'),
    {
      label: 'coverage contracts',
      execute: async () => {
        await runCoverageContracts()
      },
    },
  ],
}
const fuzzScript = runDeepGates ? 'test:fuzz:main' : 'test:fuzz'
const vitestFuzzLane = withEnv(pnpm(runDeepGates ? 'vitest fuzz main' : 'vitest fuzz default', fuzzScript), {
  BILIG_FUZZ_SKIP_BROWSER: '1',
})
const browserWebBundleBuild = withEnv(pnpm('browser web bundle build', '--filter', '@bilig/web', 'build:bundle'), {
  VITE_BILIG_REMOTE_SYNC: '0',
})
const appRuntimeDependencyBuild = pnpm('app runtime dependency build', '--filter', '@bilig/app^...', 'run', 'build')
const wasmBuildTask: CiTask = {
  label: 'wasm build',
  steps: [
    bunScript('wasm assembly build', 'packages/wasm-kernel/scripts/build.ts'),
    direct('wasm TypeScript build', workspaceBin('tsc'), '-p', 'packages/wasm-kernel/tsconfig.json'),
  ],
}
const browserLane: CiTask = {
  label: runFullGates ? 'browser tests + perf + fuzz' : 'browser ci smoke tests',
  steps: [
    withEnv(pnpm('browser tests', 'test:browser'), {
      BILIG_BROWSER_CI_SMOKE: runFullGates ? '0' : '1',
      BILIG_BROWSER_INCLUDE_PERF: runDeepGates ? '1' : '0',
      BILIG_BROWSER_INCLUDE_DEEP: runDeepGates ? '1' : '0',
      BILIG_BROWSER_INCLUDE_FUZZ: runDeepGates ? '1' : '0',
      BILIG_DEV_APP_RUNTIME_BUILD: '0',
      BILIG_DEV_WEB_PREVIEW_BUILD: '0',
      BILIG_FUZZ_PROFILE: 'main',
      BILIG_FUZZ_CAPTURE: '1',
    }),
  ],
}
const parallelFocusedCorrectnessLanes: readonly CiTask[] = [
  withEnv(directPackageScript('correctness core', 'test:correctness:core'), { BILIG_VITEST_FILE_CHUNK_SIZE: '10' }),
  withEnv(directPackageScript('correctness formula', 'test:correctness:formula'), { BILIG_VITEST_FILE_CHUNK_SIZE: '3' }),
  directPackageScript('correctness server', 'test:correctness:server'),
  directPackageScript('correctness browser runtime', 'test:correctness:browser'),
]
const corpusCorrectnessLane = withEnv(directPackageScript('correctness public workbook corpus', 'test:correctness:corpus'), {
  BILIG_VITEST_FILE_CHUNK_SIZE: '10',
})
const generatedSourceChecks: readonly CiTask[] = [
  bunScript('protocol check', 'scripts/gen-protocol.ts', '--check'),
  direct('protocol package build for generated-source imports', workspaceBin('tsc'), '-p', 'packages/protocol/tsconfig.json'),
  direct('agent API package build for generated-source imports', workspaceBin('tsc'), '-b', 'packages/agent-api/tsconfig.json'),
  bunScript('formula inventory check', 'scripts/gen-formula-inventory.ts', '--check'),
  bunScript('formula dominance check', 'scripts/gen-formula-dominance-snapshot.ts', '--check'),
  bunScript('calculation semantics scorecard check', 'scripts/gen-calculation-semantics-scorecard.ts', '--check'),
  bunScript('Microsoft Excel live calculation scorecard check', 'scripts/gen-microsoft-excel-live-calculation-scorecard.ts', '--check'),
  bunScript('Google Sheets live calculation scorecard check', 'scripts/gen-google-sheets-live-calculation-scorecard.ts', '--check'),
  bunScript('Microsoft Excel live recalculation scorecard check', 'scripts/gen-microsoft-excel-live-recalculation-scorecard.ts', '--check'),
  bunScript('Google Sheets live recalculation scorecard check', 'scripts/gen-google-sheets-live-recalculation-scorecard.ts', '--check'),
  bunScript('Microsoft Excel live structural scorecard check', 'scripts/gen-microsoft-excel-live-structural-scorecard.ts', '--check'),
  bunScript('Google Sheets live structural scorecard check', 'scripts/gen-google-sheets-live-structural-scorecard.ts', '--check'),
  bunScript(
    'Microsoft Excel live large workbook scorecard check',
    'scripts/gen-microsoft-excel-live-large-workbook-scorecard.ts',
    '--check',
  ),
  bunScript('Google Sheets live large workbook scorecard check', 'scripts/gen-google-sheets-live-large-workbook-scorecard.ts', '--check'),
  bunScript('auditability scorecard check', 'scripts/gen-auditability-scorecard.ts', '--check'),
  bunScript('reliability scorecard check', 'scripts/gen-reliability-scorecard.ts', '--check'),
  bunScript('collaboration scorecard check', 'scripts/gen-collaboration-scorecard.ts', '--check'),
  bunScript('automation scorecard check', 'scripts/gen-automation-scorecard.ts', '--check'),
  bunScript('import/export fidelity scorecard check', 'scripts/gen-import-export-fidelity-scorecard.ts', '--check'),
  bunScript(
    'public workbook corpus shared-link lifecycle plan check',
    'scripts/public-workbook-corpus.ts',
    'link-plan',
    '--source-url',
    'https://docs.google.com/spreadsheets/d/biligSharedWorkbookCheck/edit?usp=sharing',
    '--license-title',
    'Creative Commons Attribution 4.0 International',
    '--license-url',
    'https://creativecommons.org/licenses/by/4.0/',
    '--license-spdx',
    'CC-BY-4.0',
  ),
  bunScript(
    'public workbook corpus shared-link intake check',
    'scripts/public-workbook-corpus.ts',
    'add-link',
    '--dry-run',
    '--source-url',
    'https://docs.google.com/spreadsheets/d/biligSharedWorkbookCheck/edit?usp=sharing',
    '--license-title',
    'Creative Commons Attribution 4.0 International',
    '--license-url',
    'https://creativecommons.org/licenses/by/4.0/',
    '--license-spdx',
    'CC-BY-4.0',
  ),
  bunScript('public workbook corpus offline scorecard check', 'scripts/public-workbook-corpus.ts', 'check', '--skip-manifest-check'),
  bunScript('public workbook corpus resume plan check', 'scripts/public-workbook-corpus-resume-plan.ts', '--check'),
  bunScript('public workbook corpus resource-limit plan check', 'scripts/public-workbook-corpus-resource-limit-plan.ts', '--check'),
  bunScript('public workbook corpus feature-witness plan check', 'scripts/public-workbook-corpus-feature-witness-plan.ts', '--check'),
  bunScript('financial public workbook corpus plan check', 'scripts/public-workbook-corpus-financial-plan.ts', '--check'),
  directPackageScript('financial public workbook corpus resume check', 'public-workbook-corpus:resume-financial:check'),
  bunScript('public workbook corpus completion audit check', 'scripts/public-workbook-corpus-completion-audit.ts', '--check'),
  bunScript('large workbook SLO scorecard check', 'scripts/gen-large-workbook-slo-scorecard.ts', '--check'),
  {
    label: 'WorkPaper XLSX corpus fixture check',
    steps: [
      bunScript('WorkPaper XLSX corpus fixture generation check', 'scripts/gen-workpaper-xlsx-corpus-fixtures.ts', '--check'),
      bunScript(
        'WorkPaper XLSX corpus parity check',
        'scripts/check-workpaper-xlsx-corpus.ts',
        '--',
        'packages/headless/fixtures/xlsx-corpus',
      ),
    ],
  },
  bunScript(
    'UI same-corpus XLSX fixture check',
    'scripts/capture-ui-responsiveness-same-corpus.ts',
    '--emit-xlsx',
    'packages/benchmarks/baselines/ui-same-corpus',
    '--check',
  ),
  bunScript('UI responsiveness live browser scorecard check', 'scripts/gen-ui-responsiveness-live-browser-scorecard.ts', '--check'),
  bunScript('security posture scorecard check', 'scripts/gen-security-posture-scorecard.ts', '--check'),
  tsxScript('WorkPaper TrueCalc scalar benchmark check', 'scripts/gen-workpaper-vs-truecalc-benchmark.ts', '--check'),
  tsxScript('WorkPaper xlsx-calc benchmark check', 'scripts/gen-workpaper-vs-xlsx-calc-benchmark.ts', '--check'),
  bunScript('headless performance leadership scorecard check', 'scripts/gen-headless-performance-leadership-scorecard.ts', '--check'),
  bunScript('bilig dominance scorecard check', 'scripts/gen-bilig-dominance-scorecard.ts', '--check'),
  bunScript('bilig dominance audit check', 'scripts/bilig-dominance-audit.ts', '--check'),
  bunScript('public claims check', 'scripts/check-public-claims.ts'),
  bunScript('workspace resolution check', 'scripts/gen-workspace-resolution.ts', '--check'),
  bunScript('canonical naming check', 'scripts/check-canonical-naming.ts'),
  tsxScript('docs hero asset check', 'scripts/render-hero-workbook-api.ts', '--check'),
  tsxScript('docs social preview asset check', 'scripts/render-social-preview.ts', '--check'),
  tsxScript('docs benchmark card asset check', 'scripts/render-benchmark-card.ts', '--check'),
  bunScript('public evidence check', 'scripts/sync-public-evidence.ts', '--check'),
  bunScript('headless package footprint check', 'scripts/sync-headless-package-footprint.ts', '--check'),
  bunScript('create WorkPaper package check', 'scripts/check-create-workpaper-package.ts'),
  bunScript('agent discovery docs check', 'scripts/sync-agent-discovery-docs.ts', '--check'),
  tsxScript('docs discovery check', 'scripts/check-docs-discovery.ts'),
]
const semanticFastGate = pnpm('semantic correctness fast gate', 'test:semantic:fast')

try {
  assertLocalCiResourceGuardAllowsRun(rootDir)
  const allCompleted: CompletedTask[] = []
  log(`profile ${ciProfile}`)
  if (!runFullGates) {
    log(
      'fast profile runs generated checks, static checks, focused correctness tests, browser smoke, release budgets, perf smoke, and clean-diff checks; run pnpm run ci:full for coverage, fuzz, full browser, and deep benchmark gates',
    )
  }
  if (skipBrowserGates) {
    log('browser gates disabled by BILIG_CI_SKIP_BROWSER=1')
  }

  allCompleted.push(
    ...(await runSequential('static package build prerequisites', [
      skipBrowserGates ? wasmBuildTask : appRuntimeDependencyBuild,
      ...(skipBrowserGates ? [] : [pnpm('playwright chromium install', 'exec', 'playwright', 'install', 'chromium')]),
    ])),
  )

  // Keep generated-source checks serialized; later checks read artifacts validated by earlier ones.
  allCompleted.push(...(await runSequential('generated-source checks', generatedSourceChecks)))
  allCompleted.push(...(await runSequential('semantic correctness checks', [semanticFastGate])))

  allCompleted.push(
    ...(await runSequential('static direct checks', [
      direct(
        'lint',
        workspaceBin('oxlint'),
        '--config',
        '.oxlintrc.json',
        '--type-aware',
        '--deny-warnings',
        'packages',
        'apps',
        'e2e',
        'scripts',
      ),
      direct('source size check', 'bun', 'scripts/check-source-file-size.ts'),
      direct('typecheck', workspaceBin('tsc'), '-b', '--pretty', 'false'),
    ])),
  )

  if (runFullGates) {
    allCompleted.push(
      ...(await runStage('functional heavy checks', [
        {
          label: 'vitest heavy checks',
          steps: [coverageLane, vitestFuzzLane],
        },
        ...(skipBrowserGates ? [] : [browserWebBundleBuild]),
      ])),
    )
  } else {
    // Keep Vitest lanes serialized locally; running four pnpm/vitest processes concurrently is prone to child-process
    // termination before assertion output on constrained machines.
    allCompleted.push(...(await runSequential('focused correctness checks', parallelFocusedCorrectnessLanes)))
    allCompleted.push(...(await runSequential('corpus correctness benchmark', [corpusCorrectnessLane])))
    if (!skipBrowserGates) {
      allCompleted.push(...(await runStage('browser smoke setup', [browserWebBundleBuild])))
    }
  }

  if (!skipBrowserGates) {
    allCompleted.push(...(await runSequential('browser gates', [browserLane])))
  }

  allCompleted.push(
    ...(await runSequential('release bundle gate', [
      pnpm('production web bundle build', '--filter', '@bilig/web', 'build:bundle'),
      pnpm('release check', 'release:check'),
    ])),
  )

  allCompleted.push(
    ...(await runSequential('performance and clean-diff gates', [
      pnpm('perf smoke', 'bench:smoke'),
      directPackageScript('public workbook corpus synthetic memory gate', 'public-workbook-corpus:memory-gate:synthetic'),
      ...(runDeepGates ? [withEnv(pnpm('benchmark contracts', 'bench:contracts'), { CI: '1' })] : []),
      git('working tree clean', 'diff', '--exit-code'),
      git('index clean', 'diff', '--cached', '--exit-code'),
    ])),
  )

  const orderedByDuration = [...allCompleted].toSorted((left, right) => right.elapsedMs - left.elapsedMs)
  log(`all checks passed in ${formatSeconds(performance.now() - startedAt)}`)
  log(
    `slowest tasks: ${orderedByDuration
      .slice(0, 8)
      .map((task) => `${task.label} ${formatSeconds(task.elapsedMs)}`)
      .join(', ')}`,
  )
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error))
  process.exit(1)
}
