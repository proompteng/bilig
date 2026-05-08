#!/usr/bin/env bun

import { runCoverageContracts } from './coverage-contracts.ts'

import { spawn, type ChildProcess } from 'node:child_process'

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
const ciProfile = process.env['BILIG_CI_PROFILE'] === 'full' ? 'full' : 'fast'
const runFullGates = ciProfile === 'full'
const runDeepGates = runFullGates
const skipBrowserGates = process.env['BILIG_CI_SKIP_BROWSER'] === '1'
const coverageReportsDirectory = process.env['BILIG_COVERAGE_DIR'] ?? `coverage/ci-${process.pid}`

process.env['BILIG_COVERAGE_DIR'] = coverageReportsDirectory

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
    const child = spawn(task.command[0] ?? '', task.command.slice(1), {
      cwd: process.cwd(),
      env: {
        ...process.env,
        ...task.env,
      },
      stdio: 'inherit',
    })
    runningChildren.add(child)

    child.on('error', (error) => {
      runningChildren.delete(child)
      reject(error)
    })

    child.on('close', (code, signal) => {
      runningChildren.delete(child)
      const elapsedMs = performance.now() - taskStartedAt
      if (code === 0) {
        log(`done ${task.label} in ${formatSeconds(elapsedMs)}`)
        resolve({ label: task.label, elapsedMs })
        return
      }
      reject(new Error(`${task.label} failed after ${formatSeconds(elapsedMs)} (${signal ? `signal ${signal}` : `exit ${String(code)}`})`))
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
  pnpm('correctness core', 'test:correctness:core'),
  pnpm('correctness formula', 'test:correctness:formula'),
  pnpm('correctness server', 'test:correctness:server'),
  pnpm('correctness browser runtime', 'test:correctness:browser'),
]
const corpusCorrectnessLane = pnpm('correctness public workbook corpus', 'test:correctness:corpus')
const generatedSourceChecks: readonly CiTask[] = [
  pnpm('protocol check', 'protocol:check'),
  pnpm('protocol package build for generated-source imports', '--filter', '@bilig/protocol', 'build'),
  pnpm('agent API package build for generated-source imports', '--filter', '@bilig/agent-api', 'build'),
  pnpm('formula inventory check', 'formula-inventory:check'),
  pnpm('formula dominance check', 'formula:dominance:check'),
  pnpm('calculation semantics scorecard check', 'calculation:semantics:check'),
  pnpm('Microsoft Excel live calculation scorecard check', 'calculation:excel-live:check'),
  pnpm('Google Sheets live calculation scorecard check', 'calculation:google-sheets-live:check'),
  pnpm('Microsoft Excel live recalculation scorecard check', 'recalculation:excel-live:check'),
  pnpm('Google Sheets live recalculation scorecard check', 'recalculation:google-sheets-live:check'),
  pnpm('Microsoft Excel live structural scorecard check', 'structural:excel-live:check'),
  pnpm('Google Sheets live structural scorecard check', 'structural:google-sheets-live:check'),
  pnpm('Microsoft Excel live large workbook scorecard check', 'large-workbook:excel-live:check'),
  pnpm('Google Sheets live large workbook scorecard check', 'large-workbook:google-sheets-live:check'),
  pnpm('auditability scorecard check', 'auditability:check'),
  pnpm('reliability scorecard check', 'reliability:check'),
  pnpm('collaboration scorecard check', 'collaboration:check'),
  pnpm('automation scorecard check', 'automation:check'),
  pnpm('import/export fidelity scorecard check', 'import-export:fidelity:check'),
  pnpm('public workbook corpus shared-link lifecycle plan check', 'public-workbook-corpus:link-plan:check'),
  pnpm('public workbook corpus shared-link intake check', 'public-workbook-corpus:add-link:check'),
  pnpm('public workbook corpus resume plan check', 'public-workbook-corpus:resume-plan:check'),
  pnpm('public workbook corpus completion audit check', 'public-workbook-corpus:completion-audit:check'),
  pnpm('large workbook SLO scorecard check', 'large-workbook:slo:check'),
  pnpm('WorkPaper XLSX corpus fixture check', 'workpaper:xlsx-corpus:fixtures:check'),
  pnpm('UI same-corpus XLSX fixture check', 'ui:same-corpus:fixture:check'),
  pnpm('UI responsiveness live browser scorecard check', 'ui:browser-live:check'),
  pnpm('security posture scorecard check', 'security:posture:check'),
  pnpm('bilig dominance scorecard check', 'dominance:check'),
  pnpm('workspace resolution check', 'workspace-resolution:check'),
  pnpm('canonical naming check', 'naming:check'),
  pnpm('docs discovery check', 'docs:discovery:check'),
]

try {
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

  // Keep pnpm generated-source checks serialized; parallel pnpm invocations can race on .pnpm-workspace-state-v1.json.
  allCompleted.push(...(await runSequential('generated-source checks', generatedSourceChecks)))

  allCompleted.push(
    ...(await runStage('static prerequisites', [
      pnpm('lint', 'lint'),
      skipBrowserGates ? pnpm('wasm build', 'wasm:build') : appRuntimeDependencyBuild,
      pnpm('typecheck', 'typecheck'),
      ...(skipBrowserGates ? [] : [pnpm('playwright chromium install', 'exec', 'playwright', 'install', 'chromium')]),
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
    allCompleted.push(...(await runStage('focused correctness checks', parallelFocusedCorrectnessLanes)))
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
