#!/usr/bin/env bun

import { spawn, type ChildProcess } from 'node:child_process'

interface CiTask {
  readonly label: string
  readonly command?: readonly string[]
  readonly steps?: readonly CiTask[]
  readonly env?: Readonly<Record<string, string>>
}

interface CompletedTask {
  readonly label: string
  readonly elapsedMs: number
}

const startedAt = performance.now()
const ciProfile = process.env['BILIG_CI_PROFILE'] === 'full' ? 'full' : 'default'
const runDeepGates = ciProfile === 'full'
const skipBrowserGates = process.env['BILIG_CI_SKIP_BROWSER'] === '1'

function formatSeconds(ms: number): string {
  return `${(ms / 1000).toFixed(1)}s`
}

function log(message: string): void {
  console.log(`[ci] ${message}`)
}

function pnpm(label: string, ...args: string[]): CiTask {
  return { label, command: ['pnpm', ...args] }
}

function bun(label: string, ...args: string[]): CiTask {
  return { label, command: ['bun', ...args] }
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
  steps: [pnpm('coverage', 'coverage'), bun('coverage contracts', 'scripts/coverage-contracts.ts')],
}
const fuzzScript = runDeepGates ? 'test:fuzz:main' : 'test:fuzz'
const vitestFuzzLane = withEnv(pnpm(runDeepGates ? 'vitest fuzz main' : 'vitest fuzz default', fuzzScript), {
  BILIG_FUZZ_SKIP_BROWSER: '1',
})
const browserWebBundleBuild = withEnv(pnpm('browser web bundle build', '--filter', '@bilig/web', 'build:bundle'), {
  VITE_BILIG_REMOTE_SYNC: '0',
})
const browserLane: CiTask = {
  label: runDeepGates ? 'browser tests + perf + fuzz' : 'browser tests',
  steps: [
    withEnv(pnpm('browser tests', 'test:browser'), {
      BILIG_DEV_WEB_PREVIEW_BUILD: '0',
      BILIG_BROWSER_INCLUDE_PERF: runDeepGates ? '1' : '0',
      BILIG_BROWSER_INCLUDE_DEEP: runDeepGates ? '1' : '0',
      BILIG_BROWSER_INCLUDE_FUZZ: runDeepGates ? '1' : '0',
      BILIG_FUZZ_PROFILE: 'main',
      BILIG_FUZZ_CAPTURE: '1',
    }),
  ],
}

try {
  const allCompleted: CompletedTask[] = []
  log(`profile ${ciProfile}`)
  if (!runDeepGates) {
    log(
      'default profile uses the fast fuzz budget and skips browser perf, browser fuzz, and statistical benchmark contracts; run pnpm run ci:full for the deep gate',
    )
  }
  if (skipBrowserGates) {
    log('browser gates disabled by BILIG_CI_SKIP_BROWSER=1')
  }

  allCompleted.push(
    ...(await runStage('generated-source checks', [
      pnpm('protocol check', 'protocol:check'),
      pnpm('formula inventory check', 'formula-inventory:check'),
      pnpm('formula dominance check', 'formula:dominance:check'),
      pnpm('workspace resolution check', 'workspace-resolution:check'),
      pnpm('canonical naming check', 'naming:check'),
    ])),
  )

  allCompleted.push(
    ...(await runStage('static prerequisites', [
      pnpm('lint', 'lint'),
      pnpm('wasm build', 'wasm:build'),
      pnpm('typecheck', 'typecheck'),
      ...(skipBrowserGates ? [] : [pnpm('playwright chromium install', 'exec', 'playwright', 'install', 'chromium')]),
    ])),
  )

  allCompleted.push(
    ...(await runStage('functional heavy checks', [coverageLane, vitestFuzzLane, ...(skipBrowserGates ? [] : [browserWebBundleBuild])])),
  )

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
