#!/usr/bin/env bun

import { spawnSync } from 'node:child_process'
import { resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { ensureWasmKernelArtifact } from './ensure-wasm-kernel.js'

export function buildVitestArgs(args: readonly string[], env: NodeJS.ProcessEnv = process.env): string[] {
  if (!env['BILIG_CI_PROFILE'] || hasArg(args, '--maxWorkers')) {
    return [...args]
  }
  return [...args, '--maxWorkers', env['BILIG_VITEST_MAX_WORKERS'] ?? '2']
}

function hasArg(args: readonly string[], flag: string): boolean {
  return args.some((arg) => arg === flag || arg.startsWith(`${flag}=`))
}

function main(): never {
  ensureWasmKernelArtifact()

  const vitestBin = process.platform === 'win32' ? 'node_modules\\.bin\\vitest.cmd' : 'node_modules/.bin/vitest'
  const result = spawnSync(vitestBin, buildVitestArgs(process.argv.slice(2)), {
    cwd: process.cwd(),
    env: process.env,
    stdio: 'inherit',
  })

  if (result.error) {
    throw result.error
  }

  if (result.signal) {
    process.stderr.write(`vitest terminated by signal ${result.signal}\n`)
  }

  process.exit(result.status ?? 1)
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
