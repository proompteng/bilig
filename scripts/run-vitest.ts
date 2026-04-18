#!/usr/bin/env bun

import { spawnSync } from 'node:child_process'
import { ensureWasmKernelArtifact } from './ensure-wasm-kernel.js'

ensureWasmKernelArtifact()

const result = spawnSync('pnpm', ['exec', 'vitest', ...process.argv.slice(2)], {
  cwd: process.cwd(),
  env: process.env,
  stdio: 'inherit',
})

process.exit(result.status ?? 1)
