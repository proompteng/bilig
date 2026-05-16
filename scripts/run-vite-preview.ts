#!/usr/bin/env bun

import { parseVitePreviewCliArgs } from './vite-preview-cli.js'

const { port, host } = parseVitePreviewCliArgs(process.argv.slice(2))

const child = Bun.spawn(['pnpm', 'exec', 'vite', 'preview', '--host', host, '--port', String(port), '--strictPort'], {
  cwd: process.cwd(),
  stdin: 'ignore',
  stdout: 'inherit',
  stderr: 'inherit',
})

function forwardAndExit(signal: NodeJS.Signals): void {
  try {
    child.kill(signal)
  } catch (error) {
    console.warn('Failed to forward signal to preview process', String(signal), error)
  }
}

process.on('SIGINT', () => forwardAndExit('SIGINT'))
process.on('SIGTERM', () => forwardAndExit('SIGTERM'))

process.exit(await child.exited)
