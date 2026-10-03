#!/usr/bin/env node
import { runMcpChallengeCli } from './cli.js'

process.exitCode = runMcpChallengeCli({
  argv: process.argv.slice(2),
})
