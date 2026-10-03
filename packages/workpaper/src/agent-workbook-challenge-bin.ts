#!/usr/bin/env node
import { runAgentWorkbookChallengeCli } from './cli.js'

process.exitCode = runAgentWorkbookChallengeCli({
  argv: process.argv.slice(2),
})
