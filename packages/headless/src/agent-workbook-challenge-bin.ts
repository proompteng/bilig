#!/usr/bin/env node
import { runAgentWorkbookChallengeCli } from './agent-workbook-challenge-cli.js'

process.exitCode = runAgentWorkbookChallengeCli({
  argv: process.argv.slice(2),
})
