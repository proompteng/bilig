#!/usr/bin/env node
import { runAgentWorkbookChallengeCli } from '@bilig/headless/cli'

process.exitCode = runAgentWorkbookChallengeCli({
  argv: process.argv.slice(2),
})
