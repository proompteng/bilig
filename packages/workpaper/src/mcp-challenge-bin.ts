#!/usr/bin/env node
import { runMcpChallengeCli } from '@bilig/headless/cli'

process.exitCode = runMcpChallengeCli({
  argv: process.argv.slice(2),
})
