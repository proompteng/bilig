#!/usr/bin/env node
import { runAgentStartCli } from './agent-start-cli.js'

process.exitCode = runAgentStartCli({
  argv: process.argv.slice(2),
})
