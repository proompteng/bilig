#!/usr/bin/env node
import { runMcpChallengeCli } from './mcp-challenge-cli.js'

process.exitCode = runMcpChallengeCli({
  argv: process.argv.slice(2),
})
