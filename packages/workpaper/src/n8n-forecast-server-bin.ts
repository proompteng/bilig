#!/usr/bin/env node
import { runN8nForecastServerCli } from './n8n-forecast-server-cli.js'

process.exitCode = runN8nForecastServerCli({
  argv: process.argv.slice(2),
  env: process.env,
})
