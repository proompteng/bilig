#!/usr/bin/env node
import { createFileBackedWorkPaperMcpToolServer, createFileBackedWorkPaperMcpToolServerFromFile } from './work-paper-mcp-file-server.js'
import { parseWorkPaperMcpStdioCliArgs, workPaperMcpStdioHelpText } from './work-paper-mcp-stdio-cli.js'
import { runDemoWorkPaperMcpStdioServer } from './work-paper-mcp-stdio-server.js'
import { buildDemoWorkPaper } from './work-paper-mcp-server.js'

const cliOptions = parseWorkPaperMcpStdioCliArgs(process.argv.slice(2))
if (cliOptions.help) {
  process.stdout.write(workPaperMcpStdioHelpText())
  process.exit(0)
}

if (cliOptions.demoWorkPaperTools) {
  runDemoWorkPaperMcpStdioServer({
    server: createFileBackedWorkPaperMcpToolServer({
      workbook: buildDemoWorkPaper(),
      sourcePath: 'demo://bilig-workpaper',
      writable: false,
    }),
  })
} else if (cliOptions.workpaperPath === undefined) {
  runDemoWorkPaperMcpStdioServer()
} else {
  runDemoWorkPaperMcpStdioServer({
    server: createFileBackedWorkPaperMcpToolServerFromFile({
      initDemoWorkPaper: cliOptions.initDemoWorkPaper,
      workpaperPath: cliOptions.workpaperPath,
      writable: cliOptions.writable,
    }),
  })
}
