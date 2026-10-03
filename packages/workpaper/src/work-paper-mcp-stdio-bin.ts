#!/usr/bin/env node
import {
  buildDemoWorkPaper,
  createFileBackedWorkPaperMcpToolServer,
  createFileBackedWorkPaperMcpToolServerFromFile,
  createFileBackedWorkPaperMcpToolServerFromXlsxFile,
  createWorkPaperMcpToolServerFromXlsxFile,
  parseWorkPaperMcpStdioCliArgs,
  runDemoWorkPaperMcpStdioServer,
  workPaperMcpStdioHelpText,
} from './mcp.js'
import { withXlsxWorkbookRiskTool } from './work-paper-mcp-xlsx-risk-tool.js'

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
} else if (cliOptions.fromXlsxPath !== undefined) {
  const server =
    cliOptions.workpaperPath === undefined
      ? createWorkPaperMcpToolServerFromXlsxFile({
          fromXlsxPath: cliOptions.fromXlsxPath,
        })
      : createFileBackedWorkPaperMcpToolServerFromXlsxFile({
          fromXlsxPath: cliOptions.fromXlsxPath,
          overwriteWorkPaper: cliOptions.overwriteWorkPaper,
          workpaperPath: cliOptions.workpaperPath,
          writable: cliOptions.writable,
        })
  runDemoWorkPaperMcpStdioServer({
    server: withXlsxWorkbookRiskTool(server, { xlsxPath: cliOptions.fromXlsxPath }),
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
