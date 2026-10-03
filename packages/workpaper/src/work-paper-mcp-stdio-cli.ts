export interface WorkPaperMcpStdioCliOptions {
  readonly demoWorkPaperTools: boolean
  readonly fromXlsxPath?: string
  readonly initDemoWorkPaper: boolean
  readonly overwriteWorkPaper: boolean
  readonly workpaperPath?: string
  readonly writable: boolean
  readonly help: boolean
}

export function parseWorkPaperMcpStdioCliArgs(args: readonly string[]): WorkPaperMcpStdioCliOptions {
  let demoWorkPaperTools = false
  let fromXlsxPath: string | undefined
  let initDemoWorkPaper = false
  let overwriteWorkPaper = false
  let workpaperPath: string | undefined
  let writable = false
  let help = false

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === '--help' || arg === '-h') {
      help = true
      continue
    }
    if (arg === '--writable') {
      writable = true
      continue
    }
    if (arg === '--demo-workpaper-tools') {
      demoWorkPaperTools = true
      continue
    }
    if (arg === '--init-demo-workpaper') {
      initDemoWorkPaper = true
      continue
    }
    if (arg === '--overwrite-workpaper') {
      overwriteWorkPaper = true
      continue
    }
    if (arg === '--workpaper') {
      const next = args[index + 1]
      if (next === undefined || next.trim().length === 0 || next.startsWith('-')) {
        throw new Error('--workpaper requires a path')
      }
      workpaperPath = next
      index += 1
      continue
    }
    if (arg === '--from-xlsx') {
      const next = args[index + 1]
      if (next === undefined || next.trim().length === 0 || next.startsWith('-')) {
        throw new Error('--from-xlsx requires a path')
      }
      fromXlsxPath = next
      index += 1
      continue
    }
    throw new Error(`Unknown bilig-workpaper-mcp argument: ${arg}`)
  }

  if (demoWorkPaperTools && workpaperPath !== undefined) {
    throw new Error('--demo-workpaper-tools cannot be combined with --workpaper')
  }
  if (demoWorkPaperTools && fromXlsxPath !== undefined) {
    throw new Error('--demo-workpaper-tools cannot be combined with --from-xlsx')
  }
  if (initDemoWorkPaper && workpaperPath === undefined) {
    throw new Error('--init-demo-workpaper requires --workpaper')
  }
  if (fromXlsxPath !== undefined && initDemoWorkPaper) {
    throw new Error('--from-xlsx cannot be combined with --init-demo-workpaper')
  }
  if (fromXlsxPath !== undefined && workpaperPath === undefined && writable) {
    throw new Error('--writable requires --workpaper when used with --from-xlsx')
  }
  if (overwriteWorkPaper && fromXlsxPath === undefined) {
    throw new Error('--overwrite-workpaper requires --from-xlsx')
  }
  if (overwriteWorkPaper && workpaperPath === undefined) {
    throw new Error('--overwrite-workpaper requires --workpaper')
  }

  if (workpaperPath === undefined) {
    if (fromXlsxPath === undefined) {
      return { demoWorkPaperTools, help, initDemoWorkPaper, overwriteWorkPaper, writable }
    }
    return { demoWorkPaperTools, fromXlsxPath, help, initDemoWorkPaper, overwriteWorkPaper, writable }
  }
  if (fromXlsxPath === undefined) {
    return { demoWorkPaperTools, help, initDemoWorkPaper, overwriteWorkPaper, writable, workpaperPath }
  }
  return { demoWorkPaperTools, fromXlsxPath, help, initDemoWorkPaper, overwriteWorkPaper, writable, workpaperPath }
}

export function workPaperMcpStdioHelpText(): string {
  return [
    'Usage: bilig-workpaper-mcp [--workpaper ./model.workpaper.json] [--init-demo-workpaper] [--writable]',
    '       bilig-workpaper-mcp --from-xlsx ./pricing.xlsx',
    '       bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json [--overwrite-workpaper] [--writable]',
    '       bilig-workpaper-mcp --demo-workpaper-tools',
    '',
    'Without --workpaper, starts the built-in demo WorkPaper MCP server.',
    '--demo-workpaper-tools starts the built-in demo workbook with the general WorkPaper tool surface.',
    'With --workpaper, loads a persisted WorkPaper JSON document and exposes file-backed tools, resources, and prompts.',
    '--init-demo-workpaper creates a demo WorkPaper JSON at --workpaper when the file is missing.',
    '--from-xlsx imports an existing XLSX into an in-memory WorkPaper MCP server unless --workpaper is also provided.',
    '--overwrite-workpaper allows --from-xlsx to replace an existing WorkPaper JSON file.',
    '--writable persists set_cell_contents edits back to the same JSON file.',
    '',
  ].join('\n')
}
