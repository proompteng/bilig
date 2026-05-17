export interface WorkPaperMcpStdioCliOptions {
  readonly demoWorkPaperTools: boolean
  readonly workpaperPath?: string
  readonly writable: boolean
  readonly help: boolean
}

export function parseWorkPaperMcpStdioCliArgs(args: readonly string[]): WorkPaperMcpStdioCliOptions {
  let demoWorkPaperTools = false
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
    if (arg === '--workpaper') {
      const next = args[index + 1]
      if (next === undefined || next.trim().length === 0 || next.startsWith('-')) {
        throw new Error('--workpaper requires a path')
      }
      workpaperPath = next
      index += 1
      continue
    }
    throw new Error(`Unknown bilig-workpaper-mcp argument: ${arg}`)
  }

  if (demoWorkPaperTools && workpaperPath !== undefined) {
    throw new Error('--demo-workpaper-tools cannot be combined with --workpaper')
  }

  if (workpaperPath === undefined) {
    return { demoWorkPaperTools, help, writable }
  }
  return { demoWorkPaperTools, help, writable, workpaperPath }
}

export function workPaperMcpStdioHelpText(): string {
  return [
    'Usage: bilig-workpaper-mcp [--workpaper ./model.workpaper.json] [--writable]',
    '       bilig-workpaper-mcp --demo-workpaper-tools',
    '',
    'Without --workpaper, starts the built-in demo WorkPaper MCP server.',
    '--demo-workpaper-tools starts the built-in demo workbook with the general WorkPaper tool surface.',
    'With --workpaper, loads a persisted WorkPaper JSON document and exposes file-backed tools.',
    '--writable persists set_cell_contents edits back to the same JSON file.',
    '',
  ].join('\n')
}
