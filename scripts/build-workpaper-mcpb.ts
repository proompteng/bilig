import { execFileSync } from 'node:child_process'
import { copyFile, mkdir, rm, writeFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const defaultBuildRoot = join(repoRoot, 'build', 'mcpb')
const defaultExtensionDir = join(defaultBuildRoot, 'bilig-workpaper')
const defaultOutputPath = join(defaultBuildRoot, 'bilig-workpaper.mcpb')
const defaultIconPath = join(repoRoot, 'docs', 'assets', 'bilig-mcp-marketplace-logo.png')
const defaultMcpbPackage = '@anthropic-ai/mcpb@2.1.2'
const headlessPackageName = '@bilig/headless'
const workPaperMcpbPrivacyPolicyUrl = 'https://proompteng.github.io/bilig/workpaper-mcpb-privacy.html'

export type WorkPaperMcpbManifest = {
  readonly manifest_version: '0.4'
  readonly name: 'bilig-workpaper'
  readonly display_name: 'Bilig WorkPaper'
  readonly version: string
  readonly description: string
  readonly long_description: string
  readonly author: {
    readonly name: string
    readonly url: string
  }
  readonly repository: {
    readonly type: 'git'
    readonly url: string
  }
  readonly homepage: string
  readonly documentation: string
  readonly support: string
  readonly icon: 'icon.png'
  readonly server: {
    readonly type: 'node'
    readonly entry_point: 'server/index.js'
    readonly mcp_config: {
      readonly command: 'node'
      readonly args: readonly string[]
      readonly env: Record<string, string>
    }
  }
  readonly tools: readonly {
    readonly name: string
    readonly description: string
  }[]
  readonly keywords: readonly string[]
  readonly license: 'MIT'
  readonly privacy_policies: readonly string[]
  readonly compatibility: {
    readonly platforms: readonly ['darwin', 'win32', 'linux']
    readonly runtimes: {
      readonly node: '>=22.0.0'
    }
  }
}

export type WorkPaperMcpbBuildOptions = {
  readonly extensionDir: string
  readonly outputPath: string
  readonly packageVersion: string
  readonly iconPath: string
  readonly mcpbPackage: string
}

export type WorkPaperMcpbBuildResult = {
  readonly extensionDir: string
  readonly outputPath: string
  readonly packageVersion: string
}

export type WorkPaperMcpbCliOptions = {
  readonly extensionDir?: string
  readonly outputPath?: string
  readonly packageVersion?: string
  readonly iconPath?: string
  readonly mcpbPackage?: string
}

export function buildWorkPaperMcpbManifest(packageVersion: string): WorkPaperMcpbManifest {
  return {
    manifest_version: '0.4',
    name: 'bilig-workpaper',
    display_name: 'Bilig WorkPaper',
    version: packageVersion,
    description: 'Formula-backed WorkPaper tools for workbook readback, input edits, and JSON persistence.',
    long_description: [
      'Bilig WorkPaper gives Claude Desktop a local spreadsheet engine for agent workflows that need more than a screenshot of a grid.',
      'The bundle runs the published @bilig/headless MCP stdio server in file-backed writable mode, exposes tools for reading sheets and ranges, editing cells, validating formulas, exporting WorkPaper JSON, and returning calculated readback.',
      'Use it to evaluate formula-backed workbook automation before wiring the same package into a Node service, queue worker, or coding-agent tool.',
    ].join('\n\n'),
    author: {
      name: 'Proompt Engineering',
      url: 'https://github.com/proompteng',
    },
    repository: {
      type: 'git',
      url: 'https://github.com/proompteng/bilig',
    },
    homepage: 'https://proompteng.github.io/bilig/',
    documentation: 'https://proompteng.github.io/bilig/claude-desktop-mcpb-workpaper.html',
    support: 'https://github.com/proompteng/bilig/issues',
    icon: 'icon.png',
    server: {
      type: 'node',
      entry_point: 'server/index.js',
      mcp_config: {
        command: 'node',
        args: ['${__dirname}/server/index.js'],
        env: {},
      },
    },
    tools: [
      {
        name: 'list_sheets',
        description: 'List WorkPaper sheets and their current used dimensions.',
      },
      {
        name: 'read_range',
        description: 'Read evaluated values and serialized cell contents for a WorkPaper range.',
      },
      {
        name: 'read_cell',
        description: 'Read one cell with evaluated value, display value, formula, and serialized content.',
      },
      {
        name: 'set_cell_contents',
        description: 'Set one cell, recalculate dependents, and persist the updated WorkPaper JSON file.',
      },
      {
        name: 'get_cell_display_value',
        description: 'Return the formatted display value for one cell.',
      },
      {
        name: 'export_workpaper_document',
        description: 'Export the current WorkPaper JSON document.',
      },
      {
        name: 'validate_formula',
        description: 'Validate formula syntax using the WorkPaper formula parser.',
      },
    ],
    keywords: [
      'mcp',
      'mcpb',
      'claude-desktop',
      'spreadsheet',
      'workbook',
      'formula-engine',
      'workpaper',
      'agent-tools',
      'node',
      'typescript',
      'bilig',
    ],
    license: 'MIT',
    privacy_policies: [workPaperMcpbPrivacyPolicyUrl],
    compatibility: {
      platforms: ['darwin', 'win32', 'linux'],
      runtimes: {
        node: '>=22.0.0',
      },
    },
  }
}

export function renderWorkPaperMcpbLauncher(): string {
  return [
    '#!/usr/bin/env node',
    "import { createRequire } from 'node:module';",
    "import { existsSync, writeFileSync } from 'node:fs';",
    "import { dirname, join } from 'node:path';",
    "import { fileURLToPath } from 'node:url';",
    'import {',
    '  exportWorkPaperDocument,',
    '  serializeWorkPaperDocument,',
    "} from '@bilig/headless';",
    '',
    'const requirePackageJson = createRequire(import.meta.url);',
    "const packageManifest = requirePackageJson('../node_modules/@bilig/headless/package.json');",
    "const serverVersion = typeof packageManifest.version === 'string' ? packageManifest.version : '0.0.0';",
    'const serverDir = dirname(fileURLToPath(import.meta.url));',
    "const workpaperPath = join(serverDir, 'workpaper.json');",
    '',
    'async function loadMcpRuntime() {',
    '  try {',
    "    return await import('@bilig/headless/mcp');",
    '  } catch (error) {',
    "    if (typeof error === 'object' && error !== null && 'code' in error && error.code === 'ERR_PACKAGE_PATH_NOT_EXPORTED') {",
    "      return await import('@bilig/headless');",
    '    }',
    '    throw error;',
    '  }',
    '}',
    '',
    'const { buildDemoWorkPaper, createFileBackedWorkPaperMcpToolServerFromFile, runDemoWorkPaperMcpStdioServer } = await loadMcpRuntime();',
    '',
    'if (!existsSync(workpaperPath)) {',
    '  writeFileSync(workpaperPath, serializeWorkPaperDocument(exportWorkPaperDocument(buildDemoWorkPaper(), { includeConfig: true })));',
    '}',
    '',
    'runDemoWorkPaperMcpStdioServer({',
    '  serverVersion,',
    '  server: createFileBackedWorkPaperMcpToolServerFromFile({ workpaperPath, writable: true }),',
    '});',
    '',
  ].join('\n')
}

export function renderWorkPaperMcpbPackageJson(packageVersion: string): string {
  return jsonWithTrailingNewline({
    private: true,
    type: 'module',
    dependencies: {
      [headlessPackageName]: packageVersion,
    },
  })
}

export function renderWorkPaperMcpbReadme(packageVersion: string): string {
  return [
    '# Bilig WorkPaper MCPB',
    '',
    'This bundle runs the published `@bilig/headless` WorkPaper MCP stdio server inside Claude Desktop.',
    '',
    `Bundled package version: \`${packageVersion}\``,
    '',
    'After installation, ask Claude to list the Bilig WorkPaper tools, read `Summary!A1:B5`, set `Inputs!B3` to `0.4`, read `Summary!B3`, and report the recalculated value plus persistence checks.',
    '',
    '## Privacy Policy',
    '',
    'This desktop extension runs locally in Claude Desktop through stdio. It does not send workbook contents, formulas, cell values, or generated `workpaper.json` files to Proompt Engineering. Claude Desktop and Anthropic may process tool calls according to Anthropic policies when you use the extension in Claude.',
    '',
    `Privacy policy: ${workPaperMcpbPrivacyPolicyUrl}`,
    '',
  ].join('\n')
}

export async function resolvePublishedHeadlessVersion(fetchImpl: typeof fetch = fetch): Promise<string> {
  const response = await fetchImpl('https://registry.npmjs.org/%40bilig%2Fheadless/latest')
  if (!response.ok) {
    throw new Error(`Failed to resolve ${headlessPackageName} latest version: HTTP ${String(response.status)}`)
  }

  const parsed: unknown = await response.json()
  if (!isRecord(parsed) || typeof parsed['version'] !== 'string') {
    throw new Error(`Failed to parse ${headlessPackageName} latest version from npm`)
  }

  return parsed['version']
}

export function parseWorkPaperMcpbCliArgs(argv: readonly string[]): WorkPaperMcpbCliOptions {
  const options: Partial<WorkPaperMcpbCliOptions> = {}

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    if (arg === undefined) {
      continue
    }

    if (arg === '--') {
      continue
    }

    if (arg === '--extension-dir') {
      assertUniqueOption(options, 'extensionDir', arg)
      options.extensionDir = requiredValue(argv, index, arg)
      index += 1
      continue
    }
    if (arg === '--output') {
      assertUniqueOption(options, 'outputPath', arg)
      options.outputPath = requiredValue(argv, index, arg)
      index += 1
      continue
    }
    if (arg === '--package-version') {
      assertUniqueOption(options, 'packageVersion', arg)
      options.packageVersion = requiredValue(argv, index, arg)
      index += 1
      continue
    }
    if (arg === '--icon') {
      assertUniqueOption(options, 'iconPath', arg)
      options.iconPath = requiredValue(argv, index, arg)
      index += 1
      continue
    }
    if (arg === '--mcpb-package') {
      assertUniqueOption(options, 'mcpbPackage', arg)
      options.mcpbPackage = requiredValue(argv, index, arg)
      index += 1
      continue
    }
    if (arg === '--help' || arg === '-h') {
      printHelp()
      process.exit(0)
    }

    throw new Error(`Unknown argument: ${arg}`)
  }

  return options
}

export async function buildWorkPaperMcpbBundle(options: WorkPaperMcpbBuildOptions): Promise<WorkPaperMcpbBuildResult> {
  await rm(options.extensionDir, { recursive: true, force: true })
  await mkdir(join(options.extensionDir, 'server'), { recursive: true })
  await mkdir(dirname(options.outputPath), { recursive: true })

  await Promise.all([
    copyFile(options.iconPath, join(options.extensionDir, 'icon.png')),
    writeFile(join(options.extensionDir, 'manifest.json'), jsonWithTrailingNewline(buildWorkPaperMcpbManifest(options.packageVersion))),
    writeFile(join(options.extensionDir, 'package.json'), renderWorkPaperMcpbPackageJson(options.packageVersion)),
    writeFile(join(options.extensionDir, 'README.md'), renderWorkPaperMcpbReadme(options.packageVersion)),
    writeFile(join(options.extensionDir, 'server', 'index.js'), renderWorkPaperMcpbLauncher()),
  ])

  run('npm', ['install', '--prefix', options.extensionDir, '--omit=dev', '--ignore-scripts', '--no-audit', '--no-fund'], repoRoot)
  run('npx', ['-y', options.mcpbPackage, 'pack', options.extensionDir, options.outputPath], repoRoot)

  return {
    extensionDir: options.extensionDir,
    outputPath: options.outputPath,
    packageVersion: options.packageVersion,
  }
}

export async function resolveBuildOptions(cliOptions: WorkPaperMcpbCliOptions): Promise<WorkPaperMcpbBuildOptions> {
  return {
    extensionDir: resolve(cliOptions.extensionDir ?? defaultExtensionDir),
    outputPath: resolve(cliOptions.outputPath ?? defaultOutputPath),
    packageVersion: cliOptions.packageVersion ?? (await resolvePublishedHeadlessVersion()),
    iconPath: resolve(cliOptions.iconPath ?? defaultIconPath),
    mcpbPackage: cliOptions.mcpbPackage ?? defaultMcpbPackage,
  }
}

function requiredValue(argv: readonly string[], index: number, flag: string): string {
  const value = argv[index + 1]
  if (value === undefined || value.trim().length === 0 || value.startsWith('--')) {
    throw new Error(`${flag} requires a value`)
  }
  return value
}

function assertUniqueOption(options: Partial<WorkPaperMcpbCliOptions>, key: keyof WorkPaperMcpbCliOptions, flag: string): void {
  if (Object.hasOwn(options, key)) {
    throw new Error(`Duplicate argument: ${flag}`)
  }
}

function jsonWithTrailingNewline(value: unknown): string {
  return `${JSON.stringify(value, null, 2)}\n`
}

function run(command: string, args: readonly string[], cwd: string): void {
  execFileSync(command, [...args], {
    cwd,
    stdio: 'inherit',
  })
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function printHelp(): void {
  console.log(`Usage: pnpm mcpb:workpaper:build [options]

Options:
  --package-version <version>  Bundle a specific @bilig/headless version. Defaults to npm latest.
  --extension-dir <path>       Staging directory. Defaults to build/mcpb/bilig-workpaper.
  --output <path>              Output file. Defaults to build/mcpb/bilig-workpaper.mcpb.
  --icon <path>                Icon file. Defaults to docs/assets/bilig-mcp-marketplace-logo.png.
  --mcpb-package <specifier>   MCPB CLI package. Defaults to ${defaultMcpbPackage}.
`)
}

async function main(): Promise<void> {
  const cliOptions = parseWorkPaperMcpbCliArgs(process.argv.slice(2))
  const options = await resolveBuildOptions(cliOptions)
  const result = await buildWorkPaperMcpbBundle(options)
  console.log(
    jsonWithTrailingNewline({
      ok: true,
      packageVersion: result.packageVersion,
      extensionDir: result.extensionDir,
      outputPath: result.outputPath,
    }),
  )
}

if (process.argv[1] !== undefined && pathToFileURL(process.argv[1]).href === import.meta.url) {
  await main()
}
