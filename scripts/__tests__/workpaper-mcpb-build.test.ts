import { describe, expect, it } from 'vitest'

import {
  buildWorkPaperMcpbManifest,
  parseWorkPaperMcpbCliArgs,
  renderWorkPaperMcpbLauncher,
  renderWorkPaperMcpbPackageJson,
  renderWorkPaperMcpbReadme,
} from '../build-workpaper-mcpb.ts'

describe('WorkPaper MCPB builder', () => {
  it('renders a Claude Desktop MCPB manifest for the published WorkPaper stdio server', () => {
    const manifest = buildWorkPaperMcpbManifest('0.13.27')

    expect(manifest).toEqual(
      expect.objectContaining({
        manifest_version: '0.4',
        name: 'bilig-workpaper',
        display_name: 'Bilig WorkPaper',
        version: '0.13.27',
        homepage: 'https://proompteng.github.io/bilig/',
        documentation: 'https://proompteng.github.io/bilig/claude-desktop-mcpb-workpaper.html',
        icon: 'icon.png',
        license: 'MIT',
        privacy_policies: ['https://proompteng.github.io/bilig/workpaper-mcpb-privacy.html'],
      }),
    )
    expect(manifest.server).toEqual({
      type: 'node',
      entry_point: 'server/index.js',
      mcp_config: {
        command: 'node',
        args: ['${__dirname}/server/index.js'],
        env: {},
      },
    })
    expect(manifest.tools.map((tool) => tool.name)).toEqual([
      'list_sheets',
      'read_range',
      'read_cell',
      'set_cell_contents',
      'get_cell_display_value',
      'export_workpaper_document',
      'validate_formula',
    ])
    expect(manifest.keywords).toContain('mcpb')
    expect(manifest.compatibility.runtimes.node).toBe('>=22.0.0')
  })

  it('renders a module launcher and package manifest that bundle the exact published package version', () => {
    expect(renderWorkPaperMcpbLauncher()).toBe(
      [
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
      ].join('\n'),
    )
    expect(JSON.parse(renderWorkPaperMcpbPackageJson('0.13.27'))).toEqual({
      private: true,
      type: 'module',
      dependencies: {
        '@bilig/headless': '0.13.27',
      },
    })
    expect(renderWorkPaperMcpbReadme('0.13.27')).toContain('Bundled package version: `0.13.27`')
    expect(renderWorkPaperMcpbReadme('0.13.27')).toContain('## Privacy Policy')
    expect(renderWorkPaperMcpbReadme('0.13.27')).toContain('https://proompteng.github.io/bilig/workpaper-mcpb-privacy.html')
  })

  it('parses explicit CLI paths and versions', () => {
    expect(
      parseWorkPaperMcpbCliArgs([
        '--',
        '--package-version',
        '0.13.27',
        '--extension-dir',
        'tmp/bundle',
        '--output',
        'tmp/bilig-workpaper.mcpb',
        '--icon',
        'docs/assets/bilig-mcp-marketplace-logo.png',
        '--mcpb-package',
        '@anthropic-ai/mcpb@2.1.2',
      ]),
    ).toEqual({
      packageVersion: '0.13.27',
      extensionDir: 'tmp/bundle',
      outputPath: 'tmp/bilig-workpaper.mcpb',
      iconPath: 'docs/assets/bilig-mcp-marketplace-logo.png',
      mcpbPackage: '@anthropic-ai/mcpb@2.1.2',
    })
  })

  it('rejects blank MCPB CLI values before building the bundle', () => {
    expect(() => parseWorkPaperMcpbCliArgs(['--output', '   '])).toThrow('--output requires a value')
  })

  it('rejects duplicate MCPB CLI values instead of silently overriding build targets', () => {
    expect(() => parseWorkPaperMcpbCliArgs(['--package-version', '0.13.27', '--package-version', '0.13.28'])).toThrow(
      'Duplicate argument: --package-version',
    )
  })
})
