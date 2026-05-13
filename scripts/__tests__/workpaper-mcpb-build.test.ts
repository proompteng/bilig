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
    expect(manifest.tools.map((tool) => tool.name)).toEqual(['read_workpaper_summary', 'set_workpaper_input_cell'])
    expect(manifest.keywords).toContain('mcpb')
    expect(manifest.compatibility.runtimes.node).toBe('>=24.0.0')
  })

  it('renders a module launcher and package manifest that bundle the exact published package version', () => {
    expect(renderWorkPaperMcpbLauncher()).toBe(
      [
        '#!/usr/bin/env node',
        "import { createRequire } from 'node:module';",
        "import { runDemoWorkPaperMcpStdioServer } from '@bilig/headless';",
        '',
        'const requirePackageJson = createRequire(import.meta.url);',
        "const packageManifest = requirePackageJson('../node_modules/@bilig/headless/package.json');",
        "const serverVersion = typeof packageManifest.version === 'string' ? packageManifest.version : '0.0.0';",
        '',
        'runDemoWorkPaperMcpStdioServer({ serverVersion });',
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
})
