import type { AgentIdeRuleInput } from './agent-discovery-ide-rules.ts'

export function buildVscodeMcpConfig(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, workpaperPackageSpec } = input
  return `${JSON.stringify(
    {
      servers: {
        biligWorkpaperDemo: {
          type: 'http',
          url: remoteMcpEndpoint,
        },
        biligWorkpaperFile: {
          type: 'stdio',
          command: 'npm',
          args: [
            'exec',
            '--package',
            workpaperPackageSpec,
            '--',
            'bilig-workpaper-mcp',
            '--workpaper',
            '${workspaceFolder}/.bilig/pricing.workpaper.json',
            '--init-demo-workpaper',
            '--writable',
          ],
        },
      },
    },
    null,
    2,
  )}\n`
}

export function buildZedSettingsConfig(input: AgentIdeRuleInput): string {
  const { workpaperPackageSpec } = input
  return `${JSON.stringify(
    {
      context_servers: {
        'bilig-workpaper': {
          command: {
            path: 'npm',
            args: [
              'exec',
              '--yes',
              '--package',
              workpaperPackageSpec,
              '--',
              'bilig-workpaper-mcp',
              '--workpaper',
              './.bilig/pricing.workpaper.json',
              '--init-demo-workpaper',
              '--writable',
            ],
            env: {},
          },
          settings: {},
        },
      },
      agent: {
        default_profile: 'write',
      },
    },
    null,
    2,
  )}\n`
}

export function buildTraeMcpConfig(input: AgentIdeRuleInput): string {
  const { workpaperPackageSpec } = input
  return `${JSON.stringify(
    {
      mcpServers: {
        'bilig-workpaper': {
          command: 'npm',
          args: [
            'exec',
            '--yes',
            '--package',
            workpaperPackageSpec,
            '--',
            'bilig-workpaper-mcp',
            '--workpaper',
            './.bilig/pricing.workpaper.json',
            '--init-demo-workpaper',
            '--writable',
          ],
          env: {},
        },
      },
    },
    null,
    2,
  )}\n`
}

export function buildOpenCodeMcpConfig(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, workpaperPackageSpec } = input
  return `{
  "$schema": "https://opencode.ai/config.json",
  "instructions": ["AGENTS.md"],
  "mcp": {
    "bilig-workpaper": {
      "type": "local",
      "command": [
        "npm",
        "exec",
        "--yes",
        "--package",
        ${JSON.stringify(workpaperPackageSpec)},
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./.bilig/pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable",
      ],
      "enabled": true,
    },
    "bilig-workpaper-demo": {
      "type": "remote",
      "url": ${JSON.stringify(remoteMcpEndpoint)},
      "enabled": false,
    },
  },
}
`
}

export function buildFileBackedMcpServerConfig(input: {
  readonly serverKey: string
  readonly workpaperPath: string
  readonly workpaperPackageSpec: string
}): string {
  const { serverKey, workpaperPath, workpaperPackageSpec } = input
  return `${JSON.stringify(
    {
      mcpServers: {
        [serverKey]: {
          type: 'stdio',
          command: 'npm',
          args: [
            'exec',
            '--yes',
            '--package',
            workpaperPackageSpec,
            '--',
            'bilig-workpaper-mcp',
            '--workpaper',
            workpaperPath,
            '--init-demo-workpaper',
            '--writable',
          ],
          env: {},
        },
      },
    },
    null,
    2,
  )}\n`
}

export function buildClaudeCodeMcpConfig(input: AgentIdeRuleInput): string {
  return buildFileBackedMcpServerConfig({
    serverKey: 'bilig-workpaper',
    workpaperPackageSpec: input.workpaperPackageSpec,
    workpaperPath: './.bilig/pricing.workpaper.json',
  })
}

export function buildCursorMcpConfig(input: AgentIdeRuleInput): string {
  return buildFileBackedMcpServerConfig({
    serverKey: 'biligWorkpaperFile',
    workpaperPackageSpec: input.workpaperPackageSpec,
    workpaperPath: './.bilig/pricing.workpaper.json',
  })
}

export function buildJunieMcpConfig(input: AgentIdeRuleInput): string {
  return buildFileBackedMcpServerConfig({
    serverKey: 'biligWorkpaperFile',
    workpaperPackageSpec: input.workpaperPackageSpec,
    workpaperPath: './.bilig/pricing.workpaper.json',
  })
}

export function buildRooMcpConfig(input: AgentIdeRuleInput): string {
  return buildFileBackedMcpServerConfig({
    serverKey: 'bilig-workpaper',
    workpaperPackageSpec: input.workpaperPackageSpec,
    workpaperPath: './.bilig/pricing.workpaper.json',
  })
}

export function buildReusableMcpConfig(input: AgentIdeRuleInput): string {
  return buildFileBackedMcpServerConfig({
    serverKey: 'bilig-workpaper',
    workpaperPackageSpec: input.workpaperPackageSpec,
    workpaperPath: './.bilig/pricing.workpaper.json',
  })
}
