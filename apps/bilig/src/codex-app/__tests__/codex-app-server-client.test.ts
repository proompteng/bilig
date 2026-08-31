import process from 'node:process'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { CodexAppServerClient, createCodexChildEnvironment } from '../codex-app-server-client.js'

const fixturePath = fileURLToPath(new URL('./fixtures/fake-codex-app-server.mjs', import.meta.url))

function threadStartInput() {
  return {
    model: 'gpt-5.4',
    approvalPolicy: 'never' as const,
    sandbox: 'read-only' as const,
    baseInstructions: 'base',
    developerInstructions: 'developer',
    dynamicTools: [],
  }
}

describe('Codex app-server client', () => {
  let client: CodexAppServerClient | null = null

  afterEach(async () => {
    await client?.close()
    client = null
  })

  it('passes only the explicit app-server environment allowlist to child processes', () => {
    const childEnv = createCodexChildEnvironment({
      PATH: '/usr/bin',
      OPENAI_API_KEY: 'test-key',
      HOME: '/Users/alex',
      CODEX_HOME: '/Users/alex/.codex',
      XDG_CONFIG_HOME: '/Users/alex/.config',
      DATABASE_URL: 'postgres://db.example.test/app',
      BILIG_SESSION_SECRET: 'session-secret',
      OTEL_EXPORTER_OTLP_ENDPOINT: 'http://127.0.0.1:4318',
    })

    expect(childEnv).toEqual({
      PATH: '/usr/bin',
      OPENAI_API_KEY: 'test-key',
      OTEL_SDK_DISABLED: 'true',
    })
  })

  it('maps an explicit service home without inheriting user Codex configuration', () => {
    const codexHome = path.join('/var/lib/bilig', 'codex')
    const childEnv = createCodexChildEnvironment(
      {
        PATH: '/usr/bin',
        HOME: '/Users/alex',
        CODEX_HOME: '/Users/alex/.codex',
        XDG_CONFIG_HOME: '/Users/alex/.config',
      },
      { codexHome },
    )

    expect(childEnv).toMatchObject({
      PATH: '/usr/bin',
      HOME: codexHome,
      CODEX_HOME: codexHome,
      XDG_CONFIG_HOME: path.join(codexHome, 'xdg-config'),
      XDG_DATA_HOME: path.join(codexHome, 'xdg-data'),
      XDG_CACHE_HOME: path.join(codexHome, 'xdg-cache'),
    })
    expect(childEnv['CODEX_HOME']).not.toBe('/Users/alex/.codex')
  })

  it('declares experimentalApi during initialize before starting a dynamic-tools thread', async () => {
    client = new CodexAppServerClient({
      command: process.execPath,
      args: [fixturePath],
      handleDynamicToolCall: async () => ({
        success: true,
        contentItems: [],
      }),
    })

    const initialized = await client.ensureReady()
    expect(initialized).toEqual({
      userAgent: 'fake-codex-app-server',
      codexHome: '/tmp/fake-codex-home',
      platformFamily: 'unix',
      platformOs: 'macos',
    })

    const thread = await client.threadStart({
      model: 'gpt-5.4',
      approvalPolicy: 'never',
      sandbox: 'read-only',
      baseInstructions: 'base',
      developerInstructions: 'developer',
      dynamicTools: [
        {
          name: 'test_tool',
          description: 'Test dynamic tool',
          inputSchema: {
            type: 'object',
          },
        },
      ],
    })

    expect(thread.id).toBe('thr-fixture')
    expect(thread.preview).toBe('experimentalApi:true')
  })

  it('passes explicit thread permission config to the app-server', async () => {
    client = new CodexAppServerClient({
      command: process.execPath,
      args: [fixturePath, '--echo-thread-start'],
      handleDynamicToolCall: async () => ({
        success: true,
        contentItems: [],
      }),
    })

    const threadConfig = {
      approval_policy: 'never',
      sandbox_mode: 'danger-full-access',
      sandbox_workspace_write: {
        network_access: true,
      },
      web_search: 'live',
    } as const

    const thread = await client.threadStart({
      model: 'gpt-5.4',
      approvalPolicy: 'never',
      sandbox: 'danger-full-access',
      config: threadConfig,
      baseInstructions: 'base',
      developerInstructions: 'developer',
      dynamicTools: [],
    })

    expect(JSON.parse(thread.preview)).toEqual({
      experimentalApi: true,
      approvalPolicy: 'never',
      sandbox: 'danger-full-access',
      config: threadConfig,
    })
  })

  it('reasserts explicit thread permission config when resuming a thread', async () => {
    client = new CodexAppServerClient({
      command: process.execPath,
      args: [fixturePath, '--echo-thread-resume'],
      handleDynamicToolCall: async () => ({
        success: true,
        contentItems: [],
      }),
    })

    const threadConfig = {
      approval_policy: 'on-request',
      sandbox_mode: 'read-only',
      sandbox_workspace_write: {
        network_access: false,
      },
      web_search: 'disabled',
      mcp_servers: {},
    } as const
    const thread = await client.threadResume({
      threadId: 'thr-legacy',
      approvalPolicy: 'on-request',
      sandbox: 'read-only',
      cwd: '/tmp',
      runtimeWorkspaceRoots: [],
      config: threadConfig,
      baseInstructions: 'base',
      developerInstructions: 'developer',
    })

    expect(JSON.parse(thread.preview)).toEqual({
      threadId: 'thr-legacy',
      approvalPolicy: 'on-request',
      sandbox: 'read-only',
      cwd: '/tmp',
      runtimeWorkspaceRoots: [],
      config: threadConfig,
    })
  })

  it('strips inherited OTEL exporter env before spawning the app-server', async () => {
    client = new CodexAppServerClient({
      command: process.execPath,
      args: [fixturePath, '--expect-otel-stripped'],
      env: {
        OTEL_EXPORTER_OTLP_ENDPOINT: 'http://127.0.0.1:4318',
        OTEL_EXPORTER_OTLP_LOGS_ENDPOINT: 'http://127.0.0.1:4318/v1/logs',
      },
      handleDynamicToolCall: async () => ({
        success: true,
        contentItems: [],
      }),
    })

    await client.ensureReady()
    const thread = await client.threadStart({
      model: 'gpt-5.4',
      approvalPolicy: 'never',
      sandbox: 'read-only',
      baseInstructions: 'base',
      developerInstructions: 'developer',
      dynamicTools: [
        {
          name: 'test_tool',
          description: 'Test dynamic tool',
          inputSchema: {
            type: 'object',
          },
        },
      ],
    })

    expect(thread.id).toBe('thr-fixture')
  })

  it('parses reasoning delta notifications from the app-server stream', async () => {
    const notifications: unknown[] = []
    client = new CodexAppServerClient({
      command: process.execPath,
      args: [fixturePath, '--emit-reasoning-delta'],
      handleDynamicToolCall: async () => ({
        success: true,
        contentItems: [],
      }),
    })

    client.subscribe((notification) => {
      notifications.push(notification)
    })

    await client.ensureReady()
    await client.threadStart({
      model: 'gpt-5.4',
      approvalPolicy: 'never',
      sandbox: 'read-only',
      baseInstructions: 'base',
      developerInstructions: 'developer',
      dynamicTools: [],
    })
    const turn = await client.turnStart({
      threadId: 'thr-fixture',
      prompt: 'Check staged changes',
    })

    expect(turn.id).toBe('turn-fixture')
    expect(notifications).toContainEqual({
      method: 'item/reasoning/textDelta',
      params: {
        threadId: 'thr-fixture',
        turnId: 'turn-fixture',
        itemId: 'reasoning-fixture',
        delta: 'Examining staged changes',
      },
    })
  })

  it('reports notification listener failures without rejecting the stream response', async () => {
    const onLog = vi.fn()
    client = new CodexAppServerClient({
      command: process.execPath,
      args: [fixturePath, '--emit-reasoning-delta'],
      onLog,
      handleDynamicToolCall: async () => ({
        success: true,
        contentItems: [],
      }),
    })

    client.subscribe(() => {
      throw new Error('notification consumer exploded')
    })

    const thread = await client.threadStart(threadStartInput())
    await expect(client.turnStart({ threadId: thread.id, prompt: 'continue' })).resolves.toMatchObject({
      id: 'turn-fixture',
    })
    await vi.waitFor(() => {
      expect(onLog).toHaveBeenCalledWith(expect.stringContaining('notification consumer exploded'))
    })
  })

  it('rejects pending requests and restarts after the app-server exits', async () => {
    client = new CodexAppServerClient({
      command: process.execPath,
      args: [fixturePath, '--exit-during-turn-start'],
      handleDynamicToolCall: async () => ({
        success: true,
        contentItems: [],
      }),
    })

    const thread = await client.threadStart(threadStartInput())

    await expect(
      client.turnStart({
        threadId: thread.id,
        prompt: 'this request will outlive the process',
      }),
    ).rejects.toThrow('Codex app-server exited unexpectedly')

    const resumed = await client.threadResume({
      threadId: thread.id,
      baseInstructions: 'base',
      developerInstructions: 'developer',
    })
    expect(resumed.id).toBe(thread.id)
  })
})
