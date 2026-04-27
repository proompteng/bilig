import process from 'node:process'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it } from 'vitest'
import { CodexAppServerClient } from '../codex-app-server-client.js'

const fixturePath = fileURLToPath(new URL('./fixtures/fake-codex-app-server.mjs', import.meta.url))

describe('Codex app-server client', () => {
  let client: CodexAppServerClient | null = null

  afterEach(async () => {
    await client?.close()
    client = null
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
      args: [fixturePath],
      env: {
        BILIG_TEST_ECHO_THREAD_START: '1',
      },
      handleDynamicToolCall: async () => ({
        success: true,
        contentItems: [],
      }),
    })

    const threadConfig = {
      approval_policy: 'never',
      sandbox_mode: 'danger-full-access',
      network_access: true,
      web_search: 'live',
      tools: {
        view_image: true,
      },
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

  it('strips inherited OTEL exporter env before spawning the app-server', async () => {
    client = new CodexAppServerClient({
      command: process.execPath,
      args: [fixturePath],
      env: {
        BILIG_TEST_EXPECT_OTEL_STRIPPED: '1',
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
      args: [fixturePath],
      env: {
        BILIG_TEST_EMIT_REASONING_DELTA: '1',
      },
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
})
