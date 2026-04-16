import { describe, expect, it, vi } from 'vitest'
import type { CodexInitializeResponse, CodexServerNotification, CodexTurn } from '@bilig/agent-api'
import type { CodexAppServerClientOptions, CodexAppServerTransport } from './codex-app-server-client.js'
import { CodexAppServerClientPool, CodexAppServerPoolBackpressureError } from './codex-app-server-pool.js'

class FakePoolTransport implements CodexAppServerTransport {
  private readonly listeners = new Set<(notification: CodexServerNotification) => void>()
  readonly startedThreads: string[] = []
  readonly resumedThreads: string[] = []
  readonly turnStartThreadIds: string[] = []
  closeCount = 0
  nextTurn: Promise<CodexTurn> | null = null

  constructor(readonly label: string) {}

  async ensureReady(): Promise<CodexInitializeResponse> {
    return {
      userAgent: `fake-${this.label}`,
      codexHome: `/tmp/${this.label}`,
      platformFamily: 'unix',
      platformOs: 'macos',
    }
  }

  subscribe(listener: (notification: CodexServerNotification) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  async threadStart(): Promise<{ id: string; preview: string; turns: [] }> {
    const threadId = `${this.label}-thr-${String(this.startedThreads.length + 1)}`
    this.startedThreads.push(threadId)
    return {
      id: threadId,
      preview: '',
      turns: [],
    }
  }

  async threadResume(input: { threadId: string }): Promise<{ id: string; preview: string; turns: [] }> {
    this.resumedThreads.push(input.threadId)
    return {
      id: input.threadId,
      preview: '',
      turns: [],
    }
  }

  async turnStart(input: { threadId: string; prompt: string }): Promise<CodexTurn> {
    this.turnStartThreadIds.push(input.threadId)
    if (this.nextTurn) {
      return await this.nextTurn
    }
    return {
      id: `${this.label}-turn-${String(this.turnStartThreadIds.length)}`,
      status: 'inProgress',
      items: [],
      error: null,
    }
  }

  async turnInterrupt(): Promise<void> {}

  async close(): Promise<void> {
    this.closeCount += 1
  }
}

function createPool(input?: {
  maxClients?: number
  maxConcurrentTurnsPerClient?: number
  maxQueuedTurnsPerClient?: number
  transports?: FakePoolTransport[]
}) {
  const transports = input?.transports ?? []
  const capturedOptions: CodexAppServerClientOptions[] = []
  const pool = new CodexAppServerClientPool({
    codexClientFactory: (options) => {
      capturedOptions.push(options)
      const nextTransport = transports.shift() ?? new FakePoolTransport(`transport-${String(capturedOptions.length)}`)
      return nextTransport
    },
    clientOptions: {
      command: 'codex',
      args: ['app-server'],
      cwd: process.cwd(),
      env: process.env,
      onLog: vi.fn(),
      handleDynamicToolCall: vi.fn(async () => ({
        success: true,
        contentItems: [],
      })),
    },
    ...(input?.maxClients !== undefined ? { maxClients: input.maxClients } : {}),
    ...(input?.maxConcurrentTurnsPerClient !== undefined ? { maxConcurrentTurnsPerClient: input.maxConcurrentTurnsPerClient } : {}),
    ...(input?.maxQueuedTurnsPerClient !== undefined ? { maxQueuedTurnsPerClient: input.maxQueuedTurnsPerClient } : {}),
  })
  return { pool, capturedOptions }
}

function resolvePendingTurn(resolveTurn: ((value: CodexTurn) => void) | null, value: CodexTurn): void {
  if (!resolveTurn) {
    throw new Error('Expected pending turn resolver to be installed')
  }
  resolveTurn(value)
}

describe('codex-app-server-pool', () => {
  it('starts threads across a bounded pool of clients', async () => {
    const fakeA = new FakePoolTransport('A')
    const fakeB = new FakePoolTransport('B')
    const { pool, capturedOptions } = createPool({
      maxClients: 2,
      transports: [fakeA, fakeB],
    })

    try {
      const threadA = await pool.threadStart({
        model: 'gpt-5.4',
        approvalPolicy: 'never',
        sandbox: 'read-only',
        baseInstructions: 'base',
        developerInstructions: 'dev',
        dynamicTools: [],
      })
      const threadB = await pool.threadStart({
        model: 'gpt-5.4',
        approvalPolicy: 'never',
        sandbox: 'read-only',
        baseInstructions: 'base',
        developerInstructions: 'dev',
        dynamicTools: [],
      })

      expect(capturedOptions).toHaveLength(2)
      expect(threadA.id).toBe('A-thr-1')
      expect(threadB.id).toBe('B-thr-1')
      await pool.turnStart({ threadId: threadA.id, prompt: 'hello' })
      await pool.turnStart({ threadId: threadB.id, prompt: 'world' })
      expect(fakeA.turnStartThreadIds).toEqual(['A-thr-1'])
      expect(fakeB.turnStartThreadIds).toEqual(['B-thr-1'])
    } finally {
      await pool.close()
    }
  })

  it('queues turns per client when concurrency is saturated', async () => {
    let resolveFirstTurn: ((value: CodexTurn) => void) | null = null
    const firstTurnPromise = new Promise<CodexTurn>((resolve) => {
      resolveFirstTurn = resolve
    })
    const fake = new FakePoolTransport('A')
    fake.nextTurn = firstTurnPromise
    const { pool } = createPool({
      maxClients: 1,
      maxConcurrentTurnsPerClient: 1,
      maxQueuedTurnsPerClient: 1,
      transports: [fake],
    })

    try {
      const thread = await pool.threadStart({
        model: 'gpt-5.4',
        approvalPolicy: 'never',
        sandbox: 'read-only',
        baseInstructions: 'base',
        developerInstructions: 'dev',
        dynamicTools: [],
      })
      const firstTurn = pool.turnStart({ threadId: thread.id, prompt: 'first' })
      const secondTurn = pool.turnStart({ threadId: thread.id, prompt: 'second' })

      await Promise.resolve()
      expect(fake.turnStartThreadIds).toEqual([thread.id])

      resolvePendingTurn(resolveFirstTurn, {
        id: 'turn-1',
        status: 'inProgress',
        items: [],
        error: null,
      })
      fake.nextTurn = null
      await firstTurn
      await secondTurn

      expect(fake.turnStartThreadIds).toEqual([thread.id, thread.id])
    } finally {
      await pool.close()
    }
  })

  it('applies backpressure when the per-client turn queue is full', async () => {
    let resolveFirstTurn: ((value: CodexTurn) => void) | null = null
    const firstTurnPromise = new Promise<CodexTurn>((resolve) => {
      resolveFirstTurn = resolve
    })
    const fake = new FakePoolTransport('A')
    fake.nextTurn = firstTurnPromise
    const { pool } = createPool({
      maxClients: 1,
      maxConcurrentTurnsPerClient: 1,
      maxQueuedTurnsPerClient: 1,
      transports: [fake],
    })

    try {
      const thread = await pool.threadStart({
        model: 'gpt-5.4',
        approvalPolicy: 'never',
        sandbox: 'read-only',
        baseInstructions: 'base',
        developerInstructions: 'dev',
        dynamicTools: [],
      })
      const firstTurn = pool.turnStart({ threadId: thread.id, prompt: 'first' })
      const secondTurn = pool.turnStart({ threadId: thread.id, prompt: 'second' })
      await expect(pool.turnStart({ threadId: thread.id, prompt: 'third' })).rejects.toBeInstanceOf(CodexAppServerPoolBackpressureError)
      resolvePendingTurn(resolveFirstTurn, {
        id: 'turn-1',
        status: 'inProgress',
        items: [],
        error: null,
      })
      fake.nextTurn = null
      await firstTurn
      await secondTurn
    } finally {
      await pool.close()
    }
  })

  it('releases and closes idle clients when their threads are evicted', async () => {
    const fake = new FakePoolTransport('A')
    const { pool } = createPool({
      maxClients: 1,
      transports: [fake],
    })

    try {
      const thread = await pool.threadStart({
        model: 'gpt-5.4',
        approvalPolicy: 'never',
        sandbox: 'read-only',
        baseInstructions: 'base',
        developerInstructions: 'dev',
        dynamicTools: [],
      })
      pool.releaseThread(thread.id)
      await vi.waitFor(() => {
        expect(fake.closeCount).toBe(1)
      })
    } finally {
      await pool.close()
    }
  })

  it('reports pool stats for observability', async () => {
    let resolveFirstTurn: ((value: CodexTurn) => void) | null = null
    const firstTurnPromise = new Promise<CodexTurn>((resolve) => {
      resolveFirstTurn = resolve
    })
    const fake = new FakePoolTransport('A')
    fake.nextTurn = firstTurnPromise
    const { pool } = createPool({
      maxClients: 2,
      maxConcurrentTurnsPerClient: 1,
      maxQueuedTurnsPerClient: 2,
      transports: [fake],
    })

    try {
      const thread = await pool.threadStart({
        model: 'gpt-5.4',
        approvalPolicy: 'never',
        sandbox: 'read-only',
        baseInstructions: 'base',
        developerInstructions: 'dev',
        dynamicTools: [],
      })
      const firstTurn = pool.turnStart({ threadId: thread.id, prompt: 'first' })
      const secondTurn = pool.turnStart({ threadId: thread.id, prompt: 'second' })

      await Promise.resolve()
      expect(pool.getStats()).toEqual({
        slotCount: 1,
        boundThreadCount: 1,
        activeTurnCount: 1,
        queuedTurnCount: 1,
        maxClients: 2,
        maxConcurrentTurnsPerClient: 1,
        maxQueuedTurnsPerClient: 2,
      })

      resolvePendingTurn(resolveFirstTurn, {
        id: 'turn-1',
        status: 'inProgress',
        items: [],
        error: null,
      })
      fake.nextTurn = null
      await firstTurn
      await secondTurn
    } finally {
      await pool.close()
    }
  })
})
