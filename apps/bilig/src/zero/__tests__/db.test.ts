import { beforeEach, describe, expect, it, vi } from 'vitest'

const { Pool, poolCtor, poolInstances } = vi.hoisted(() => {
  const instances: MockPool[] = []
  const ctorSpy = vi.fn()

  class MockPool {
    readonly options: { readonly connectionString: string }
    private readonly listeners = new Map<string, ((error: unknown) => void)[]>()

    constructor(options: { readonly connectionString: string }) {
      ctorSpy(options)
      this.options = options
      instances.push(this)
    }

    on(eventName: string, listener: (error: unknown) => void) {
      const nextListeners = this.listeners.get(eventName) ?? []
      nextListeners.push(listener)
      this.listeners.set(eventName, nextListeners)
      return this
    }

    emit(eventName: string, payload: unknown) {
      for (const listener of this.listeners.get(eventName) ?? []) {
        listener(payload)
      }
      return (this.listeners.get(eventName)?.length ?? 0) > 0
    }

    listenerCount(eventName: string) {
      return this.listeners.get(eventName)?.length ?? 0
    }
  }

  return {
    poolCtor: ctorSpy,
    poolInstances: instances,
    Pool: MockPool,
  }
})

vi.mock('pg', () => ({
  Pool,
}))

import { createZeroPool } from '../db.js'

describe('createZeroPool', () => {
  beforeEach(() => {
    poolCtor.mockClear()
    poolInstances.length = 0
  })

  it('attaches a pool error listener so idle client disconnects do not crash the process', () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {})
    const pool = createZeroPool('postgres://example.test/bilig')
    expect(poolCtor).toHaveBeenCalledWith({
      connectionString: 'postgres://example.test/bilig',
    })

    expect(pool.listenerCount('error')).toBeGreaterThan(0)
    pool.emit('error', new Error('Connection terminated unexpectedly'))

    expect(consoleError).toHaveBeenCalledWith(
      'Zero Postgres pool error',
      expect.objectContaining({
        message: 'Connection terminated unexpectedly',
      }),
    )

    consoleError.mockRestore()
  })
})
