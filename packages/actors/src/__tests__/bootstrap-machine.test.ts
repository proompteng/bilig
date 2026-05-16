import { describe, expect, it } from 'vitest'
import { createActor } from 'xstate'
import { createBootstrapMachine } from '../index.js'

describe('@bilig/actors bootstrap machine', () => {
  it('reaches ready after loading config and session', async () => {
    const machine = createBootstrapMachine<{ defaultDocumentId: string }, { authToken: string }>()
    const actor = createActor(machine, {
      input: {
        loadConfig: async () => ({ defaultDocumentId: 'bilig-demo' }),
        loadSession: async () => ({ authToken: 'token-123' }),
      },
    })

    const done = new Promise<void>((resolve, reject) => {
      const subscription = actor.subscribe((snapshot) => {
        if (snapshot.matches('ready')) {
          subscription.unsubscribe()
          resolve()
          return
        }
        if (snapshot.matches('failed')) {
          subscription.unsubscribe()
          reject(new Error(snapshot.context.error ?? 'bootstrap failed'))
        }
      })
    })

    actor.start()
    await done

    expect(actor.getSnapshot().context.config).toEqual({
      defaultDocumentId: 'bilig-demo',
    })
    expect(actor.getSnapshot().context.session).toEqual({
      authToken: 'token-123',
    })
  })

  it('supports retry after a failed config load', async () => {
    let attempts = 0
    const machine = createBootstrapMachine<{ defaultDocumentId: string }, { authToken: string }>()
    const actor = createActor(machine, {
      input: {
        loadConfig: async () => {
          attempts += 1
          if (attempts === 1) {
            throw new Error('temporary failure')
          }
          return { defaultDocumentId: 'bilig-demo' }
        },
        loadSession: async () => ({ authToken: 'token-123' }),
      },
    })

    actor.start()
    await new Promise<void>((resolve) => {
      const subscription = actor.subscribe((snapshot) => {
        if (snapshot.matches('failed')) {
          subscription.unsubscribe()
          resolve()
        }
      })
    })

    actor.send({ type: 'retry' })
    await new Promise<void>((resolve, reject) => {
      const subscription = actor.subscribe((snapshot) => {
        if (snapshot.matches('ready')) {
          subscription.unsubscribe()
          resolve()
          return
        }
        if (snapshot.matches('failed') && attempts > 1) {
          subscription.unsubscribe()
          reject(new Error(snapshot.context.error ?? 'retry failed'))
        }
      })
    })

    expect(attempts).toBe(2)
  })

  it('automatically retries transient config failures when configured', async () => {
    let attempts = 0
    let sawRetrying = false
    const machine = createBootstrapMachine<{ defaultDocumentId: string }, { authToken: string }>()
    const actor = createActor(machine, {
      input: {
        autoRetryDelayMs: 1,
        maxAutoRetryAttempts: 1,
        loadConfig: async () => {
          attempts += 1
          if (attempts === 1) {
            throw new Error('temporary failure')
          }
          return { defaultDocumentId: 'bilig-demo' }
        },
        loadSession: async () => ({ authToken: 'token-123' }),
      },
    })

    const done = new Promise<void>((resolve, reject) => {
      const subscription = actor.subscribe((snapshot) => {
        if (snapshot.matches('retrying')) {
          sawRetrying = true
          expect(snapshot.context.error).toBe('temporary failure')
          return
        }
        if (snapshot.matches('ready')) {
          subscription.unsubscribe()
          resolve()
          return
        }
        if (snapshot.matches('failed')) {
          subscription.unsubscribe()
          reject(new Error(snapshot.context.error ?? 'automatic retry failed'))
        }
      })
    })

    actor.start()
    await done

    expect(sawRetrying).toBe(true)
    expect(attempts).toBe(2)
    expect(actor.getSnapshot().context.error).toBeNull()
    actor.stop()
  })

  it('keeps recovering from a failed bootstrap when failed retry is configured', async () => {
    let attempts = 0
    let sawFailed = false
    const machine = createBootstrapMachine<{ defaultDocumentId: string }, { authToken: string }>()
    const actor = createActor(machine, {
      input: {
        autoRetryDelayMs: 1,
        failedRetryDelayMs: 1,
        maxAutoRetryAttempts: 1,
        loadConfig: async () => {
          attempts += 1
          if (attempts < 3) {
            throw new Error(`temporary failure ${attempts}`)
          }
          return { defaultDocumentId: 'bilig-demo' }
        },
        loadSession: async () => ({ authToken: 'token-123' }),
      },
    })

    const done = new Promise<void>((resolve, reject) => {
      const subscription = actor.subscribe((snapshot) => {
        if (snapshot.matches('failed')) {
          if (attempts <= 2) {
            sawFailed = true
            expect(snapshot.context.error).toBe('temporary failure 2')
            return
          }
          subscription.unsubscribe()
          reject(new Error(snapshot.context.error ?? 'failed retry did not recover'))
          return
        }
        if (snapshot.matches('ready')) {
          subscription.unsubscribe()
          resolve()
          return
        }
      })
    })

    actor.start()
    await done

    expect(sawFailed).toBe(true)
    expect(attempts).toBe(3)
    expect(actor.getSnapshot().context.error).toBeNull()
    actor.stop()
  })
})
