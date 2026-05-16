import { describe, expect, it, vi } from 'vitest'
import { loadRuntimeConfig, parseRuntimeConfig } from '../runtime-config.js'

const validConfig = {
  zeroCacheUrl: 'http://127.0.0.1:4848',
  defaultDocumentId: 'bilig-demo',
  persistState: true,
  currentUserId: 'guest:test',
} as const

describe('runtime config', () => {
  it('loads and validates runtime config responses', async () => {
    const fetchImpl = vi.fn<typeof fetch>(async () => {
      return new Response(JSON.stringify(validConfig), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    })

    await expect(loadRuntimeConfig(fetchImpl)).resolves.toEqual(validConfig)
    expect(fetchImpl).toHaveBeenCalledWith('/runtime-config.json', {
      headers: { accept: 'application/json' },
    })
  })

  it('surfaces failed runtime config responses with status', async () => {
    const fetchImpl = vi.fn<typeof fetch>(async () => {
      return new Response('not found', { status: 503 })
    })

    await expect(loadRuntimeConfig(fetchImpl)).rejects.toThrow('Failed to load runtime config (503)')
  })

  it('surfaces malformed runtime config JSON with stable copy', async () => {
    const fetchImpl = vi.fn<typeof fetch>(async () => {
      return new Response('{', {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    })

    await expect(loadRuntimeConfig(fetchImpl)).rejects.toThrow('Runtime config response returned malformed JSON')
  })

  it('rejects malformed runtime config fields', () => {
    expect(() => parseRuntimeConfig({ ...validConfig, persistState: 'true' })).toThrow(
      'Runtime config field persistState must be a boolean',
    )
  })
})
