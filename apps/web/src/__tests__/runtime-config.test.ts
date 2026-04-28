// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { normalizeRuntimeConfigUserId, resolveRuntimeConfig } from '../runtime-config'
import { resolveZeroCacheUrl } from '../zero-connection'

const BASE_CONFIG = {
  zeroCacheUrl: 'http://127.0.0.1:4848',
  defaultDocumentId: 'bilig-demo',
  persistState: true,
  currentUserId: 'guest:test',
} as const

describe('resolveRuntimeConfig', () => {
  afterEach(() => {
    window.history.replaceState({}, '', '/')
    vi.unstubAllGlobals()
  })

  it('keeps the explicit document id when one is provided', () => {
    window.history.replaceState({}, '', '/?document=multiplayer-debug')

    expect(resolveRuntimeConfig(BASE_CONFIG)).toMatchObject({
      documentId: 'multiplayer-debug',
      currentUserId: 'guest:test',
      persistState: true,
      workbookAgentEnabled: false,
    })
  })

  it('passes through the assistant availability flag from the app runtime config', () => {
    expect(resolveRuntimeConfig({ ...BASE_CONFIG, workbookAgentEnabled: true })).toMatchObject({
      workbookAgentEnabled: true,
    })
  })

  it('uses an ephemeral document under webdriver when no explicit document is present', () => {
    vi.stubGlobal('navigator', { webdriver: true })

    expect(resolveRuntimeConfig(BASE_CONFIG)).toMatchObject({
      persistState: false,
    })
    expect(resolveRuntimeConfig(BASE_CONFIG).documentId).toMatch(/^bilig-demo:/)
  })

  it('resolves relative zero cache URLs against the current origin', () => {
    expect(resolveZeroCacheUrl('/zero', 'http://127.0.0.1:4180')).toBe('http://127.0.0.1:4180/zero')
  })

  it('normalizes the workbook current user id to the resolved runtime session user', () => {
    expect(
      normalizeRuntimeConfigUserId(BASE_CONFIG, {
        userId: 'guest:session-user',
      }),
    ).toEqual({
      ...BASE_CONFIG,
      currentUserId: 'guest:session-user',
    })
  })
})
