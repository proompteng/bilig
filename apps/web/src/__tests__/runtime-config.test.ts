// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  createLocalOnlyRuntimeConfig,
  normalizeRuntimeConfigUserId,
  resolveRemoteSyncEnabled,
  resolveRuntimeConfig,
} from '../runtime-config'
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

  it('allows explicit document sessions to opt out of Zero client persistence for browser QA', () => {
    window.history.replaceState({}, '', '/?document=visual-smoke&persist=0')

    expect(resolveRuntimeConfig(BASE_CONFIG)).toMatchObject({
      documentId: 'visual-smoke',
      persistState: false,
    })
  })

  it('rejects malformed persistence query overrides instead of silently using defaults', () => {
    window.history.replaceState({}, '', '/?document=visual-smoke&persist=yes')

    expect(() => resolveRuntimeConfig(BASE_CONFIG)).toThrow(
      'persist query parameter must be "1", "true", "0", or "false" when set, got yes',
    )
  })

  it('passes through the assistant availability flag from the app runtime config', () => {
    expect(resolveRuntimeConfig({ ...BASE_CONFIG, workbookAgentEnabled: true })).toMatchObject({
      workbookAgentEnabled: true,
    })
  })

  it('uses the configured default document under webdriver when no explicit document is present', () => {
    vi.stubGlobal('navigator', { webdriver: true })

    expect(resolveRuntimeConfig(BASE_CONFIG)).toMatchObject({
      documentId: 'bilig-demo',
      persistState: true,
    })
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

  it('creates a local-only runtime config for standalone web development', () => {
    expect(createLocalOnlyRuntimeConfig('guest:local-dev')).toEqual({
      zeroCacheUrl: '/zero',
      defaultDocumentId: 'local-workbook',
      persistState: true,
      currentUserId: 'guest:local-dev',
      workbookAgentEnabled: false,
    })
  })

  it('defaults standalone Vite dev to local-only mode while preserving explicit and production remote sync', () => {
    expect(resolveRemoteSyncEnabled({ DEV: true })).toBe(false)
    expect(resolveRemoteSyncEnabled({ DEV: true, VITE_BILIG_REMOTE_SYNC: '' })).toBe(false)
    expect(resolveRemoteSyncEnabled({ DEV: true, VITE_BILIG_REMOTE_SYNC: '1' })).toBe(true)
    expect(resolveRemoteSyncEnabled({ DEV: true, VITE_BILIG_REMOTE_SYNC: 'true' })).toBe(true)
    expect(resolveRemoteSyncEnabled({ DEV: true, VITE_BILIG_REMOTE_SYNC: '0' })).toBe(false)
    expect(resolveRemoteSyncEnabled({ DEV: true, VITE_BILIG_REMOTE_SYNC: 'false' })).toBe(false)
    expect(resolveRemoteSyncEnabled({ DEV: false })).toBe(true)
    expect(() => resolveRemoteSyncEnabled({ DEV: true, VITE_BILIG_REMOTE_SYNC: 'yes' })).toThrow(
      'VITE_BILIG_REMOTE_SYNC must be "1", "true", "0", or "false" when set, got yes',
    )
  })
})
