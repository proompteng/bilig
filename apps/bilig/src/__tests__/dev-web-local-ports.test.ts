import { describe, expect, it, vi } from 'vitest'

import {
  canUsePort,
  resolvePreferredPort,
  resolvePreferredZeroPort,
  resolveRequestedOrAvailablePort,
} from '../../../../scripts/dev-web-local-ports.js'

describe('resolvePreferredPort', () => {
  it('uses the configured port when present', () => {
    expect(resolvePreferredPort('6001', 55432)).toBe(6001)
  })

  it('falls back to the default port', () => {
    expect(resolvePreferredPort(undefined, 55432)).toBe(55432)
  })
})

describe('resolvePreferredZeroPort', () => {
  it('prefers an explicit zero port', () => {
    expect(resolvePreferredZeroPort('4900', 'http://127.0.0.1:4848', 4848)).toBe(4900)
  })

  it('falls back to the configured upstream port', () => {
    expect(resolvePreferredZeroPort(undefined, 'http://127.0.0.1:4848', 4900)).toBe(4848)
  })

  it('falls back to the default port when nothing is configured', () => {
    expect(resolvePreferredZeroPort(undefined, undefined, 4848)).toBe(4848)
  })
})

describe('resolveRequestedOrAvailablePort', () => {
  it('keeps an explicit port when it is available', async () => {
    const canUseRequestedPort = vi.fn(async () => true)

    await expect(
      resolveRequestedOrAvailablePort({
        preferredPort: 55432,
        explicitPort: '55432',
        label: 'Postgres port',
        canUseRequestedPort,
      }),
    ).resolves.toBe(55432)
    expect(canUseRequestedPort).toHaveBeenCalledWith(55432)
  })

  it('throws when an explicit port is unavailable', async () => {
    await expect(
      resolveRequestedOrAvailablePort({
        preferredPort: 55432,
        explicitPort: '55432',
        label: 'Postgres port',
        canUseRequestedPort: async () => false,
      }),
    ).rejects.toThrow('Postgres port 55432 is already in use.')
  })

  it('walks forward until it finds an open port', async () => {
    const canUseRequestedPort = vi.fn(async (port: number) => port === 55434)

    await expect(
      resolveRequestedOrAvailablePort({
        preferredPort: 55432,
        explicitPort: undefined,
        label: 'Postgres port starting at 55432',
        canUseRequestedPort,
      }),
    ).resolves.toBe(55434)
    expect(canUseRequestedPort).toHaveBeenCalledTimes(3)
  })
})

describe('canUsePort', () => {
  it('rejects ports that already have listeners before probing', async () => {
    const bindProbe = vi.fn(async () => true)

    await expect(
      canUsePort({
        port: 55432,
        listListeningPids: () => ['952'],
        bindProbe,
      }),
    ).resolves.toBe(false)
    expect(bindProbe).not.toHaveBeenCalled()
  })

  it('falls back to a bind probe when there are no listeners', async () => {
    const bindProbe = vi.fn(async () => true)

    await expect(
      canUsePort({
        port: 55434,
        listListeningPids: () => [],
        bindProbe,
      }),
    ).resolves.toBe(true)
    expect(bindProbe).toHaveBeenCalledWith(55434)
  })
})
