import { afterEach, describe, expect, it, vi } from 'vitest'

import { createBrowserMetadataStore } from '../index.js'

describe('storage-browser', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('creates a persistence facade', async () => {
    const persistence = createBrowserMetadataStore({ databaseName: 'spec', storeName: 'state' })
    expect(typeof persistence.loadJson).toBe('function')
    expect(typeof persistence.saveJson).toBe('function')
    expect(typeof persistence.remove).toBe('function')
  })

  it('falls back when localStorage reads fail', async () => {
    vi.stubGlobal('localStorage', {
      getItem() {
        throw new Error('storage denied')
      },
      removeItem() {
        throw new Error('storage denied')
      },
      setItem() {
        throw new Error('storage denied')
      },
    })

    const persistence = createBrowserMetadataStore({ databaseName: 'spec', storeName: 'state' })

    await expect(persistence.loadJson('layout', () => ({ ok: true }))).resolves.toBeNull()
  })

  it('does not throw when invalid cached JSON cannot be removed', async () => {
    vi.stubGlobal('localStorage', {
      getItem() {
        return '{'
      },
      removeItem() {
        throw new Error('storage denied')
      },
      setItem() {
        throw new Error('storage denied')
      },
    })

    const persistence = createBrowserMetadataStore({ databaseName: 'spec', storeName: 'state' })

    await expect(persistence.loadJson('layout', () => ({ ok: true }))).resolves.toBeNull()
  })

  it('does not throw when localStorage cleanup fails during save and remove', async () => {
    vi.stubGlobal('localStorage', {
      getItem() {
        return null
      },
      removeItem() {
        throw new Error('storage denied')
      },
      setItem() {
        throw new Error('storage denied')
      },
    })

    const persistence = createBrowserMetadataStore({ databaseName: 'spec', storeName: 'state' })

    await expect(persistence.saveJson('layout', { ok: true })).resolves.toBeUndefined()
    await expect(persistence.remove('layout')).resolves.toBeUndefined()
  })
})
