import { describe, expect, test } from 'vitest'
import { resolveWorkbookRendererMode } from '../workbookRendererMode.js'

describe('resolveWorkbookRendererMode', () => {
  test('defaults to the current typegpu renderer', () => {
    expect(resolveWorkbookRendererMode()).toBe('typegpu-v1')
  })

  test('prefers explicit mode over query and env', () => {
    expect(
      resolveWorkbookRendererMode({
        env: 'typegpu-v2',
        explicit: 'canvas-fallback',
        search: '?workbookRenderer=typegpu-v1',
      }),
    ).toBe('canvas-fallback')
  })

  test('resolves query mode before env mode', () => {
    expect(resolveWorkbookRendererMode({ env: 'typegpu-v1', search: '?workbookRenderer=typegpu-v2' })).toBe('typegpu-v2')
  })

  test('ignores invalid modes', () => {
    expect(resolveWorkbookRendererMode({ env: 'wat', search: '?workbookRenderer=nope' })).toBe('typegpu-v1')
  })
})
