import { describe, expect, test } from 'vitest'
import {
  MIN_TYPEGPU_RECT_VERTEX_CAPACITY,
  MIN_TYPEGPU_TEXT_VERTEX_CAPACITY,
  resolveTypeGpuVertexBufferCapacity,
} from '../renderer-v2/typegpu-backend.js'

describe('typegpu backend buffer capacity policy', () => {
  test('uses fixed minimum buckets for new rect and text buffers', () => {
    expect(
      resolveTypeGpuVertexBufferCapacity({
        currentCapacity: 0,
        minimumCapacity: MIN_TYPEGPU_RECT_VERTEX_CAPACITY,
        nextCount: 12,
      }),
    ).toBe(MIN_TYPEGPU_RECT_VERTEX_CAPACITY)
    expect(
      resolveTypeGpuVertexBufferCapacity({
        currentCapacity: 0,
        minimumCapacity: MIN_TYPEGPU_TEXT_VERTEX_CAPACITY,
        nextCount: 12,
      }),
    ).toBe(MIN_TYPEGPU_TEXT_VERTEX_CAPACITY)
  })

  test('keeps sufficient existing buffers and grows by deterministic buckets', () => {
    expect(
      resolveTypeGpuVertexBufferCapacity({
        currentCapacity: 8192,
        minimumCapacity: MIN_TYPEGPU_RECT_VERTEX_CAPACITY,
        nextCount: 8192,
      }),
    ).toBe(8192)
    expect(
      resolveTypeGpuVertexBufferCapacity({
        currentCapacity: 8192,
        minimumCapacity: MIN_TYPEGPU_RECT_VERTEX_CAPACITY,
        nextCount: 8193,
      }),
    ).toBe(16384)
  })
})
