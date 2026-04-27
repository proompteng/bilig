// @vitest-environment jsdom
import { describe, expect, test, vi } from 'vitest'
import {
  EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3,
  WorkbookPaneSurfaceRuntimeV3,
  resolveWorkbookPaneSurfaceSizeV3,
  type WorkbookPaneSurfaceSnapshotV3,
} from '../renderer-v3/workbook-pane-surface-runtime.js'

function defineHostSize(host: HTMLElement, width: number, height: number): void {
  Object.defineProperty(host, 'clientWidth', { configurable: true, value: width })
  Object.defineProperty(host, 'clientHeight', { configurable: true, value: height })
}

async function flushMicrotasks(): Promise<void> {
  await Promise.resolve()
  await Promise.resolve()
}

describe('WorkbookPaneSurfaceRuntimeV3', () => {
  test('resolves canvas surface size from host dimensions and DPR', () => {
    const host = document.createElement('div')
    defineHostSize(host, 640, 360)

    expect(resolveWorkbookPaneSurfaceSizeV3({ dpr: 2, host })).toEqual({
      dpr: 2,
      height: 360,
      pixelHeight: 720,
      pixelWidth: 1280,
      width: 640,
    })
  })

  test('owns backend startup, surface sync, resize updates, and teardown', async () => {
    const backend = { id: 'backend-1' }
    const host = document.createElement('div')
    const canvas = document.createElement('canvas')
    defineHostSize(host, 640, 360)
    let resizeListener: ResizeObserverCallback | null = null
    const disconnect = vi.fn()
    const observe = vi.fn()
    const createBackend = vi.fn(async () => backend)
    const destroyBackend = vi.fn()
    const syncSurface = vi.fn()
    const snapshots: WorkbookPaneSurfaceSnapshotV3[] = []
    const runtime = new WorkbookPaneSurfaceRuntimeV3({
      createBackend,
      createResizeObserver: (listener) => {
        resizeListener = listener
        return { disconnect, observe }
      },
      destroyBackend,
      getDevicePixelRatio: () => 2,
      syncSurface,
    })

    runtime.subscribe((snapshot) => snapshots.push(snapshot))
    runtime.setHost(host)
    runtime.setActive(true)
    runtime.setCanvas(canvas)
    await flushMicrotasks()

    expect(createBackend).toHaveBeenCalledWith(canvas)
    expect(observe).toHaveBeenCalledWith(host)
    expect(snapshots.at(-1)).toMatchObject({
      backend,
      surface: { height: 360, pixelHeight: 720, pixelWidth: 1280, width: 640 },
      webGpuReady: true,
    })
    expect(syncSurface).toHaveBeenCalledWith({
      backend,
      canvas,
      size: expect.objectContaining({ height: 360, pixelHeight: 720, pixelWidth: 1280, width: 640 }),
    })

    defineHostSize(host, 800, 400)
    resizeListener?.([], {
      disconnect() {},
      observe() {},
      unobserve() {},
    } as ResizeObserver)

    expect(snapshots.at(-1)).toMatchObject({
      surface: { height: 400, pixelHeight: 800, pixelWidth: 1600, width: 800 },
      webGpuReady: true,
    })
    expect(syncSurface).toHaveBeenLastCalledWith({
      backend,
      canvas,
      size: expect.objectContaining({ height: 400, pixelHeight: 800, pixelWidth: 1600, width: 800 }),
    })

    runtime.setActive(false)

    expect(destroyBackend).toHaveBeenCalledWith(backend)
    expect(snapshots.at(-1)).toMatchObject({
      backend: null,
      webGpuReady: false,
    })
    expect(runtime.getSnapshot()).toMatchObject({
      backend: null,
      webGpuReady: false,
    })

    runtime.dispose()
    expect(disconnect).toHaveBeenCalled()
  })

  test('destroys a stale async backend that resolves after deactivation', async () => {
    const host = document.createElement('div')
    const canvas = document.createElement('canvas')
    const staleBackend = { id: 'stale' }
    defineHostSize(host, 640, 360)
    let resolveBackend: ((backend: object | null) => void) | null = null
    const pendingBackend = new Promise<object | null>((resolve) => {
      resolveBackend = resolve
    })
    const destroyBackend = vi.fn()
    const snapshots: WorkbookPaneSurfaceSnapshotV3[] = []
    const runtime = new WorkbookPaneSurfaceRuntimeV3({
      createBackend: () => pendingBackend,
      createResizeObserver: () => null,
      destroyBackend,
      syncSurface: vi.fn(),
    })

    runtime.subscribe((snapshot) => snapshots.push(snapshot))
    runtime.setHost(host)
    runtime.setActive(true)
    runtime.setCanvas(canvas)
    runtime.setActive(false)
    resolveBackend?.(staleBackend)
    await pendingBackend
    await flushMicrotasks()

    expect(destroyBackend).toHaveBeenCalledWith(staleBackend)
    expect(runtime.getSnapshot()).toEqual({
      ...EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3,
      surface: expect.objectContaining({ height: 360, width: 640 }),
    })
    expect(snapshots.at(-1)).toMatchObject({
      backend: null,
      webGpuReady: false,
    })
  })
})
