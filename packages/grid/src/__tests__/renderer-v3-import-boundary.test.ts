import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { dirname, join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, test } from 'vitest'

const GRID_SRC_ROOT = dirname(dirname(fileURLToPath(import.meta.url)))

function collectSourceFiles(path: string): string[] {
  const stat = statSync(path)
  if (stat.isFile()) {
    return path.endsWith('.ts') || path.endsWith('.tsx') ? [path] : []
  }
  return readdirSync(path).flatMap((entry) => collectSourceFiles(join(path, entry)))
}

describe('renderer v3 import boundary', () => {
  test('legacy renderer-v2 directory is deleted', () => {
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2'))).toBe(false)
  })

  test('legacy mounted V2 pane renderer surface is deleted', () => {
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/WorkbookPaneRendererV2.tsx'))).toBe(false)
  })

  test('legacy V2 workbook backend surface is deleted', () => {
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/workbook-typegpu-backend.ts'))).toBe(false)
  })

  test('legacy V2 draw and surface runtime surfaces are deleted', () => {
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/typegpu-render-pass.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/typegpu-surface.ts'))).toBe(false)
  })

  test('legacy V2 pane resource stack is deleted', () => {
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/pane-layout.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/pane-scene-types.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/pane-buffer-cache.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/typegpu-buffer-pool.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/typegpu-backend.ts'))).toBe(false)
  })

  test('legacy V2 text and atlas resource stack is deleted', () => {
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/glyphAtlasV2.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/text-glyph-buffer.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/line-text-layout.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/line-text-quad-buffer.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/typegpu-atlas-manager.ts'))).toBe(false)
  })

  test('legacy V2 loop and camera store shims are deleted', () => {
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/gridRenderLoop.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/gridCameraStore.ts'))).toBe(false)
  })

  test('legacy V2 scene-packet, cache, debug, and barrel files are deleted', () => {
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/index.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/scene-packet-v2.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/scene-packet-validator.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/tile-gpu-cache.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/grid-render-contract.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/grid-render-counters.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/grid-render-layers.ts'))).toBe(false)
    expect(existsSync(join(GRID_SRC_ROOT, 'renderer-v2/render-debug-hud.ts'))).toBe(false)
  })

  test('product grid paths do not import renderer-v2 modules', () => {
    const files = collectSourceFiles(GRID_SRC_ROOT).filter((file) => {
      const local = relative(GRID_SRC_ROOT, file)
      return !local.startsWith('__tests__/') && !local.startsWith('renderer-v2/')
    })
    const offenders = files.flatMap((file) => {
      const source = readFileSync(file, 'utf8')
      return source.includes('renderer-v2') ? [relative(GRID_SRC_ROOT, file)] : []
    })

    expect(offenders).toEqual([])
  })

  test('V3 pane renderer shell does not own renderer readiness in React state', () => {
    const source = readFileSync(join(GRID_SRC_ROOT, 'renderer-v3/WorkbookPaneRendererV3.tsx'), 'utf8')

    expect(source).not.toContain('useState')
    expect(source).not.toContain('WorkbookPaneRendererRuntimeV3')
    expect(source).not.toContain('WorkbookPaneSurfaceRuntimeV3')
  })
})
