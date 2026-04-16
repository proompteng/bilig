import { fileURLToPath, URL } from 'node:url'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { createViteAliasRecord } from '../../scripts/workspace-resolution.js'

const syncServerTarget = process.env['BILIG_SYNC_SERVER_TARGET'] ?? `http://127.0.0.1:${process.env['BILIG_SYNC_SERVER_PORT'] ?? '4321'}`

export const crossOriginIsolationHeaders = {
  'Cross-Origin-Opener-Policy': 'same-origin',
  'Cross-Origin-Embedder-Policy': 'require-corp',
  'Origin-Agent-Cluster': '?1',
} as const

function includesAny(id: string, patterns: readonly string[]): boolean {
  const normalizedId = id.replaceAll('\\', '/')
  return patterns.some((pattern) => normalizedId.includes(pattern))
}

const codeSplittingGroups = [
  {
    name: 'react-vendor',
    priority: 70,
    test(id: string) {
      return includesAny(id, ['/node_modules/react/', '/node_modules/react-dom/', '/node_modules/scheduler/'])
    },
  },
  {
    name: 'sync-vendor',
    priority: 60,
    test(id: string) {
      return includesAny(id, ['/node_modules/@rocicorp/zero/', '/packages/zero-sync/'])
    },
  },
  {
    name: 'grid-vendor',
    priority: 50,
    test(id: string) {
      return includesAny(id, [
        '/node_modules/marked/',
        '/node_modules/react-number-format/',
        '/node_modules/react-responsive-carousel/',
        '/node_modules/lodash/',
      ])
    },
  },
  {
    name: 'icons-vendor',
    priority: 40,
    test(id: string) {
      return includesAny(id, ['/node_modules/lucide-react/'])
    },
  },
  {
    name: 'formula-vendor',
    priority: 30,
    test(id: string) {
      return includesAny(id, ['/packages/formula/'])
    },
  },
  {
    name: 'engine-vendor',
    priority: 20,
    test(id: string) {
      return includesAny(id, ['/packages/binary-protocol/', '/packages/protocol/', '/packages/core/', '/packages/wasm-kernel/'])
    },
  },
  {
    name: 'workbook-vendor',
    priority: 10,
    test(id: string) {
      return includesAny(id, [
        '/packages/grid/',
        '/packages/renderer/',
        '/packages/storage-browser/',
        '/packages/worker-transport/',
        '/packages/workbook-domain/',
        '/apps/web/src/WorkerWorkbookApp.tsx',
        '/apps/web/src/projected-viewport-store.ts',
        '/apps/web/src/worker-runtime.ts',
        '/apps/web/src/zero/',
      ])
    },
  },
]

export default defineConfig({
  plugins: [react(), tailwindcss()],
  optimizeDeps: {
    exclude: ['@sqlite.org/sqlite-wasm'],
  },
  build: {
    // Keep the startup shell bounded to the entry module and shell CSS. The
    // heavy workbook and sync chunks still load through the module graph, but
    // they should not all count as eager shell bytes in the release contract.
    modulePreload: false,
    rolldownOptions: {
      output: {
        codeSplitting: {
          groups: codeSplittingGroups,
        },
      },
    },
  },
  resolve: {
    alias: createViteAliasRecord({
      '@bilig/formula/program-arena': fileURLToPath(new URL('../../packages/formula/src/program-arena.ts', import.meta.url)),
    }),
  },
  server: {
    headers: crossOriginIsolationHeaders,
    proxy: {
      '/runtime-config.json': {
        target: syncServerTarget,
        changeOrigin: true,
      },
      '/v2': {
        target: syncServerTarget,
        changeOrigin: true,
      },
      '/api/zero': {
        target: syncServerTarget,
        changeOrigin: true,
      },
      '/zero': {
        target: syncServerTarget,
        changeOrigin: true,
        ws: true,
      },
      '/healthz': {
        target: syncServerTarget,
        changeOrigin: true,
      },
    },
  },
  preview: {
    headers: crossOriginIsolationHeaders,
  },
})
