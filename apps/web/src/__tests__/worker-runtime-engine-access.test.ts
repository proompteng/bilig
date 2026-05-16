import { describe, expect, it } from 'vitest'
import { WorkerRuntimeSnapshotCaches } from '../worker-runtime-snapshot-caches.js'
import { ensureAuthoritativeEngine } from '../worker-runtime-engine-access.js'

describe('worker runtime engine access', () => {
  it('seeds authoritative caches when lazily creating an engine', async () => {
    const caches = new WorkerRuntimeSnapshotCaches()

    const engine = await ensureAuthoritativeEngine({
      authoritativeEngine: null,
      documentId: 'worker-runtime-engine-access-doc',
      replicaId: 'replica-2',
      snapshotCaches: caches,
      async resolveAuthoritativeStateInput() {
        return { snapshot: null, replica: null }
      },
    })

    const resolved = caches.resolveAuthoritativeState({
      exportSnapshot: null,
      exportReplica: null,
    })
    expect(resolved.snapshot).toEqual(engine.exportSnapshot())
    expect(resolved.replica).toEqual(engine.exportReplicaSnapshot())
  })
})
