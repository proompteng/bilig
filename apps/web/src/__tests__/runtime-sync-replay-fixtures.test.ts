import { readdirSync } from 'node:fs'
import { extname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, it } from 'vitest'
import { loadReplayFixture } from '@bilig/test-fuzz'
import {
  applyRuntimeSyncAction,
  assertRuntimeSyncState,
  createRuntimeSyncHarness,
  type RuntimeSyncAction,
} from './runtime-sync-fuzz-helpers.js'

const fixturesDir = fileURLToPath(new URL('./fixtures/fuzz-replays', import.meta.url))

describe('runtime sync replay fixtures', () => {
  for (const fixture of loadRuntimeSyncReplayFixtures()) {
    it(`replays ${fixture.name}`, async () => {
      const { runtime, model } = await createRuntimeSyncHarness()
      try {
        await fixture.actions.reduce<Promise<void>>(async (previous, action) => {
          await previous
          await applyRuntimeSyncAction(runtime, model, action)
          assertRuntimeSyncState(runtime, model)
        }, Promise.resolve())
      } finally {
        runtime.dispose()
      }
    })
  }
})

type RuntimeSyncReplayFixture = {
  name: string
  actions: RuntimeSyncAction[]
}

function loadRuntimeSyncReplayFixtures(): RuntimeSyncReplayFixture[] {
  return readdirSync(fixturesDir)
    .filter((fileName) => extname(fileName) === '.json')
    .toSorted((left, right) => left.localeCompare(right))
    .map((fileName) => {
      const fixture = loadReplayFixture(`${fixturesDir}/${fileName}`)
      if (!Array.isArray(fixture.counterexample)) {
        throw new Error(`Runtime sync replay fixture ${fileName} is missing counterexample actions`)
      }
      return {
        name: fileName.replace(/\.json$/u, ''),
        actions: fixture.counterexample.map((action) => parseRuntimeSyncAction(action, fileName)),
      }
    })
}

function parseRuntimeSyncAction(value: unknown, fileName: string): RuntimeSyncAction {
  if (!isRecord(value) || typeof value['kind'] !== 'string') {
    throw new Error(`Invalid runtime sync replay action in ${fileName}`)
  }
  const kind = value['kind']
  switch (kind) {
    case 'submit':
    case 'ack':
      return { kind }
    case 'local':
    case 'remote':
      if (typeof value['address'] === 'string' && typeof value['value'] === 'number') {
        return {
          kind,
          address: value['address'],
          value: value['value'],
        }
      }
      break
  }
  throw new Error(`Invalid runtime sync replay action in ${fileName}: ${JSON.stringify(value)}`)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
