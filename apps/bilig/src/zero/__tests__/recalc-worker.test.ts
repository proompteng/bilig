import { WorkbookRuntimeManager } from '../../workbook-runtime/runtime-manager.js'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ZeroRecalcWorker } from '../recalc-worker.js'
import type { Queryable, QueryResultRow, WorkbookRuntimeStoreConnection } from '../store.js'

function createRuntimeStore(): WorkbookRuntimeStoreConnection {
  return {
    async query<T extends QueryResultRow = QueryResultRow>(): Promise<{ rows: T[] }> {
      return { rows: [] }
    },
    async run(): Promise<never> {
      throw new Error('not used')
    },
  }
}

class FailingRuntimeManager extends WorkbookRuntimeManager {
  readonly invalidateSpy = vi.fn()

  override async runExclusive<T>(): Promise<T> {
    throw new Error('engine failed')
  }

  override invalidate(documentId: string): void {
    this.invalidateSpy(documentId)
  }
}

describe('ZeroRecalcWorker', () => {
  let stderrWrite: ReturnType<typeof vi.spyOn>

  beforeEach(() => {
    vi.useFakeTimers()
    stderrWrite = vi.spyOn(process.stderr, 'write').mockImplementation(() => true)
  })

  afterEach(() => {
    stderrWrite.mockRestore()
    vi.useRealTimers()
  })

  it('logs lease failures instead of letting the scheduled tick crash the process', async () => {
    const query = vi.fn(async () => {
      throw new Error('database restarting')
    })
    const db: Queryable = {
      query,
    }
    const worker = new ZeroRecalcWorker(db, createRuntimeStore(), new WorkbookRuntimeManager(), 'worker-test')

    worker.start()
    await vi.runOnlyPendingTimersAsync()
    worker.stop()

    expect(query).toHaveBeenCalledOnce()
    expect(stderrWrite).toHaveBeenCalledWith(expect.stringContaining('[bilig] Zero recalc worker tick failed Error: database restarting'))
  })

  it('logs failure-persistence errors after invalidating the affected workbook', async () => {
    const query = vi
      .fn()
      .mockResolvedValueOnce({
        rows: [
          {
            id: 'job-1',
            workbook_id: 'workbook-1',
            from_revision: 1,
            to_revision: 2,
            dirty_regions_json: null,
            attempts: 1,
          },
        ],
      })
      .mockRejectedValueOnce(new Error('connection reset while recording failure'))
    const db: Queryable = {
      query,
    }
    const runtimeManager = new FailingRuntimeManager()
    const worker = new ZeroRecalcWorker(db, createRuntimeStore(), runtimeManager, 'worker-test')

    worker.start()
    await vi.runOnlyPendingTimersAsync()
    worker.stop()

    expect(runtimeManager.invalidateSpy).toHaveBeenCalledWith('workbook-1')
    expect(stderrWrite).toHaveBeenCalledWith(
      expect.stringContaining(
        '[bilig] Zero recalc worker failed to mark recalc job failed Error: connection reset while recording failure',
      ),
    )
  })
})
