import { describe, expect, it } from 'vitest'
import { attachRuntimeSnapshot, SpreadsheetEngine } from '@bilig/core'

import { WorkPaper } from '../index.js'

function expectWorkPaperTimeout(run: () => void): void {
  let thrown: unknown
  try {
    run()
  } catch (error) {
    thrown = error
  }

  expect(thrown).toBeInstanceOf(Error)
  if (!(thrown instanceof Error)) {
    throw new Error('Expected WorkPaper evaluation timeout')
  }
  expect(thrown.name).toBe('WorkPaperEvaluationTimeoutError')
  expect(thrown.message).toContain('timed out')
}

async function buildRuntimeImageSnapshot(): Promise<ReturnType<SpreadsheetEngine['exportSnapshot']>> {
  const source = new SpreadsheetEngine({
    workbookName: 'evaluation-timeout-source',
    replicaId: 'evaluation-timeout-source',
  })
  await source.ready()
  source.createSheet('Sheet1')
  source.setCellValue('Sheet1', 'A1', 1)
  source.setCellFormula('Sheet1', 'B1', 'A1+1')
  return source.exportSnapshot()
}

describe('evaluation timeout coverage', () => {
  it('applies the evaluation budget while importing compatible runtime snapshots from sheets', async () => {
    const runtimeSnapshot = await buildRuntimeImageSnapshot()
    const sheets = attachRuntimeSnapshot({ Sheet1: [[1, '=A1+1']] }, runtimeSnapshot)

    expectWorkPaperTimeout(() => {
      WorkPaper.buildFromSheets(sheets, {
        evaluationTimeoutMs: 0,
        maxColumns: 2,
        maxRows: 1,
        useColumnIndex: true,
      })
    })
  })

  it('applies the evaluation budget while importing compatible runtime snapshots directly', async () => {
    const runtimeSnapshot = await buildRuntimeImageSnapshot()

    expectWorkPaperTimeout(() => {
      WorkPaper.buildFromSnapshot(runtimeSnapshot, {
        evaluationTimeoutMs: 0,
        maxColumns: 2,
        maxRows: 1,
        useColumnIndex: true,
      })
    })
  })
})
