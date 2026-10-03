import { describe, expect, it, vi } from 'vitest'
import { WorkPaperEmitter } from '../work-paper-emitter.js'
import type { WorkPaperChange } from '../work-paper-types.js'

describe('WorkPaperEmitter', () => {
  it('dispatches legacy and detailed sheet events', () => {
    const emitter = new WorkPaperEmitter()
    const legacy = vi.fn()
    const detailed = vi.fn()

    emitter.on('sheetRenamed', legacy)
    emitter.onDetailed('sheetRenamed', detailed)

    emitter.emitDetailed({
      eventName: 'sheetRenamed',
      payload: { sheetId: 7, oldName: 'Sheet1', newName: 'Data' },
    })

    expect(legacy).toHaveBeenCalledWith('Sheet1', 'Data')
    expect(detailed).toHaveBeenCalledWith({ sheetId: 7, oldName: 'Sheet1', newName: 'Data' })
  })

  it('removes once listeners after the first detailed event', () => {
    const emitter = new WorkPaperEmitter()
    const changes: WorkPaperChange[] = []
    const legacy = vi.fn()
    const detailed = vi.fn()

    emitter.once('valuesUpdated', legacy)
    emitter.onceDetailed('valuesUpdated', detailed)

    emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })
    emitter.emitDetailed({ eventName: 'valuesUpdated', payload: { changes } })

    expect(legacy).toHaveBeenCalledTimes(1)
    expect(legacy).toHaveBeenCalledWith(changes)
    expect(detailed).toHaveBeenCalledTimes(1)
    expect(detailed).toHaveBeenCalledWith({ changes })
  })

  it('tracks listener presence across off and clear', () => {
    const emitter = new WorkPaperEmitter()
    const legacy = vi.fn()
    const detailed = vi.fn()

    expect(emitter.hasAnyListeners()).toBe(false)
    expect(emitter.hasListeners('evaluationSuspended')).toBe(false)

    emitter.on('evaluationSuspended', legacy)
    emitter.onDetailed('valuesUpdated', detailed)

    expect(emitter.hasAnyListeners()).toBe(true)
    expect(emitter.hasListeners('evaluationSuspended')).toBe(true)
    expect(emitter.hasListeners('valuesUpdated')).toBe(true)

    emitter.off('evaluationSuspended', legacy)

    expect(emitter.hasListeners('evaluationSuspended')).toBe(false)
    expect(emitter.hasAnyListeners()).toBe(true)

    emitter.clear()

    expect(emitter.hasAnyListeners()).toBe(false)
    expect(emitter.hasListeners('valuesUpdated')).toBe(false)
  })

  it('does not double count duplicate listener registrations', () => {
    const emitter = new WorkPaperEmitter()
    const legacy = vi.fn()
    const detailed = vi.fn()

    emitter.on('valuesUpdated', legacy)
    emitter.on('valuesUpdated', legacy)
    emitter.onDetailed('valuesUpdated', detailed)
    emitter.onDetailed('valuesUpdated', detailed)

    expect(emitter.hasAnyListeners()).toBe(true)
    expect(emitter.hasListeners('valuesUpdated')).toBe(true)

    emitter.off('valuesUpdated', legacy)
    emitter.offDetailed('valuesUpdated', detailed)

    expect(emitter.hasAnyListeners()).toBe(false)
    expect(emitter.hasListeners('valuesUpdated')).toBe(false)
  })

  it('does not dispatch listeners removed with offDetailed', () => {
    const emitter = new WorkPaperEmitter()
    const detailed = vi.fn()

    emitter.onDetailed('sheetAdded', detailed)
    emitter.offDetailed('sheetAdded', detailed)
    emitter.emitDetailed({ eventName: 'sheetAdded', payload: { sheetId: 1, sheetName: 'Sheet1' } })

    expect(detailed).not.toHaveBeenCalled()
  })
})
