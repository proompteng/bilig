import { utcDateToExcelSerial } from '@bilig/formula'
import type { RecalcVolatileState, U32 } from '../runtime-state.js'

export function createRecalcVolatileState(now: () => Date): RecalcVolatileState {
  return {
    nowSerial: utcDateToExcelSerial(now()),
    randomValues: [],
    randomCursor: 0,
  }
}

function ensureVolatileRandomValues(state: RecalcVolatileState, count: number, random: () => number): void {
  const needed = state.randomCursor + count - state.randomValues.length
  if (needed <= 0) {
    return
  }
  for (let index = 0; index < needed; index += 1) {
    state.randomValues.push(random())
  }
}

export function consumeVolatileRandomValues(state: RecalcVolatileState, count: number, random: () => number): Float64Array {
  ensureVolatileRandomValues(state, count, random)
  const values = state.randomValues.slice(state.randomCursor, state.randomCursor + count)
  state.randomCursor += count
  return Float64Array.from(values)
}

export function toOrderedUint32(ordered: readonly number[] | U32, orderedCount: number): U32 {
  if (ordered instanceof Uint32Array) {
    return ordered
  }
  const next = new Uint32Array(orderedCount)
  for (let index = 0; index < orderedCount; index += 1) {
    next[index] = ordered[index] ?? 0
  }
  return next
}
