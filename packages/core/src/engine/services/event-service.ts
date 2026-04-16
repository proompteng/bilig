import { Effect } from 'effect'
import type { EngineEvent } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import type { EngineRuntimeState } from '../runtime-state.js'

export interface EngineEventService {
  readonly subscribe: (listener: (event: EngineEvent) => void) => Effect.Effect<() => void>
  readonly subscribeCell: (sheetName: string, address: string, listener: () => void) => Effect.Effect<() => void>
  readonly subscribeCells: (sheetName: string, addresses: readonly string[], listener: () => void) => Effect.Effect<() => void>
  readonly subscribeBatches: (listener: (batch: EngineOpBatch) => void) => Effect.Effect<() => void>
}

export function createEngineEventService(state: Pick<EngineRuntimeState, 'events' | 'workbook' | 'batchListeners'>): EngineEventService {
  return {
    subscribe(listener) {
      return Effect.sync(() => state.events.subscribe(listener))
    },
    subscribeCell(sheetName, address, listener) {
      return Effect.sync(() => {
        const cellIndex = state.workbook.getCellIndex(sheetName, address)
        if (cellIndex !== undefined) {
          return state.events.subscribeCellIndex(cellIndex, listener)
        }
        return state.events.subscribeCellAddress(`${sheetName}!${address}`, listener)
      })
    },
    subscribeCells(sheetName, addresses, listener) {
      return Effect.sync(() => {
        const cellIndices: number[] = []
        const qualifiedAddresses: string[] = []
        addresses.forEach((address) => {
          const cellIndex = state.workbook.getCellIndex(sheetName, address)
          if (cellIndex !== undefined) {
            cellIndices.push(cellIndex)
            return
          }
          qualifiedAddresses.push(`${sheetName}!${address}`)
        })
        return state.events.subscribeCells(cellIndices, qualifiedAddresses, listener)
      })
    },
    subscribeBatches(listener) {
      return Effect.sync(() => {
        state.batchListeners.add(listener)
        return () => {
          state.batchListeners.delete(listener)
        }
      })
    },
  }
}
