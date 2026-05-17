import type { ReplicaState } from '../replica-state.js'
import type { SpreadsheetEngine } from '../engine.js'
import type { EngineOperationService } from '../engine/services/operation-service.js'

function isEngineOperationService(value: unknown): value is EngineOperationService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return typeof Reflect.get(value, 'applyBatch') === 'function' && typeof Reflect.get(value, 'applyDerivedOp') === 'function'
}

function isReplicaState(value: unknown): value is ReplicaState {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'replicaId') === 'string' &&
    typeof Reflect.get(value, 'clock') === 'object' &&
    Reflect.get(value, 'appliedBatchIds') instanceof Set
  )
}

export function getOperationService(engine: SpreadsheetEngine): EngineOperationService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (!isEngineOperationService(operations)) {
    throw new TypeError('Expected engine operation service')
  }
  return operations
}

export function getReplicaState(engine: SpreadsheetEngine): ReplicaState {
  const replicaState = Reflect.get(engine, 'replicaState')
  if (!isReplicaState(replicaState)) {
    throw new TypeError('Expected engine replica state')
  }
  return replicaState
}
