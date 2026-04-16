import type { CompiledFormula } from '@bilig/formula'
import type { CompiledPlanRecord } from '../runtime-state.js'

export interface EngineCompiledPlanService {
  readonly intern: (source: string, compiled: CompiledFormula) => CompiledPlanRecord
  readonly get: (planId: number) => CompiledPlanRecord | undefined
  readonly release: (planId: number) => void
}

interface MutableCompiledPlanRecord extends CompiledPlanRecord {
  refCount: number
}

export function createEngineCompiledPlanService(): EngineCompiledPlanService {
  const records = new Map<number, MutableCompiledPlanRecord>()
  const planByCompiled = new WeakMap<CompiledFormula, MutableCompiledPlanRecord>()
  let nextPlanId = 1

  return {
    intern(source, compiled) {
      const existing = planByCompiled.get(compiled)
      if (existing !== undefined) {
        existing.refCount += 1
        return existing
      }
      const record: MutableCompiledPlanRecord = {
        id: nextPlanId,
        source,
        compiled,
        refCount: 1,
      }
      nextPlanId += 1
      records.set(record.id, record)
      planByCompiled.set(compiled, record)
      return record
    },
    get(planId) {
      return records.get(planId)
    },
    release(planId) {
      const existing = records.get(planId)
      if (!existing) {
        return
      }
      existing.refCount -= 1
      if (existing.refCount > 0) {
        return
      }
      records.delete(planId)
    },
  }
}
