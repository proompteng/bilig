import type { CompiledFormula } from '@bilig/formula'
import type { CompiledPlanRecord } from '../runtime-state.js'

export interface EngineCompiledPlanService {
  readonly intern: (source: string, compiled: CompiledFormula) => CompiledPlanRecord
  readonly replace: (planId: number, source: string, compiled: CompiledFormula) => CompiledPlanRecord
  readonly get: (planId: number) => CompiledPlanRecord | undefined
  readonly release: (planId: number) => void
}

interface MutableCompiledPlanRecord extends Omit<CompiledPlanRecord, 'source' | 'compiled'> {
  source: string
  compiled: CompiledFormula
  refCount: number
}

export function createEngineCompiledPlanService(): EngineCompiledPlanService {
  const records = new Map<number, MutableCompiledPlanRecord>()
  const planByCompiled = new WeakMap<CompiledFormula, MutableCompiledPlanRecord>()
  let nextPlanId = 1
  const releasePlanReference = (record: MutableCompiledPlanRecord): void => {
    record.refCount -= 1
    if (record.refCount > 0) {
      return
    }
    records.delete(record.id)
    planByCompiled.delete(record.compiled)
  }

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
    replace(planId, source, compiled) {
      const existing = records.get(planId)
      if (!existing) {
        return this.intern(source, compiled)
      }
      if (existing.compiled === compiled) {
        existing.source = source
        return existing
      }
      const alreadyInterned = planByCompiled.get(compiled)
      if (alreadyInterned && alreadyInterned !== existing) {
        releasePlanReference(existing)
        alreadyInterned.refCount += 1
        return alreadyInterned
      }
      if (existing.refCount === 1) {
        planByCompiled.delete(existing.compiled)
        existing.source = source
        existing.compiled = compiled
        planByCompiled.set(compiled, existing)
        return existing
      }
      releasePlanReference(existing)
      return this.intern(source, compiled)
    },
    get(planId) {
      return records.get(planId)
    },
    release(planId) {
      const existing = records.get(planId)
      if (!existing) {
        return
      }
      releasePlanReference(existing)
    },
  }
}
