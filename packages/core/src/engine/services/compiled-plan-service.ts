import type { CompiledFormula } from '@bilig/formula'
import type { CompiledPlanRecord } from '../runtime-state.js'

export interface EngineCompiledPlanService {
  readonly intern: (source: string, compiled: CompiledFormula, templateId?: number) => CompiledPlanRecord
  readonly replace: (planId: number, source: string, compiled: CompiledFormula, templateId?: number) => CompiledPlanRecord
  readonly get: (planId: number) => CompiledPlanRecord | undefined
  readonly release: (planId: number) => void
  readonly clear: () => void
}

interface MutableCompiledPlanRecord extends Omit<CompiledPlanRecord, 'source' | 'compiled' | 'templateId'> {
  source: string
  compiled: CompiledFormula
  templateId?: number
  refCount: number
}

export function createEngineCompiledPlanService(): EngineCompiledPlanService {
  const records = new Map<number, MutableCompiledPlanRecord>()
  let planByCompiled = new WeakMap<CompiledFormula, MutableCompiledPlanRecord>()
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
    intern(source, compiled, templateId) {
      const existing = planByCompiled.get(compiled)
      if (existing !== undefined) {
        existing.refCount += 1
        if (templateId !== undefined) {
          existing.templateId = templateId
        } else {
          delete existing.templateId
        }
        return existing
      }
      const record: MutableCompiledPlanRecord = {
        id: nextPlanId,
        source,
        compiled,
        ...(templateId !== undefined ? { templateId } : {}),
        refCount: 1,
      }
      nextPlanId += 1
      records.set(record.id, record)
      planByCompiled.set(compiled, record)
      return record
    },
    replace(planId, source, compiled, templateId) {
      const existing = records.get(planId)
      if (!existing) {
        return this.intern(source, compiled, templateId)
      }
      if (existing.compiled === compiled) {
        existing.source = source
        if (templateId !== undefined) {
          existing.templateId = templateId
        } else {
          delete existing.templateId
        }
        return existing
      }
      const alreadyInterned = planByCompiled.get(compiled)
      if (alreadyInterned && alreadyInterned !== existing) {
        releasePlanReference(existing)
        alreadyInterned.refCount += 1
        if (templateId !== undefined) {
          alreadyInterned.templateId = templateId
        } else {
          delete alreadyInterned.templateId
        }
        return alreadyInterned
      }
      if (existing.refCount === 1) {
        planByCompiled.delete(existing.compiled)
        existing.source = source
        existing.compiled = compiled
        if (templateId !== undefined) {
          existing.templateId = templateId
        } else {
          delete existing.templateId
        }
        planByCompiled.set(compiled, existing)
        return existing
      }
      releasePlanReference(existing)
      return this.intern(source, compiled, templateId)
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
    clear() {
      records.clear()
      planByCompiled = new WeakMap<CompiledFormula, MutableCompiledPlanRecord>()
      nextPlanId = 1
    },
  }
}
