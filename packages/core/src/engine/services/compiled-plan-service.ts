import type { CompiledFormula } from "@bilig/formula";
import type { CompiledPlanRecord } from "../runtime-state.js";

export interface EngineCompiledPlanService {
  readonly intern: (source: string, compiled: CompiledFormula) => CompiledPlanRecord;
  readonly get: (planId: number) => CompiledPlanRecord | undefined;
  readonly release: (planId: number) => void;
}

interface MutableCompiledPlanRecord extends CompiledPlanRecord {
  refCount: number;
  signature: string;
}

function compiledPlanSignature(source: string, compiled: CompiledFormula): string {
  return JSON.stringify({
    source,
    mode: compiled.mode,
    deps: compiled.deps,
    symbolicNames: compiled.symbolicNames,
    symbolicTables: compiled.symbolicTables,
    symbolicSpills: compiled.symbolicSpills,
    symbolicRefs: compiled.symbolicRefs,
    symbolicRanges: compiled.symbolicRanges,
    symbolicStrings: compiled.symbolicStrings,
    program: [...compiled.program],
    constants: [...compiled.constants],
    volatile: compiled.volatile,
    randCallCount: compiled.randCallCount,
    producesSpill: compiled.producesSpill,
    jsPlan: compiled.jsPlan,
  });
}

export function createEngineCompiledPlanService(): EngineCompiledPlanService {
  const records = new Map<number, MutableCompiledPlanRecord>();
  const planIdBySignature = new Map<string, number>();
  let nextPlanId = 1;

  return {
    intern(source, compiled) {
      const signature = compiledPlanSignature(source, compiled);
      const existingId = planIdBySignature.get(signature);
      if (existingId !== undefined) {
        const existing = records.get(existingId);
        if (!existing) {
          throw new Error(`Missing compiled plan record for id ${existingId}`);
        }
        existing.refCount += 1;
        return existing;
      }
      const record: MutableCompiledPlanRecord = {
        id: nextPlanId,
        source,
        compiled,
        refCount: 1,
        signature,
      };
      nextPlanId += 1;
      records.set(record.id, record);
      planIdBySignature.set(signature, record.id);
      return record;
    },
    get(planId) {
      return records.get(planId);
    },
    release(planId) {
      const existing = records.get(planId);
      if (!existing) {
        return;
      }
      existing.refCount -= 1;
      if (existing.refCount > 0) {
        return;
      }
      records.delete(planId);
      planIdBySignature.delete(existing.signature);
    },
  };
}
