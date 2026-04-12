import { describe, expect, it } from "vitest";
import { compileFormula, compileFormulaAst, parseFormula } from "@bilig/formula";
import { createEngineCompiledPlanService } from "../engine/services/compiled-plan-service.js";

describe("EngineCompiledPlanService", () => {
  it("reuses one immutable compiled plan for identical compiled formulas and releases it by refcount", () => {
    const service = createEngineCompiledPlanService();

    const first = service.intern("1+2", compileFormula("1+2"));
    const second = service.intern("1+2", compileFormula("1+2"));
    const different = service.intern("2+3", compileFormula("2+3"));

    expect(second.id).toBe(first.id);
    expect(second.compiled).toBe(first.compiled);
    expect(different.id).not.toBe(first.id);

    service.release(first.id);
    expect(service.get(first.id)).toBeDefined();

    service.release(second.id);
    expect(service.get(first.id)).toBeUndefined();
    expect(service.get(different.id)).toBeDefined();
  });

  it("does not merge distinct compiled plans that happen to share a source string", () => {
    const service = createEngineCompiledPlanService();

    const first = service.intern("Shared", compileFormulaAst("Shared", parseFormula("1+2")));
    const second = service.intern("Shared", compileFormulaAst("Shared", parseFormula("2+3")));

    expect(second.id).not.toBe(first.id);
    expect([...second.compiled.constants]).not.toEqual([...first.compiled.constants]);
  });
});
