import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  extractClassSurface,
  extractInterfaceKeys,
  parseHyperFormulaSurfaceSnapshot,
} from "../../../../scripts/workpaper-surface-contract.js";

const ALLOWED_BILIG_INSTANCE_METHODS = [
  "dispose",
  "offDetailed",
  "onDetailed",
  "onceDetailed",
] as const;

describe("WorkPaper HyperFormula snapshot parity", () => {
  it("matches the checked-in HyperFormula class surface snapshot", () => {
    const snapshot = loadSnapshot();
    const currentSurface = extractClassSurface(
      readFileSync(new URL("../headless-workbook.ts", import.meta.url), "utf8"),
      "HeadlessWorkbook",
    );

    expect(currentSurface.staticMembers).toEqual(snapshot.classSurface.staticMembers);
    expect(currentSurface.staticMethods).toEqual(snapshot.classSurface.staticMethods);
    expect(currentSurface.instanceAccessors).toEqual(snapshot.classSurface.instanceAccessors);
    expect(currentSurface.instanceMethods).toEqual(
      [...snapshot.classSurface.instanceMethods, ...ALLOWED_BILIG_INSTANCE_METHODS].toSorted(),
    );
  });

  it("matches the checked-in HyperFormula config-key snapshot", () => {
    const snapshot = loadSnapshot();
    const currentConfigKeys = extractInterfaceKeys(
      readFileSync(new URL("../types.ts", import.meta.url), "utf8"),
      "HeadlessConfig",
    );

    expect(currentConfigKeys).toEqual(snapshot.configKeys);
  });
});

function loadSnapshot() {
  return parseHyperFormulaSurfaceSnapshot(
    readFileSync(new URL("./fixtures/hyperformula-surface.json", import.meta.url), "utf8"),
  );
}
