import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { evaluateAst } from "../js-evaluator.js";
import { parseFormula } from "../parser.js";
import {
  renameFormulaSheetReferences,
  rewriteFormulaForStructuralTransform,
  serializeFormula,
  translateFormulaReferences,
} from "../translation.js";
import { evaluationContext } from "./formula-fuzz-helpers.js";
import { loadFormulaReplayFixtures } from "./formula-fuzz-replay-fixtures.js";

function expectTranslationFixture(
  canonical: string,
  translation: NonNullable<ReturnType<typeof loadFormulaReplayFixtures>[number]["translation"]>,
): void {
  const translated = translateFormulaReferences(
    canonical,
    translation.rowDelta,
    translation.colDelta,
  );
  expect(translated).toBe(translation.translated);
  expect(translateFormulaReferences(translated, -translation.rowDelta, -translation.colDelta)).toBe(
    translation.restored,
  );
}

function expectRenameFixture(
  canonical: string,
  rename: NonNullable<ReturnType<typeof loadFormulaReplayFixtures>[number]["rename"]>,
): void {
  const renamed = renameFormulaSheetReferences(canonical, rename.oldSheetName, rename.newSheetName);
  expect(renamed).toBe(rename.renamed);
  expect(renameFormulaSheetReferences(renamed, rename.newSheetName, rename.oldSheetName)).toBe(
    rename.restored,
  );
}

function expectStructuralFixture(
  canonical: string,
  structural: NonNullable<ReturnType<typeof loadFormulaReplayFixtures>[number]["structural"]>,
): void {
  const rewritten = rewriteFormulaForStructuralTransform(
    canonical,
    structural.ownerSheetName,
    structural.targetSheetName,
    structural.transform,
  );
  expect(rewritten).toBe(structural.rewritten);
  expect(serializeFormula(parseFormula(rewritten))).toBe(structural.rewritten);
  if (!structural.reversed) {
    return;
  }
  expect(
    rewriteFormulaForStructuralTransform(
      rewritten,
      structural.ownerSheetName,
      structural.targetSheetName,
      {
        kind: structural.reversed.kind,
        axis: structural.reversed.axis,
        start: structural.reversed.start,
        count: structural.reversed.count,
      },
    ),
  ).toBe(structural.reversed.restored);
}

function expectEvaluationFixture(
  canonical: string,
  evaluation: NonNullable<ReturnType<typeof loadFormulaReplayFixtures>[number]["evaluation"]>,
): void {
  const actual = evaluateAst(parseFormula(canonical), evaluationContext);
  const expected =
    evaluation.expected.kind === "number"
      ? { tag: ValueTag.Number, value: evaluation.expected.value }
      : evaluation.expected.kind === "string"
        ? { tag: ValueTag.String, value: evaluation.expected.value }
        : { tag: ValueTag.Boolean, value: evaluation.expected.value };
  expect(actual).toMatchObject(expected);
}

describe("formula replay fixtures", () => {
  for (const fixture of loadFormulaReplayFixtures()) {
    it(`replays ${fixture.name}`, () => {
      const canonical = serializeFormula(parseFormula(fixture.source));
      expect(canonical).toBe(fixture.canonical);

      if (fixture.translation) {
        expectTranslationFixture(canonical, fixture.translation);
      }

      if (fixture.rename) {
        expectRenameFixture(canonical, fixture.rename);
      }

      if (fixture.structural) {
        expectStructuralFixture(canonical, fixture.structural);
      }

      if (fixture.evaluation) {
        expectEvaluationFixture(canonical, fixture.evaluation);
      }
    });
  }
});
